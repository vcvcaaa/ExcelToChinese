import os
import google.generativeai as genai
from flask import Flask, render_template, request, send_from_directory, jsonify, abort, after_this_request
from openpyxl import load_workbook
import uuid
import time
import math
import json
import sys

# --- 設定區 ---
# 建立 Flask 應用
app = Flask(__name__)

# 設定一個簡單的密碼
AUTH_PASSWORD = '123'
UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
GLOSSARY_FILE = 'dic.json'  # 指定您的詞彙表檔案名稱
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

# --- 【新功能】載入專業詞彙表 ---
def load_glossary(filepath):
    """從 JSON 檔案載入專業詞彙表，並建立一個 越南文 -> 繁體中文 的字典。"""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        # 建立字典，並將越南文關鍵字轉為小寫，以便進行不分大小寫的比對
        # {'giá trị ph': 'pH值', 'giao phối': '交配'}
        glossary = {item['vietnamese'].lower(): item['chinese'] for item in data}
        print(f"✅ 成功載入專業詞彙表，共 {len(glossary)} 條目。")
        return glossary
    except FileNotFoundError:
        print(f"❌ 錯誤：找不到詞彙表檔案 '{filepath}'。請確保檔案存在於正確的位置。")
        return None
    except json.JSONDecodeError:
        print(f"❌ 錯誤：詞彙表檔案 '{filepath}' 格式不正確，無法解析。")
        return None
    except Exception as e:
        print(f"❌ 載入詞彙表時發生未知錯誤: {e}")
        return None

# 在程式啟動時，載入一次詞彙表
VIET_TO_CHI_GLOSSARY = load_glossary(GLOSSARY_FILE)

# --- 設定您的 Gemini API 金鑰 ---
try:
    gemini_api_key = os.getenv("GEMINI_API_KEY")
    if not gemini_api_key:
        raise ValueError("錯誤：請設定 GEMINI_API_KEY 環境變數")
    genai.configure(api_key=gemini_api_key)
    model = genai.GenerativeModel('gemini-1.5-flash-latest')
except Exception as e:
    print(f"初始化 Gemini 模型時發生錯誤: {e}")
    model = None

# 確保暫存資料夾存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

# --- 核心功能 (已整合智慧提示詞注入) ---

def translate_text_batch_with_gemini(texts, separator):
    """
    使用 Gemini API 批次翻譯一個文字塊(chunk)，並智慧注入相關的專業詞彙。
    """
    if not texts or not model:
        return texts

    combined_text = separator.join(texts)
    
    # --- 【修改點】智慧注入邏輯 ---
    # 檢查此批次的原文中，是否包含詞彙表中的任何越南文術語
    relevant_terms_list = []
    if VIET_TO_CHI_GLOSSARY:
        source_text_lower = combined_text.lower()
        for viet_term, chi_translation in VIET_TO_CHI_GLOSSARY.items():
            # viet_term 已經是小寫
            if viet_term in source_text_lower:
                relevant_terms_list.append(f"- 越南原文 '{viet_term}' 應翻譯為 '{chi_translation}'")
    
    # 建立注入到提示詞中的詞彙表部分
    if relevant_terms_list:
        glossary_section = (
            "**【優先翻譯詞彙】**\n"
            "在翻譯時，請務必遵循以下術語對照，優先使用指定的翻譯：\n"
            + "\n".join(relevant_terms_list)
        )
    else:
        glossary_section = "**【優先翻譯詞彙】**\n無特定術語需要優先處理。"
    # --- 智慧注入邏輯結束 ---

    prompt = f"""
    任務：請將以下由特殊分隔符 `"{separator}"` 串連起來的越南文文字，逐一精準地翻譯成繁體中文。

    規則：
    1.  翻譯完成後，必須使用完全相同的分隔符 `"{separator}"` 將所有翻譯結果重新串連起來。
    2.  除了翻譯後的文字和分隔符外，不要包含任何前言、解釋或額外字元。
    3.  如果某個片段是數字、英文、純粹的標點符號或看起來不是越南文，請直接返回該片段原文。
    4.  保持片段的順序與原文完全一致。
    5.  返回的片段數量必須與原始片段數量完全相同。

    {glossary_section}

    ---
    待翻譯的原文組合:
    "{combined_text}"

    ---
    翻譯後的繁體中文組合:
    """

    for i in range(3): # 重試機制
        try:
            response = model.generate_content(prompt)
            if response.text:
                translated_texts = response.text.strip().split(separator)
                if len(translated_texts) == len(texts):
                    return translated_texts # 成功，返回結果
                else:
                    print(f"批次錯誤：API回傳的片段數量 ({len(translated_texts)}) 與原始數量 ({len(texts)}) 不符。原文: {combined_text[:100]}...")
                    # 數量不符時，不再重試，直接返回原文以保證安全
                    return texts
            else:
                print(f"警告: Gemini 為批次任務回傳了空的結果。")
                # 返回原文以保證安全
                return texts
        except Exception as e:
            if '429' in str(e): # 速率限制錯誤
                wait_time = 2**i
                print(f"觸發 API 速率限制，等待 {wait_time} 秒後重試...")
                time.sleep(wait_time)
            else:
                print(f"呼叫 Gemini API 時發生錯誤: {e}")
                return texts # 其他錯誤，返回原文

    print("警告: 重試多次後，批次翻譯仍然失敗。將使用原文填充。")
    return texts

def process_excel_file_optimized(input_path, output_path):
    """
    【已升級為雙語模式】讀取 Excel，使用分塊 + 批次處理進行翻譯，並將原文與譯文合併儲存。
    """
    workbook = load_workbook(input_path)
    SEPARATOR = "|||$$$|||"
    CHUNK_SIZE = 150 # 每個批次處理的儲存格數量

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        print(f"--- 正在掃描工作表: {sheet_name} ---")

        texts_to_translate = []
        cell_locations = []

        for row_idx, row in enumerate(sheet.iter_rows(min_row=1), 1):
            for col_idx, cell in enumerate(row, 1):
                if cell.value and isinstance(cell.value, str) and cell.value.strip():
                    texts_to_translate.append(cell.value)
                    cell_locations.append((row_idx, col_idx))

        if not texts_to_translate:
            print(f"工作表 '{sheet_name}' 中沒有需要翻譯的文字。")
            continue

        print(f"在 '{sheet_name}' 中找到 {len(texts_to_translate)} 個待翻譯的儲存格。開始分塊處理...")

        all_translated_texts = []
        total_chunks = math.ceil(len(texts_to_translate) / CHUNK_SIZE)

        for i in range(0, len(texts_to_translate), CHUNK_SIZE):
            chunk = texts_to_translate[i:i + CHUNK_SIZE]
            current_chunk_num = (i // CHUNK_SIZE) + 1
            print(f"正在處理第 {current_chunk_num} / {total_chunks} 批... (本批次共 {len(chunk)} 個儲存格)")
            
            translated_chunk = translate_text_batch_with_gemini(chunk, SEPARATOR)
            
            if len(translated_chunk) != len(chunk):
                print(f"警告：第 {current_chunk_num} 批翻譯失敗，為保護資料，此批次將使用原文。")
                all_translated_texts.extend(chunk) # 使用原文
            else:
                all_translated_texts.extend(translated_chunk)
        
        print("所有批次翻譯完成，正在將原文與譯文合併寫回 Excel...")
        if len(all_translated_texts) == len(cell_locations):
            for i, location in enumerate(cell_locations):
                row, col = location
                original_text = texts_to_translate[i]
                translated_text = all_translated_texts[i].strip()

                # 只有在翻譯結果與原文不同時，才合併文字
                if original_text.strip() != translated_text and translated_text:
                    # 新格式：原文 + 換行 + 譯文
                    new_value = f"{original_text}\n{translated_text}"
                    sheet.cell(row=row, column=col).value = new_value
                else:
                    # 如果沒有翻譯（例如是數字或英文），則保持原文不變
                    sheet.cell(row=row, column=col).value = original_text
        else:
             print(f"【嚴重錯誤】：最終文本數量與位置數量不符，為防止數據損壞，工作表 '{sheet_name}' 不進行任何修改。")

    print("--- 所有工作表處理完畢，正在儲存檔案... ---")
    workbook.save(output_path)

# --- 路由 (Web Routes) ---

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    password = request.form.get('password')
    if password != AUTH_PASSWORD:
        return jsonify({'success': False, 'error': '密碼錯誤！'}), 401

    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '沒有檔案被上傳。'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': '沒有選擇檔案。'}), 400

    if file and file.filename.endswith('.xlsx'):
        unique_id = str(uuid.uuid4())
        original_filename = f"{unique_id}_original.xlsx"
        translated_filename = f"{unique_id}_translated.xlsx"

        input_path = os.path.join(app.config['UPLOAD_FOLDER'], original_filename)
        output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], translated_filename)

        file.save(input_path)

        try:
            if not model:
                 return jsonify({'success': False, 'error': '伺服器端 Gemini 模型未成功初始化，請檢查 API 金鑰設定。'}), 500

            process_excel_file_optimized(input_path, output_path)
            
            # 處理完畢後，立即刪除上傳的原始檔案
            os.remove(input_path)

            return jsonify({
                'success': True,
                'download_url': f'/download/{translated_filename}'
            })
        except Exception as e:
            print(f"處理檔案時發生嚴重錯誤: {e}")
            # 如果發生錯誤，也嘗試刪除可能已儲存的原始檔案
            if os.path.exists(input_path):
                os.remove(input_path)
            return jsonify({'success': False, 'error': f'處理檔案時發生錯誤: {e}'}), 500
    else:
        return jsonify({'success': False, 'error': '只支援 .xlsx 格式的 Excel 檔案。'}), 400

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(app.config['DOWNLOAD_FOLDER'], filename)
    if not os.path.exists(file_path):
        abort(404, "找不到檔案，可能已被刪除或處理失敗。")

    @after_this_request
    def cleanup(response):
        """下載請求結束後，刪除伺服器上的已翻譯檔案"""
        try:
            os.remove(file_path)
            print(f"已刪除已下載的檔案: {filename}")
        except OSError as e:
            print(f"刪除檔案時出錯: {e.strerror}")
        return response

    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True)

# --- 啟動伺服器 ---
if __name__ == '__main__':
    # 啟動前進行最終檢查
    if not model:
        print("致命錯誤：Gemini 模型未被初始化。請檢查您的 GEMINI_API_KEY 環境變數後再重新啟動。")
        sys.exit(1) # 退出程式
    
    if not VIET_TO_CHI_GLOSSARY:
        print("致命錯誤：專業詞彙表未能成功載入。請檢查 'dic.json' 檔案後再重新啟動。")
        sys.exit(1) # 退出程式

    print("Gemini 模型及專業詞彙表已成功載入。伺服器準備就緒。")
    app.run(host='0.0.0.0', port=5000, debug=False)