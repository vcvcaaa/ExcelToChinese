import os
import google.generativeai as genai # <-- 修改: 匯入新的函式庫
from flask import Flask, render_template, request, send_from_directory, jsonify, abort, after_this_request
from openpyxl import load_workbook
import uuid
import time # <-- 新增: 用於處理 API 速率限制

# --- 設定區 ---
# 建立 Flask 應用
app = Flask(__name__)

# 設定一個簡單的密碼
AUTH_PASSWORD = '123'
UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

# --- 修改: 設定您的 Gemini API 金鑰 ---
# 最佳實踐是將 API 金鑰儲存在環境變數中
# 執行前請在終端機設定:
# Linux/macOS: export GEMINI_API_KEY="YOUR_GEMINI_API_KEY"
# Windows: set GEMINI_API_KEY="YOUR_GEMINI_API_KEY"
try:
    gemini_api_key = os.getenv("GEMINI_API_KEY")
    if not gemini_api_key:
        raise ValueError("錯誤：請設定 GEMINI_API_KEY 環境變數")
    genai.configure(api_key=gemini_api_key)
    # 使用一個速度快且穩定的模型
    model = genai.GenerativeModel('gemini-1.5-flash-latest') 
except Exception as e:
    print(e)
    model = None

# 確保暫存資料夾存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

# --- 核心功能 (已修改為使用 Gemini) ---

def translate_text_with_gemini(text_to_translate):
    """使用 Gemini API 翻譯文字"""
    # 如果文字為空、不是字串或模型未成功初始化，則直接返回原文
    if not text_to_translate or not isinstance(text_to_translate, str) or not model:
        return text_to_translate

    # --- 提示工程 (Prompt Engineering) ---
    # 這是引導 Gemini 做出正確翻譯的關鍵
    # 我們明確指示它：1. 任務目標 2. 來源語言 3. 目標語言 4. 輸出格式要求
    prompt = f"""
    請將以下的越南文文字翻譯成繁體中文。

    規則：
    1.  直接提供翻譯結果，不要有任何額外的解釋、前言或附註。
    2.  保留原始文字中的換行符號 (如有)。
    3.  如果內容是數字、英文、或看起來不是越南文，請直接返回原文。

    越南文原文: "{text_to_translate}"
    繁體中文翻譯:
    """
    
    # 加入重試機制和速率限制處理
    for i in range(3): # 最多重試 3 次
        try:
            response = model.generate_content(prompt)
            # 確保 response.text 存在且不為空
            if response.text:
                return response.text.strip()
            else:
                # 如果 Gemini 回傳空內容，則返回原文以避免儲存格變空
                print(f"警告: Gemini 為 '{text_to_translate}' 回傳了空的結果。")
                return text_to_translate
                
        except Exception as e:
            # 處理 API 的速率限制錯誤 (ResourceExhausted)
            if '429' in str(e): 
                print(f"觸發 API 速率限制，等待 {2**i} 秒後重試...")
                time.sleep(2**i) # 指數退避等待
            else:
                print(f"呼叫 Gemini API 時發生錯誤: {e}")
                # 其他錯誤發生時，直接返回原文
                return text_to_translate
    
    # 如果重試三次都失敗，返回原文
    print(f"警告: 重試多次後，翻譯 '{text_to_translate}' 仍然失敗。")
    return text_to_translate


def process_excel_file(input_path, output_path):
    """讀取 Excel 檔案，翻譯越南文儲存格並儲存"""
    workbook = load_workbook(input_path)
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        print(f"--- 正在處理工作表: {sheet_name} ---")
        for row_idx, row in enumerate(sheet.iter_rows(), 1):
            for col_idx, cell in enumerate(row, 1):
                if cell.value and isinstance(cell.value, str) and cell.value.strip():
                    original_text = cell.value
                    # Gemini 會自行判斷語言，所以我們直接送出翻譯
                    translated_text = translate_text_with_gemini(original_text)
                    
                    # 只有在翻譯結果與原文不同時才印出，方便觀察
                    if original_text != translated_text:
                        print(f"單元格({row_idx}, {col_idx}) | 原文: '{original_text}' -> 譯文: '{translated_text}'")
                    
                    cell.value = translated_text
    workbook.save(output_path)


# --- 路由 (Web Routes) ---
# 以下的路由部分與前一版本完全相同，無需修改

@app.route('/')
def index():
    """渲染主頁面"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """處理檔案上傳與翻譯"""
    password = request.form.get('password')
    if password != AUTH_PASSWORD:
        return jsonify({'error': '密碼錯誤！'}), 401

    if 'file' not in request.files:
        return jsonify({'error': '沒有檔案被上傳。'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '沒有選擇檔案。'}), 400

    if file and file.filename.endswith('.xlsx'):
        unique_id = str(uuid.uuid4())
        original_filename = f"{unique_id}_original.xlsx"
        translated_filename = f"{unique_id}_translated.xlsx"
        
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], original_filename)
        output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], translated_filename)
        
        file.save(input_path)
        
        try:
            # 檢查模型是否成功初始化
            if not model:
                 return jsonify({'error': '伺服器端 Gemini 模型未成功初始化，請檢查 API 金鑰設定。'}), 500

            process_excel_file(input_path, output_path)
            os.remove(input_path)
            
            return jsonify({
                'success': True,
                'download_url': f'/download/{translated_filename}'
            })
        except Exception as e:
            print(f"處理檔案時發生錯誤: {e}")
            return jsonify({'error': f'處理檔案時發生錯誤: {e}'}), 500
    else:
        return jsonify({'error': '只支援 .xlsx 格式的 Excel 檔案。'}), 400

@app.route('/download/<filename>')
def download_file(filename):
    """提供翻譯後檔案的下載"""
    file_path = os.path.join(app.config['DOWNLOAD_FOLDER'], filename)
    if not os.path.exists(file_path):
        abort(404, "找不到檔案，可能已被刪除或處理失敗。")
        
    @after_this_request
    def cleanup(response):
        try:
            os.remove(file_path)
            print(f"已刪除檔案: {filename}")
        except OSError as e:
            print(f"刪除檔案時出錯: {e.strerror}")
        return response
        
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True)

# --- 啟動伺服器 ---
if __name__ == '__main__':
    if not model:
        print("致命錯誤：Gemini 模型未被初始化。請檢查您的 GEMINI_API_KEY 環境變數。")
    else:
        print("Gemini 模型已成功載入。伺服器準備就緒。")
        app.run(host='0.0.0.0', port=5000, debug=True)