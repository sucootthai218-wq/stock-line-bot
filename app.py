import os
import glob
import re
import json
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
import gspread
import pandas as pd
import gdown
from openpyxl import load_workbook

app = Flask(__name__)

LINE_CHANNEL_ACCESS_TOKEN = os.environ.get('LINE_CHANNEL_ACCESS_TOKEN')
LINE_CHANNEL_SECRET = os.environ.get('LINE_CHANNEL_SECRET')

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

DRIVE_FOLDER_ID = "19DLipG-4_C0qWTOsFGXyJWfhLsNvR4V8"
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1ngb3u6xzE6m0QSre1gTwFxm0_hElAavyPGKkOHY98Vc/edit?gid=0#gid=0"

HEADER_ROW = 4
CODE_B_INDEX = 1
CODE_C_INDEX = 2
DESC_COL_INDEX = 3     
NEW_COL_INDEX = 12
OLD_COL_INDEX = 13
BALANCE_COL_INDEX = 63

def clean_num(val):
    if pd.isna(val) or val is None: return 0.0
    s = str(val).replace(',', '').strip()
    if s in ["-", "_", "", "nan", "None"]: return 0.0
    try: return float(s)
    except: return 0.0

def normalize_code(code_str):
    if not code_str: return ""
    return re.sub(r'[^A-Z0-9]', '', str(code_str).strip().upper())

def get_google_credentials():
    google_creds_json = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if google_creds_json:
        creds_dict = json.loads(google_creds_json)
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        return Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    else:
        return Credentials.from_service_account_file("credentials.json", scopes=SCOPES)

def get_latest_file_ids_from_folder(creds, folder_id, count=2):
    """ ดึงเฉพาะ File ID ของไฟล์ Excel 2 ไฟล์ล่าสุดจากโฟลเดอร์ """
    service = build('drive', 'v3', credentials=creds)
    query = f"'{folder_id}' in parents and (mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType='application/vnd.ms-excel') and trashed=false"
    results = service.files().list(q=query, orderBy="modifiedTime desc", pageSize=count, fields="files(id, name)").execute()
    items = results.get('files', [])
    return [item['id'] for item in items]

def process_order_and_get_summary(user_msg):
    creds = get_google_credentials()
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_url(GSHEET_URL)

    ws_input = spreadsheet.worksheet("Input_Order")
    ws_input.clear()
    
    lines = user_msg.strip().split('\n')
    input_data = [["Code", "Qty"]]
    for line in lines:
        parts = line.strip().split()
        if len(parts) >= 2:
            input_data.append([parts[0], parts[1]])
        elif len(parts) == 1:
            input_data.append([parts[0], "1"])
    ws_input.update(input_data, 'A1')

    # ลบไฟล์เก่าทิ้ง
    for f in glob.glob("*.xlsx"):
        try: os.remove(f)
        except: pass

    # ดึง 2 ไฟล์ล่าสุดจากโฟลเดอร์ My Stock
    file_ids = get_latest_file_ids_from_folder(creds, DRIVE_FOLDER_ID, count=2)
    for fid in file_ids:
        gdown.download(id=fid, quiet=False)

    all_xlsx = [f for f in glob.glob("*.xlsx") if not os.path.basename(f).startswith("~$")]
    global_code_map = {}

    # อ่านไฟล์แบบ read_only=True เพื่อประหยัด RAM ไม่ให้ Render ตัดทำงาน
    for file_path in all_xlsx:
        try:
            wb = load_workbook(filename=file_path, read_only=True, data_only=True)
            sheet = wb.active
            rows = list(sheet.iter_rows(values_only=True))
            wb.close()
            
            for r_idx in range(HEADER_ROW + 1, len(rows)):
                row_data = rows[r_idx]
                if not row_data:
                    continue
                
                for col_idx in [CODE_B_INDEX, CODE_C_INDEX]:
                    if col_idx < len(row_data):
                        c_val = row_data[col_idx]
                        if c_val is not None:
                            norm_c = normalize_code(c_val)
                            if norm_c and norm_c not in ["NAN", "NONE", "0", "CODE", "รหัสสินค้า", "รหัส", "NO"]:
                                if norm_c not in global_code_map:
                                    b_val = row_data[CODE_B_INDEX] if CODE_B_INDEX < len(row_data) else ""
                                    desc_val = row_data[DESC_COL_INDEX] if DESC_COL_INDEX < len(row_data) else ""
                                    bal_val = row_data[BALANCE_COL_INDEX] if BALANCE_COL_INDEX < len(row_data) else 0
                                    
                                    global_code_map[norm_c] = {
                                        'code': str(b_val or c_val),
                                        'desc': str(desc_val or ""),
                                        'balance': clean_num(bal_val)
                                    }
        except Exception as e:
            print(f"Error reading {file_path}: {e}")

    df_input = pd.DataFrame(ws_input.get_all_records())
    report_items = []
    
    for _, row in df_input.iterrows():
        raw_code = str(row[df_input.columns[0]]).strip()
        norm_input = normalize_code(raw_code)
        qty_needed = clean_num(row[df_input.columns[1]])
        
        match_data = global_code_map.get(norm_input)
        if match_data is not None:
            balance = match_data['balance']
            shortage = qty_needed if balance < 0 else max(0, qty_needed - balance)
            report_items.append({
                'code': match_data['code'], 
                'desc': match_data['desc'],
                'shortage': "มีของ" if shortage == 0 else int(shortage),
                'balance': int(balance)
            })
        else:
            report_items.append({'code': raw_code, 'desc': "ไม่พบรหัส", 'shortage': int(qty_needed), 'balance': "-"})

    summary_text = "📊 รายงานสรุปสต็อก:\n"
    for item in report_items:
        summary_text += f"- {item['code']} ({item['desc']}): คงเหลือ {item['balance']} | ขาด {item['shortage']}\n"

    header = ["รหัสสินค้า", "รายการ", "สต็อก Balance", "ขาด"]
    table_data = [header] + [[i['code'], i['desc'], i['balance'], i['shortage']] for i in report_items]
    ws_summary = spreadsheet.worksheet("Summary") if "Summary" in [w.title for w in spreadsheet.worksheets()] else spreadsheet.add_worksheet(title="Summary", rows="100", cols="30")
    ws_summary.clear()
    ws_summary.update(table_data, 'A1')
    return summary_text

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try: handler.handle(body, signature)
    except InvalidSignatureError: abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    try: reply_text = process_order_and_get_summary(event.message.text.strip())
    except Exception as e: reply_text = f"❌ เกิดข้อผิดพลาด: {str(e)}"
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply_text))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
