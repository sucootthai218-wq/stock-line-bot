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

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

DRIVE_FOLDER_ID = "19DLipG-4_C0qWTOsFGXyJWfhLsNvR4V8"
GSHEET_URL = "https://docs.google.com/spreadsheets/d/1ngb3u6xzE6m0QSre1gTwFxm0_hElAavyPGKkOHY98Vc/edit?gid=0#gid=0"

# การตั้งค่าตำแหน่งข้อมูล (ตรวจสอบให้แน่ใจว่า Balance อยู่ที่คอลัมน์ 63 หรือไม่ ถ้าไม่ใช่ให้ปรับเลขนี้ครับ)
HEADER_ROW = 4
CODE_B_INDEX = 1
CODE_C_INDEX = 2
DESC_COL_INDEX = 3     
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
    service = build('drive', 'v3', credentials=creds)
    query = f"'{folder_id}' in parents and (mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType='application/vnd.ms-excel') and trashed=false"
    results = service.files().list(q=query, orderBy="modifiedTime desc", pageSize=count, fields="files(id, name)").execute()
    return [item['id'] for item in results.get('files', [])]

def process_order_and_get_summary(user_msg):
    creds = get_google_credentials()
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_url(GSHEET_URL)
    ws_input = spreadsheet.worksheet("Input_Order")
    
    # ดึงข้อมูลจาก LINE
    lines = user_msg.strip().split('\n')
    input_data = [["Code", "Qty"]]
    for line in lines:
        parts = line.strip().split()
        if len(parts) >= 2: input_data.append([parts[0], parts[1]])
        elif len(parts) == 1: input_data.append([parts[0], "1"])
    ws_input.clear()
    ws_input.update(input_data, 'A1')

    # ล้างไฟล์เก่า
    for f in glob.glob("*.xlsx"):
        try: os.remove(f)
        except: pass

    file_ids = get_latest_file_ids_from_folder(creds, DRIVE_FOLDER_ID, count=2)
    if not file_ids: return "❌ ไม่พบไฟล์ในโฟลเดอร์ My Stock"

    for fid in file_ids:
        gdown.download(id=fid, quiet=False)

    global_code_map = {}
    for file_path in glob.glob("*.xlsx"):
        try:
            wb = load_workbook(filename=file_path, read_only=True, data_only=True)
            sheet = wb.active
            for r_idx, row in enumerate(sheet.iter_rows(values_only=True)):
                if r_idx <= HEADER_ROW or not row: continue
                # ค้นหาโค้ดจากคอลัมน์ที่ระบุ
                for col_idx in [CODE_B_INDEX, CODE_C_INDEX]:
                    if col_idx < len(row) and row[col_idx]:
                        norm_c = normalize_code(row[col_idx])
                        if norm_c and len(norm_c) > 3:
                            global_code_map[norm_c] = {
                                'code': str(row[col_idx]),
                                'desc': str(row[DESC_COL_INDEX] if DESC_COL_INDEX < len(row) else ""),
                                'balance': clean_num(row[BALANCE_COL_INDEX] if BALANCE_COL_INDEX < len(row) else 0)
                            }
            wb.close()
        except Exception as e:
            return f"❌ อ่านไฟล์ผิดพลาด: {str(e)}"

    # คำนวณ
    df_input = pd.DataFrame(ws_input.get_all_records())
    report_items = []
    for _, row in df_input.iterrows():
        norm_input = normalize_code(str(row[df_input.columns[0]]))
        qty_needed = clean_num(row[df_input.columns[1]])
        data = global_code_map.get(norm_input)
        if data:
            bal = data['balance']
            shortage = qty_needed if bal < 0 else max(0, qty_needed - bal)
            report_items.append({'code': data['code'], 'desc': data['desc'], 'shortage': "มีของ" if shortage == 0 else int(shortage), 'balance': int(bal)})
        else:
            report_items.append({'code': str(row[df_input.columns[0]]), 'desc': "ไม่พบรหัส", 'shortage': int(qty_needed), 'balance': "-"})

    # สรุปผล
    summary_text = "📊 รายงานสรุปสต็อก:\n" + "".join([f"- {i['code']} ({i['desc']}): คงเหลือ {i['balance']} | ขาด {i['shortage']}\n" for i in report_items])
    ws_summary = spreadsheet.worksheet("Summary")
    ws_summary.clear()
    ws_summary.update([["รหัสสินค้า", "รายการ", "สต็อก Balance", "ขาด"]] + [[i['code'], i['desc'], i['balance'], i['shortage']] for i in report_items], 'A1')
    return summary_text

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try: handler.handle(body, signature)
    except: abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    try: reply = process_order_and_get_summary(event.message.text.strip())
    except Exception as e: reply = f"❌ Error: {str(e)}"
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
