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

app = Flask(__name__)

# ตั้งค่าตำแหน่งคอลัมน์ (Excel Column Index เริ่มต้นที่ 0)
HEADER_ROW = 4
CODE_B_INDEX = 1
DESC_COL_INDEX = 3
NEW_COL_INDEX = 12
OLD_COL_INDEX = 13
BALANCE_COL_INDEX = 63
# ขยายขอบเขตการหาโครงการให้ไกลขึ้น (คอลัมน์ 14 ถึง 100)
PROJECT_START_COL = 14
PROJECT_END_COL = 100

def clean_num(val):
    if pd.isna(val) or val is None: return 0.0
    s = str(val).replace(',', '').strip()
    try: return float(s)
    except: return 0.0

def normalize_code(code_str):
    return re.sub(r'[^A-Z0-9]', '', str(code_str).strip().upper())

def process_order_and_get_summary(user_msg):
    # เชื่อมต่อ Google Sheets
    creds = Credentials.from_service_account_info(json.loads(os.environ.get('GOOGLE_CREDENTIALS_JSON')))
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1ngb3u6xzE6m0QSre1gTwFxm0_hElAavyPGKkOHY98Vc/edit")

    # ดึงค่าจาก Input
    lines = user_msg.strip().split('\n')
    input_dict = {}
    for line in lines:
        parts = line.strip().split()
        if len(parts) >= 2: input_dict[normalize_code(parts[0])] = clean_num(parts[1])
        elif len(parts) == 1: input_dict[normalize_code(parts[0])] = 1.0

    # ประมวลผลไฟล์ Excel
    all_xlsx = glob.glob("*.xlsx")
    global_code_map = {}
    for file_path in all_xlsx:
        df_raw = pd.read_excel(file_path, header=None, engine='openpyxl')
        headers = df_raw.iloc[HEADER_ROW, :].astype(str).tolist()
        for r in range(HEADER_ROW + 1, len(df_raw)):
            code = normalize_code(df_raw.iloc[r, CODE_B_INDEX])
            if code: global_code_map[code] = (df_raw, r, headers)

    report_items = []
    for code, qty_needed in input_dict.items():
        if code in global_code_map:
            df_target, target_row, headers = global_code_map[code]
            balance = clean_num(df_target.iloc[target_row, BALANCE_COL_INDEX])
            
            # --- ปรับ Logic ช่อง ขาด ---
            if balance >= qty_needed:
                shortage = "มีของ"
            else:
                shortage = int(qty_needed) # คงตัวเลขเดิมไว้
            
            # --- ปรับ Logic ค้นหาโครงการ (ขยายขอบเขตและกรองเข้มงวด) ---
            bookings = {}
            for col_idx in range(PROJECT_START_COL, PROJECT_END_COL):
                proj_name = str(headers[col_idx]).strip()
                # กรอง: ต้องไม่ว่าง, ไม่ใช่ตัวเลขล้วน (เช่น '100'), และยาวไม่เกิน 4 ตัว
                if proj_name and not proj_name.isdigit() and len(proj_name) <= 4:
                    val = clean_num(df_target.iloc[target_row, col_idx])
                    if val > 0: bookings[proj_name] = int(val)
            
            report_items.append({
                'code': str(df_target.iloc[target_row, CODE_B_INDEX]),
                'desc': str(df_target.iloc[target_row, DESC_COL_INDEX]),
                'shortage': shortage,
                'balance': int(balance),
                'new': int(clean_num(df_target.iloc[target_row, NEW_COL_INDEX])),
                'old': int(clean_num(df_target.iloc[target_row, OLD_COL_INDEX])),
                'bookings': bookings
            })

    # เตรียมบันทึกและส่งกลับ
    summary_text = "📊 รายงานสรุปสต็อก:\n"
    table_data = [["รหัสสินค้า", "รายการ", "สต็อก Balance", "ขาด", "ของใหม่", "ของเก่า", "ติดจอง"]]
    
    for item in report_items:
        summary_text += f"\n📦 {item['code']} ({item['desc']})\n- สต็อกรวม: {item['balance']}\n- ขาด: {item['shortage']}\n- ของใหม่: {item['new']}\n- ของเก่า: {item['old']}\n"
        booking_str = ""
        if item['bookings']:
            summary_text += "- ติดจอง:\n"
            booking_items = []
            for proj, qty in item['bookings'].items():
                summary_text += f"  • {proj}: {qty}\n"
                booking_items.append(f"{proj}:{qty}")
            booking_str = ", ".join(booking_items)
        
        table_data.append([item['code'], item['desc'], item['balance'], item['shortage'], item['new'], item['old'], booking_str])

    # บันทึกลงชีต Summary
    ws = spreadsheet.worksheet("Summary") if "Summary" in [w.title for w in spreadsheet.worksheets()] else spreadsheet.add_worksheet(title="Summary", rows="100", cols="10")
    ws.clear()
    ws.update(table_data, 'A1')

    return summary_text

# (ส่วน Callback และ Handler เหมือนเดิม)
@app.route("/callback", methods=['POST'])
def callback():
    body = request.get_data(as_text=True)
    try: handler.handle(body, request.headers['X-Line-Signature'])
    except: abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    reply_text = process_order_and_get_summary(event.message.text)
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply_text))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
