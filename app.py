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
# ดึงค่า LINE Token และ Secret จาก Environment Variables บน Render
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
ON_HAND_COL_INDEX = 62   
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

def process_order_and_get_summary(user_msg):
    creds = get_google_credentials()
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_url(GSHEET_URL)

    # 1. บันทึกข้อความที่ส่งมาลงใน Input_Order (กรองเฉพาะบรรทัดที่เป็นรหัสสินค้า)
    ws_input = spreadsheet.worksheet("Input_Order")
    ws_input.clear()
    
    lines = user_msg.strip().split('\n')
    input_data = [["Code", "Qty"]]
    for line in lines:
        parts = line.strip().split()
        if len(parts) >= 1:
            raw_code = parts[0]
            norm_c = normalize_code(raw_code)
            if not norm_c:
                continue
            qty = parts[1] if len(parts) >= 2 else "1"
            input_data.append([raw_code, qty])

    if len(input_data) <= 1:
        return "❌ กรุณาระบุรหัสสินค้าที่ต้องการตรวจสอบให้ถูกต้อง"

    ws_input.update(input_data, 'A1')

    # 2. ล้างไฟล์ Excel เก่าในเซิร์ฟเวอร์ทิ้งก่อนทุกครั้ง
    for f in glob.glob("*.xlsx"):
        try:
            os.remove(f)
        except:
            pass

    # 3. ค้นหาและดาวน์โหลดไฟล์ .xlsx ทั้งหมดจาก Google Drive Folder อัตโนมัติ
    drive_service = build('drive', 'v3', credentials=creds)
    query = f"'{DRIVE_FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])

    for file in files:
        file_id = file['id']
        file_name = file['name']
        gdown.download(id=file_id, output=file_name, quiet=False)

    # 4. ประมวลผลไฟล์ Excel ที่อยู่ในโฟลเดอร์
    all_xlsx = [f for f in glob.glob("*.xlsx") if not os.path.basename(f).startswith("~$")]
    global_code_map = {}

    excluded_kw = [
        "no", "code", "รายการ", "order", "category", "weight", "quantity", "qty", 
        "on hand", "balance", "ขาด", "old", "new", "broken", "po", "repair", 
        "total", "sum", "รวม", "dift", "maintenance", "import", "pending", "update"
    ]

    for file_path in all_xlsx:
        excel_file_obj = pd.ExcelFile(file_path)
        sheet_name = excel_file_obj.sheet_names[0]
        df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
        num_cols = df_raw.shape[1]
        
        for r in range(HEADER_ROW + 1, len(df_raw)):
            for col_idx in [CODE_B_INDEX, CODE_C_INDEX]:
                if col_idx < num_cols:
                    c_val = df_raw.iloc[r, col_idx]
                    if pd.notna(c_val):
                        norm_c = normalize_code(c_val)
                        if norm_c and norm_c not in ["NAN", "NONE", "0", "CODE", "รหัสสินค้า", "รหัส", "NO"]:
                            if norm_c not in global_code_map:
                                global_code_map[norm_c] = (df_raw, r)

    df_input = pd.DataFrame(ws_input.get_all_records())
    report_items = []
    input_code_col = df_input.columns[0]
    input_qty_col = df_input.columns[1]

    for _, row in df_input.iterrows():
        raw_code = str(row[input_code_col]).strip()
        norm_input = normalize_code(raw_code)
        qty_needed = clean_num(row[input_qty_col])
        
        match_data = global_code_map.get(norm_input)
        if match_data is not None:
            df_target, target_row = match_data
            
            if target_row >= len(df_target):
                continue
                
            on_hand_val = clean_num(df_target.iloc[target_row, ON_HAND_COL_INDEX]) if ON_HAND_COL_INDEX < df_target.shape[1] else 0.0
            balance = clean_num(df_target.iloc[target_row, BALANCE_COL_INDEX]) if BALANCE_COL_INDEX < df_target.shape[1] else 0.0
            new_val = clean_num(df_target.iloc[target_row, NEW_COL_INDEX]) if NEW_COL_INDEX < df_target.shape[1] else 0.0
            old_val = clean_num(df_target.iloc[target_row, OLD_COL_INDEX]) if OLD_COL_INDEX < df_target.shape[1] else 0.0
            shortage = qty_needed if balance < 0 else max(0, qty_needed - balance)
            
            proj_bookings = {}
            num_cols = df_target.shape[1]
            start_search_col = max(BALANCE_COL_INDEX + 1, 64)
            
            # วนลูปกวาดหาเฉพาะคอลัมน์โครงการที่อยู่ทางขวา
            for c in range(start_search_col, num_cols):
                # ดึงข้อความหัวตารางมาตรวจสอบก่อนว่าเข้าสู่โซน "หักจอง" หรือยัง
                txts = [str(df_target.iloc[r, c]).strip() for r in range(HEADER_ROW, min(HEADER_ROW + 4, len(df_target)))]
                header_block_str = " ".join(txts).lower()
                
                # [จุดปรับปรุง] ถ้าเจอคำว่า "หักจอง" หรือคำที่เกี่ยวข้อง ให้หยุดกวาดทันทีเพื่อไม่ให้ลามไปอ่านโซนอื่น
                if "หักจอง" in header_block_str or "ยอดรวม" in header_block_str:
                    break
                
                val = clean_num(df_target.iloc[target_row, c])
                if val > 0:
                    proj_name = " ".join([t for t in txts if t and t.lower() not in ["nan", "none"]])
                    proj_lower = proj_name.lower()
                    if proj_name and not any(k in proj_lower for k in excluded_kw):
                        proj_bookings[proj_name] = int(val)

            report_items.append({
                'code': str(df_target.iloc[target_row, CODE_B_INDEX]) if CODE_B_INDEX < df_target.shape[1] else raw_code, 
                'desc': str(df_target.iloc[target_row, DESC_COL_INDEX]) if DESC_COL_INDEX < df_target.shape[1] else "",
                'shortage': "มีของ" if shortage == 0 else int(shortage),
                'on_hand': int(on_hand_val),
                'balance': int(balance),
                'new': int(new_val) if new_val != 0 else "-",
                'old': int(old_val) if old_val != 0 else "-",
                'bookings': proj_bookings
            })
        else:
            report_items.append({
                'code': raw_code, 
                'desc': "ไม่พบรหัส",
                'shortage': int(qty_needed),
                'on_hand': "-",
                'balance': "-",
                'new': "-",
                'old': "-",
                'bookings': {}
            })

    # 5. สร้างข้อความสรุปผลส่งกลับไปที่ LINE
    summary_text = "📊 รายงานสรุปสต็อก:\n"
    for item in report_items:
        summary_text += f"\n📦 {item['code']} ({item['desc']})\n"
        summary_text += f"- On hand: {item['on_hand']}\n"
        summary_text += f"- สต็อก balance: {item['balance']}\n"
        summary_text += f"- ขาด: {item['shortage']}\n"
        summary_text += f"- ของใหม่: {item['new']}\n"
        summary_text += f"- ของเก่า: {item['old']}\n"
        
        if item['bookings']:
            summary_text += "- ติดจอง:\n"
            for p, q in item['bookings'].items():
                summary_text += f"  • {p}: {q}\n"
        else:
            summary_text += "- ติดจอง: -\n"

    # อัปเดตลงชีต Summary ใน Google Sheets
    try:
        active_projects = sorted(list({p for item in report_items for p in item['bookings'].keys()}))
        header = ["รหัสสินค้า", "รายการ", "On hand", "สต็อก Balance", "ขาด", "ของใหม่", "ของเก่า"] + active_projects
        table_data = [header] + [
            [i['code'], i['desc'], i['on_hand'], i['balance'], i['shortage'], i['new'], i['old']] + 
            [i['bookings'].get(p, "-") for p in active_projects] 
            for i in report_items
        ]
        
        ws_summary = spreadsheet.worksheet("Summary") if "Summary" in [w.title for w in spreadsheet.worksheets()] else spreadsheet.add_worksheet(title="Summary", rows="100", cols="30")
        ws_summary.clear()
        ws_summary.update(table_data, 'A1')
    except Exception as e:
        print(f"Warning: Could not update Summary sheet: {e}")

    return summary_text

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_msg = event.message.text.strip()
    try:
        reply_text = process_order_and_get_summary(user_msg)
    except Exception as e:
        reply_text = f"❌ เกิดข้อผิดพลาด: {str(e)}"

    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=reply_text))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
