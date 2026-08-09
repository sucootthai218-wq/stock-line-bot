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

FILE_IDS = [
    "1M7YJzGsNRTSQswyxdHHiDDXAuX1h7U4a",
    "1HwUNEZ1wwwne2-ZG0ogluODTdmAAdTSg",
]

GSHEET_URL = "https://docs.google.com/spreadsheets/d/1ngb3u6xzE6m0QSre1gTwFxm0_hElAavyPGKkOHY98Vc/edit?gid=0#gid=0"

HEADER_ROW = 4
CODE_B_INDEX = 1        # คอลัมน์ B (รหัสสินค้าหลัก)
CODE_C_INDEX = 2        # คอลัมน์ C (รหัสสำรอง/Code Express)
DESC_COL_INDEX = 3     
NEW_COL_INDEX = 12      # คอลัมน์ M
OLD_COL_INDEX = 13      # คอลัมน์ N
BALANCE_COL_INDEX = 63  # คอลัมน์ BL

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
    
    # 1. บันทึกข้อความที่ส่งมาลงใน Input_Order
    try:
        ws_input = spreadsheet.worksheet("Input_Order")
    except:
        ws_input = spreadsheet.add_worksheet(title="Input_Order", rows="100", cols="10")
        
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

    # 2. รันระบบประมวลผลสต็อก (ดึงไฟล์และคำนวณ)
    existing_files = glob.glob("*.xlsx")
    if len(existing_files) < len(FILE_IDS):
        for f in glob.glob("*.xlsx"):
            try: os.remove(f)
            except: pass
        for file_id in FILE_IDS:
            gdown.download(id=file_id, output=f"{file_id}.xlsx", quiet=True)
            
    all_xlsx = [f for f in glob.glob("*.xlsx") if not os.path.basename(f).startswith("~$")]
    
    global_code_map = {} 
    project_col_map = {}
    combined_headers = []

    for file_path in all_xlsx:
        excel_file_obj = pd.ExcelFile(file_path)
        sheet_name = excel_file_obj.sheet_names[0]
        df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
        num_cols = df_raw.shape[1]
        
        for r in range(HEADER_ROW + 1, len(df_raw)):
            for col_idx in [CODE_B_INDEX, CODE_C_INDEX]:
                c_val = df_raw.iloc[r, col_idx]
                if pd.notna(c_val):
                    norm_c = normalize_code(c_val)
                    if norm_c and norm_c not in ["NAN", "NONE", "0", "CODE", "รหัสสินค้า", "รหัส", "NO"]:
                        if norm_c not in global_code_map:
                            global_code_map[norm_c] = (df_raw, r)
                            
        for c in range(num_cols):
            txts = [str(df_raw.iloc[r, c]).strip() for r in range(HEADER_ROW, min(HEADER_ROW + 3, len(df_raw))) if pd.notna(df_raw.iloc[r, c])]
            combined_headers.append(" ".join(txts))

    df_input = pd.DataFrame(ws_input.get_all_records())
    excluded_kw = ["no", "code", "รายการ", "order", "category", "weight", "quantity", "qty", "on hand", "balance", "ขาด", "old", "new", "broken", "po", "repair", "total", "dift"]

    sample_df = all_xlsx and pd.read_excel(all_xlsx[0], sheet_name=0, header=None, engine='openpyxl')
    num_cols_sample = sample_df.shape[1] if sample_df is not None else 0

    for c in range(num_cols_sample):
        if c in [CODE_B_INDEX, CODE_C_INDEX, DESC_COL_INDEX, NEW_COL_INDEX, OLD_COL_INDEX, BALANCE_COL_INDEX]: continue
        h_text = combined_headers[c].lower() if c < len(combined_headers) else ""
        if any(k in h_text for k in excluded_kw): continue
        
        if c < len(combined_headers):
            pcode_match = re.search(r'\b([A-Z0-9]{2,6})\b', combined_headers[c].upper())
            if pcode_match:
                pcode = pcode_match.group(1)
                if pcode not in ["NEW", "OLD", "QTY", "KG", "TOTAL", "PO", "BROKEN", "REPAIR"]:
                    if pcode not in project_col_map: project_col_map[pcode] = []
                    project_col_map[pcode].append(c)

    report_items = []
    input_code_col = next((c for c in df_input.columns if "code" in str(c).lower() or "รหัส" in str(c)), df_input.columns[0])
    input_qty_col = next((c for c in df_input.columns if "qty" in str(c).lower() or "จำนวน" in str(c)), df_input.columns[1])

    for _, row in df_input.iterrows():
        raw_code = str(row[input_code_col]).strip()
        norm_input = normalize_code(raw_code)
        qty_needed = clean_num(row[input_qty_col])
        
        match_data = global_code_map.get(norm_input)
        if match_data is not None:
            df_target, target_row = match_data
            
            new_qty = clean_num(df_target.iloc[target_row, NEW_COL_INDEX])
            old_qty = clean_num(df_target.iloc[target_row, OLD_COL_INDEX])
            balance = clean_num(df_target.iloc[target_row, BALANCE_COL_INDEX])
            
            shortage = qty_needed if balance < 0 else max(0, qty_needed - balance)
            
            proj_bookings = {p: int(sum(clean_num(df_target.iloc[target_row, c]) for c in cols)) 
                             for p, cols in project_col_map.items() if sum(clean_num(df_target.iloc[target_row, c]) for c in cols) != 0}

            report_items.append({
                'code': str(df_target.iloc[target_row, CODE_B_INDEX]), 
                'desc': str(df_target.iloc[target_row, DESC_COL_INDEX]),
                'new': int(new_qty),
                'old': int(old_qty),
                'on_hand': int(new_qty + old_qty),
                'balance': int(balance),
                'shortage': "มีของ" if shortage == 0 else int(shortage),
                'bookings': proj_bookings
            })
        else:
            report_items.append({
                'code': raw_code, 
                'desc': "⚠️ ไม่พบรหัส",
                'new': "-",
                'old': "-",
                'on_hand': "-",
                'balance': "-",
                'shortage': int(qty_needed),
                'bookings': {}
            })

    # 3. สร้างข้อความสรุปผลส่งกลับไปที่หน้าจอ LINE และอัปเดตชีต Summary
    summary_text = "📊 รายงานสรุปสต็อก:\n"
    for item in report_items:
        summary_text += f"\n📦 {item['code']} ({item['desc']})\n- สต็อกรวม: {item['balance']}\n- ขาด: {item['shortage']}\n- ของใหม่: {item['new']}\n- ของเก่า: {item['old']}\n"
        if item['bookings']:
            summary_text += "- ติดจอง:\n"
            for p, q in item['bookings'].items():
                summary_text += f"  • {p}: {q}\n"

    active_projects = sorted(list({p for item in report_items for p in item['bookings'].keys()}))
    header = ["รหัสสินค้า", "รายการ", "New", "Old", "On Hand", "สต็อก Balance", "ขาด"] + active_projects
    table_data = [header] + [
        [i['code'], i['desc'], i['new'], i['old'], i['on_hand'], i['balance'], i['shortage']] + 
        [i['bookings'].get(p, "-") for p in active_projects] 
        for i in report_items
    ]

    try:
        ws_summary = spreadsheet.worksheet("Summary") if "Summary" in [w.title for w in spreadsheet.worksheets()] else spreadsheet.add_worksheet(title="Summary", rows="100", cols="30")
        ws_summary.clear()
        ws_summary.update(table_data, 'A1')
    except Exception as e:
        print(f"Error updating Summary: {e}")

    return summary_text

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers.get('X-Line-Signature', '')
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
