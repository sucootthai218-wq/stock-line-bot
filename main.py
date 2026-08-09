import glob
import os
import re
from google.oauth2.service_account import Credentials
import gspread
import pandas as pd
import gdown

# ========================================================
# 📌 ตั้งค่าพิกัดคอลัมน์และการเชื่อมต่อ
# ========================================================
SERVICE_ACCOUNT_FILE = "credentials.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# File ID ของไฟล์ Excel ทั้ง 2 ไฟล์บน Google Drive
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

print("🔒 กำลังเชื่อมต่อกับ Google Sheets API...")
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
client = gspread.authorize(creds)
spreadsheet = client.open_by_url(GSHEET_URL)

def clean_num(val):
    if pd.isna(val) or val is None: return 0.0
    s = str(val).replace(',', '').strip()
    if s in ["-", "_", "", "nan", "None"]: return 0.0
    try: return float(s)
    except: return 0.0

def normalize_code(code_str):
    if not code_str: return ""
    return re.sub(r'[^A-Z0-9]', '', str(code_str).strip().upper())

# ========================================================
# [ขั้นตอนที่ 0] ตรวจสอบและดาวน์โหลดไฟล์จาก Google Drive
# ========================================================
print("\n--- [ขั้นตอนที่ 0/3] ตรวจสอบและดาวน์โหลดไฟล์จาก Google Drive ---")
existing_files = glob.glob("*.xlsx")
if len(existing_files) < len(FILE_IDS):
    for file_id in FILE_IDS:
        gdown.download(id=file_id, quiet=False)
    print("✅ ดาวน์โหลดไฟล์ครบถ้วนแล้ว")
else:
    print("📁 พบไฟล์ Excel ในเครื่องอยู่แล้ว ข้ามการดาวน์โหลด")

# ========================================================
# [ขั้นตอนที่ 1/3] อ่านไฟล์ Excel ทุกไฟล์ในเครื่องและรวมฐานข้อมูล
# ========================================================
print("\n--- [ขั้นตอนที่ 1/3] กำลังอ่านและรวมข้อมูลจากไฟล์ Excel ทั้งหมด ---")
all_xlsx = [f for f in glob.glob("*.xlsx") if not os.path.basename(f).startswith("~$")]

global_code_map = {} # เก็บข้อมูลรหัสสินค้าทั้งหมดจากทุกไฟล์
project_col_map = {}
combined_headers = []

for file_path in all_xlsx:
    print(f"📄 กำลังอ่านไฟล์: {file_path}")
    excel_file_obj = pd.ExcelFile(file_path)
    sheet_name = excel_file_obj.sheet_names[0] # ใช้ชีตแรกของแต่ละไฟล์
    
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
    num_cols = df_raw.shape[1]
    
    # สแกนหาพิกัดรหัสสินค้าจากคอลัมน์ B และ C
    for r in range(HEADER_ROW + 1, len(df_raw)):
        for col_idx in [CODE_B_INDEX, CODE_C_INDEX]:
            c_val = df_raw.iloc[r, col_idx]
            if pd.notna(c_val):
                norm_c = normalize_code(c_val)
                if norm_c and norm_c not in ["NAN", "NONE", "0", "CODE", "รหัสสินค้า", "รหัส", "NO"]:
                    if norm_c not in global_code_map:
                        global_code_map[norm_c] = (df_raw, r)
                        
    # เก็บ Headers สำหรับหาชื่อโปรเจกต์
    for c in range(num_cols):
        txts = [str(df_raw.iloc[r, c]).strip() for r in range(HEADER_ROW, min(HEADER_ROW + 3, len(df_raw))) if pd.notna(df_raw.iloc[r, c])]
        combined_headers.append(" ".join(txts))

# ========================================================
# [ขั้นตอนที่ 2/3] โหลดข้อมูล Input_Order และคำนวณ
# ========================================================
print("\n--- [ขั้นตอนที่ 2/3] โหลดข้อมูลจากแท็บ Input_Order ---")
ws_input = spreadsheet.worksheet("Input_Order")
df_input = pd.DataFrame(ws_input.get_all_records())

excluded_kw = ["no", "code", "รายการ", "order", "category", "weight", "quantity", "qty", "on hand", "balance", "ขาด", "old", "new", "broken", "po", "repair", "total", "dift"]

# ใช้ df_raw ตัวแรกสุดหรือตัวล่าสุดในการแมปโปรเจกต์
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
found_count = 0
input_code_col = next((c for c in df_input.columns if "code" in str(c).lower() or "รหัส" in str(c)), df_input.columns[0])
input_qty_col = next((c for c in df_input.columns if "qty" in str(c).lower() or "จำนวน" in str(c)), df_input.columns[1])

for _, row in df_input.iterrows():
    raw_code = str(row[input_code_col]).strip()
    norm_input = normalize_code(raw_code)
    qty_needed = clean_num(row[input_qty_col])
    
    match_data = global_code_map.get(norm_input)
    if match_data is not None:
        found_count += 1
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
            'desc': "⚠️ ไม่พบรหัสนี้ในไฟล์ Stock",
            'new': "-", 
            'old': "-", 
            'on_hand': "-", 
            'balance': "-", 
            'shortage': int(qty_needed), 
            'bookings': {}
        })

# ========================================================
# [ขั้นตอนที่ 3/3] บันทึกข้อมูล
# ========================================================
active_projects = sorted(list({p for item in report_items for p in item['bookings'].keys()}))
header = ["รหัสสินค้า", "รายการ", "New", "Old", "On Hand", "สต็อก Balance", "ขาด"] + active_projects
table_data = [header] + [[i['code'], i['desc'], i['new'], i['old'], i['on_hand'], i['balance'], i['shortage']] + 
                         [i['bookings'].get(p, "-") for p in active_projects] for i in report_items]

ws_summary = spreadsheet.worksheet("Summary") if "Summary" in [w.title for w in spreadsheet.worksheets()] else spreadsheet.add_worksheet(title="Summary", rows="100", cols="30")
ws_summary.clear()
ws_summary.update(table_data, 'A1')
print(f"\n✅ อัปเดตข้อมูลเสร็จสิ้น! พบ {found_count} รายการ, ไม่พบ {len(report_items)-found_count} รายการ")

input("\nกด Enter เพื่อปิดโปรแกรม...")