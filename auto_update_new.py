import os
import sys
import io
import time
import shutil
import requests
import datetime
import re
import pandas as pd
from pathlib import Path
import zipfile
from mailmerge import MailMerge
from docx2pdf import convert

# Đảm bảo in tiếng Việt không lỗi trên terminal
if not hasattr(sys.stdout, 'original_stdout'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import json

# ==========================================
# 1. CẤU HÌNH API LARK
# ==========================================
try:
    with open('config.json', 'r', encoding='utf-8') as f:
        _config = json.load(f)
        APP_ID = _config.get('APP_ID')
        APP_SECRET = _config.get('APP_SECRET')
        APP_TOKEN = _config.get('APP_TOKEN')
        TABLE_ID = _config.get('TABLE_ID')
except FileNotFoundError:
    print("LỖI: Không tìm thấy file config.json. Hãy tạo file config.json chứa các API keys.")
    sys.exit(1)
except Exception as e:
    print(f"LỖI đọc file config.json: {e}")
    sys.exit(1)

FILE_FIELD_NAME = 'File học viên'
SKIP_EXISTING = True  # Đổi thành True để bỏ qua những người đã có file, chỉ update người mới

# ==========================================
# 2. BỘ TỪ ĐIỂN VÀ HÀM LÀM SẠCH (CLEAN DATA)
# ==========================================
AUX_WORDS = sorted([
    'số nhà', 'thành phố', 'thị trấn', 'thị xã', 'khu phố', 'dân phố',
    'tỉnh', 'xã', 'thôn', 'xóm', 'phường', 'quận', 'huyện',
    'khối', 'tổ', 'ngõ', 'ngách', 'đường', 'phố',
    'số', 'khu', 'ấp', 'làng', 'hẻm', 'lô',
], key=len, reverse=True)

PLUS_CODE_RE = re.compile(r'\b[A-Z0-9]{4,8}\+[A-Z0-9]{2,4}\b', re.IGNORECASE)

ABBREV_MAP = [
    (re.compile(r'\bTP\.\s*', re.IGNORECASE), 'Thành phố '),
    (re.compile(r'\bTX\.\s*', re.IGNORECASE), 'Thị xã '),
    (re.compile(r'\bTT\.\s*', re.IGNORECASE), 'Thị trấn '),
    (re.compile(r'\bP\.\s*',  re.IGNORECASE), 'Phường '),
    (re.compile(r'\bQ\.\s*',  re.IGNORECASE), 'Quận '),
    (re.compile(r'\bH\.\s*',  re.IGNORECASE), 'Huyện '),
    (re.compile(r'\bX\.\s*',  re.IGNORECASE), 'Xã '),
    (re.compile(r'\bHN\b',    re.IGNORECASE), 'Hà Nội'),
    (re.compile(r'\bHCM\b',   re.IGNORECASE), 'Hồ Chí Minh'),
    (re.compile(r'\bTPHCM\b', re.IGNORECASE), 'Hồ Chí Minh'),
    (re.compile(r'\bĐN\b',    re.IGNORECASE), 'Đà Nẵng'),
]

PROVINCES = [
    'Hà Nội', 'Hồ Chí Minh', 'Hải Phòng', 'Đà Nẵng', 'Cần Thơ',
    'An Giang', 'Bà Rịa Vũng Tàu', 'Bắc Giang', 'Bắc Kạn', 'Bạc Liêu',
    'Bắc Ninh', 'Bến Tre', 'Bình Định', 'Bình Dương', 'Bình Phước',
    'Bình Thuận', 'Cà Mau', 'Cao Bằng', 'Đắk Lắk', 'Đắk Nông',
    'Điện Biên', 'Đồng Nai', 'Đồng Tháp', 'Gia Lai', 'Hà Giang',
    'Hà Nam', 'Hà Tĩnh', 'Hải Dương', 'Hậu Giang', 'Hòa Bình',
    'Hưng Yên', 'Khánh Hòa', 'Kiên Giang', 'Kon Tum', 'Lai Châu',
    'Lâm Đồng', 'Lạng Sơn', 'Lào Cai', 'Long An', 'Nam Định',
    'Nghệ An', 'Ninh Bình', 'Ninh Thuận', 'Phú Thọ', 'Phú Yên',
    'Quảng Bình', 'Quảng Nam', 'Quảng Ngãi', 'Quảng Ninh', 'Quảng Trị',
    'Sóc Trăng', 'Sơn La', 'Tây Ninh', 'Thái Bình', 'Thái Nguyên',
    'Thanh Hóa', 'Thừa Thiên Huế', 'Tiền Giang', 'Trà Vinh', 'Tuyên Quang',
    'Vĩnh Long', 'Vĩnh Phúc', 'Yên Bái',
]
_PROV_LOWER = {p.lower(): p for p in PROVINCES}

def title_viet(word):
    return word[0].upper() + word[1:] if word else word

def normalize_abbrev(text):
    for pattern, replacement in ABBREV_MAP:
        text = pattern.sub(replacement, text)
    return text

def clean_ten(val):
    if not val: return ''
    s = re.sub(r'\s+', ' ', re.sub(r'\d+', '', str(val))).strip()
    return ' '.join(w[0].upper() + w[1:].lower() if w else w for w in s.split())

def clean_cccd(val):
    if val is None: return ''
    try: s = str(int(round(float(str(val)))))
    except Exception: s = re.sub(r'\D', '', str(val))
    s = re.sub(r'\D', '', s)
    return s.zfill(12)[:12]

def clean_sdt(val):
    if val is None: return ''
    try: s = str(int(round(float(str(val)))))
    except Exception: s = re.sub(r'\D', '', str(val))
    s = re.sub(r'\D', '', s)
    if len(s) == 9: s = '0' + s
    return s.zfill(10)[-10:]

def clean_ngaysinh_val(val):
    if not val: return ''
    try:
        # Nếu val là string timestamp mili giây từ Lark
        if str(val).isdigit() and len(str(val)) > 10:
            dt = datetime.datetime.fromtimestamp(int(val)/1000)
            return dt.strftime('%d/%m/%Y')
        return pd.to_datetime(val).strftime('%d/%m/%Y')
    except Exception:
        return str(val)

def clean_diachi(addr):
    if not addr: return ''
    s = str(addr)
    s = PLUS_CODE_RE.sub('', s)
    s = re.sub(r',\s*,', ',', s)
    s = re.sub(r'^[\s,]+|[\s,]+$', '', s)
    s = re.sub(r'\s+', ' ', s)
    s = s.rstrip('.,;!:')

    parts = [normalize_abbrev(p.strip()) for p in s.split(',')]
    new_parts = []
    for part in parts:
        part = part.strip().rstrip('.,;!:')
        if not part: continue
        part_lower = part.lower()
        matched = None
        for aux in AUX_WORDS:
            if re.match(r'^' + re.escape(aux) + r'(\s|$)', part_lower):
                matched = aux
                break
        if matched:
            aux_cap = matched[0].upper() + matched[1:]
            rest = part[len(matched):].strip()
            rest_cap = ' '.join(title_viet(w) for w in rest.split())
            new_parts.append(aux_cap + (' ' + rest_cap if rest_cap else ''))
        else:
            new_parts.append(' '.join(title_viet(w) for w in part.split()))

    if not new_parts: return ''
    return ', '.join(new_parts) + '.'

def extract_diaphuong(cleaned_addr):
    if not cleaned_addr: return ''
    s = cleaned_addr.rstrip('.')
    parts = [p.strip() for p in re.split(r'[,\-–.]', s) if p.strip()]
    if not parts: return ''
    
    last_part = parts[-1]
    for aux in AUX_WORDS:
        if re.match(r'^' + re.escape(aux) + r'(\s|$)', last_part, re.IGNORECASE):
            last_part = last_part[len(aux):].strip()
            break

    result = ' '.join(w[0].upper() + w[1:] if w else w for w in last_part.split())
    if result.lower() in _PROV_LOWER: return _PROV_LOWER[result.lower()]

    words = s.split()
    for n in [3, 2, 1]:
        if len(words) >= n:
            candidate = ' '.join(words[-n:])
            if candidate.lower() in _PROV_LOWER:
                return _PROV_LOWER[candidate.lower()]
    return result

# ==========================================
# 3. KẾT NỐI API LARK
# ==========================================

def get_tenant_access_token():
    print("🔑 Đang lấy Access Token...")
    url = "https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal"
    res = requests.post(url, json={"app_id": APP_ID, "app_secret": APP_SECRET}).json()
    token = res.get('tenant_access_token')
    if not token:
        print("Không thể lấy token:", res)
        sys.exit(1)
    return token

def parse_lark_text(val):
    if isinstance(val, list) and len(val) > 0:
        return str(val[0].get('text', val[0])).strip()
    return str(val).strip() if val else ""

def fetch_and_clean_records(token):
    print("📥 Đang lấy dữ liệu thô từ Lark và TỰ ĐỘNG LÀM SẠCH...")
    records = []
    url = f"https://open.larksuite.com/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records"
    headers = {"Authorization": f"Bearer {token}"}
    params = {"page_size": 100}
    
    while True:
        res = requests.get(url, headers=headers, params=params).json()
        if res.get('code') != 0:
            print("Lỗi khi lấy records:", res)
            break
            
        for item in res.get('data', {}).get('items', []):
            fields = item.get('fields', {})
            # 1. Trích xuất raw
            raw_cccd = parse_lark_text(fields.get('CCCD/ CMT'))
            if not raw_cccd: continue  
            
            raw_hoten = parse_lark_text(fields.get('Họ và tên học viên'))
            raw_ngaysinh = fields.get('Ngày sinh học viên')
            raw_sdt = parse_lark_text(fields.get('SĐT học viên'))
            raw_diachi = parse_lark_text(fields.get('Địa chỉ thường trú'))
            nganh = parse_lark_text(fields.get('Loại hình khóa học'))
            da_co_file = bool(fields.get(FILE_FIELD_NAME))
            
            # 2. Làm sạch (Clean Data)
            clean_hoten_val = clean_ten(raw_hoten)
            clean_cccd_val = clean_cccd(raw_cccd)
            clean_ngaysinh_val_str = clean_ngaysinh_val(raw_ngaysinh)
            clean_sdt_val = clean_sdt(raw_sdt)
            clean_diachi_val = clean_diachi(raw_diachi)
            
            # Tự động trích xuất Nơi sinh từ địa chỉ đã clean, hoặc lấy trực tiếp từ cột Địa Phương
            raw_diaphuong = parse_lark_text(fields.get('Địa Phương'))
            noi_sinh = raw_diaphuong if raw_diaphuong else extract_diaphuong(clean_diachi_val)
            
            raw_thoigian = parse_lark_text(fields.get('Thời gian gửi email đủ ĐK trúng tuyển'))
            
            rec = {
                'record_id': item.get('record_id'),
                'Họ và tên': clean_hoten_val,
                'CCCD': clean_cccd_val,
                'Ngày sinh': clean_ngaysinh_val_str,
                'Nơi sinh': noi_sinh,
                'Địa chỉ': clean_diachi_val,
                'SĐT': clean_sdt_val,
                'Ngành': nganh,
                'Thời gian email': raw_thoigian,
                'Đã có file': da_co_file
            }
            records.append(rec)
                
        if not res.get('data', {}).get('has_more', False):
            break
        params['page_token'] = res.get('data', {}).get('page_token')
        
    print(f"-> Đã lấy và chuẩn hóa thành công {len(records)} hồ sơ có số CCCD.")
    return records

def create_excel_backup(records):
    print("💾 Đang xuất dữ liệu ra file Excel Backup (Chia theo ngày)...")
    df = pd.DataFrame(records)
    if 'record_id' in df.columns:
        df = df.drop(columns=['record_id'])
    
    # Định dạng cột text để không bị mất số 0
    df['CCCD'] = df['CCCD'].astype(str)
    df['SĐT'] = df['SĐT'].astype(str)
    
    # Phân loại và xuất nhiều file Excel
    groups = df.groupby('Thời gian email', dropna=False)
    for name, group in groups:
        safe_name = str(name).strip()
        if not safe_name or safe_name.lower() == 'nan' or safe_name.lower() == 'none':
            file_name = 'Data_Lark_Cleaned_Backup_ChuaCoNgay.xlsx'
        else:
            safe_name_file = re.sub(r'[\\/*?:"<>|]', '-', safe_name)
            file_name = f'Data_Lark_Cleaned_Backup_{safe_name_file}.xlsx'
            
        group.to_excel(file_name, index=False)
        print(f"-> Đã lưu file {file_name} ({len(group)} dòng)")

def upload_file_to_lark(token, file_name, file_bytes):
    url = "https://open.larksuite.com/open-apis/drive/v1/medias/upload_all"
    headers = {"Authorization": f"Bearer {token}"}
    files = {'file': (file_name, file_bytes, 'application/pdf')}
    data = {
        'file_name': file_name,
        'parent_type': 'bitable_file',
        'parent_node': APP_TOKEN,
        'size': len(file_bytes)
    }
    
    res = requests.post(url, headers=headers, data=data, files=files).json()
    if res.get('code') != 0:
        print(f" [!] Lỗi upload {file_name}: {res}")
        return None
    return res.get('data', {}).get('file_token')

def update_bitable_record(token, record_id, file_token):
    url = f"https://open.larksuite.com/open-apis/bitable/v1/apps/{APP_TOKEN}/tables/{TABLE_ID}/records/{record_id}"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    
    # 1. Thao tác xóa cột file cũ trước để tránh lỗi xung đột hoặc tích tụ nhiều file trong một ô Bitable
    try:
        clear_payload = {"fields": {FILE_FIELD_NAME: []}}
        requests.put(url, headers=headers, json=clear_payload)
    except Exception:
        pass
        
    # 2. Cập nhật file học viên mới tinh
    payload = {"fields": {FILE_FIELD_NAME: [{"file_token": file_token}]}}
    res = requests.put(url, headers=headers, json=payload).json()
    return res.get('code') == 0

# ==========================================
# 4. CHẠY PIPELINE
# ==========================================
def run_automation():
    token = get_tenant_access_token()
    records = fetch_and_clean_records(token)
    
    if not records:
        print("Không có dữ liệu hợp lệ để xử lý.")
        return
        
    create_excel_backup(records)
    
    # Chuẩn bị thư mục tạm
    temp_dir = Path("_temp_all_in_one")
    if temp_dir.exists(): shutil.rmtree(temp_dir, ignore_errors=True)
    temp_dir.mkdir()
    
    # Lọc học sinh cần xử lý
    students_to_process = []
    for r in records:
        if SKIP_EXISTING and r['Đã có file']:
            continue
        students_to_process.append(r)
        
    if not students_to_process:
        print("🎉 Tất cả học sinh đều đã có File học viên trên Lark. Không cần chạy thêm.")
        # Xuất toàn bộ học sinh sang JSON với trạng thái done
        processed_students = []
        for r in records:
            processed_students.append({
                "name": r['Họ và tên'],
                "cccd": r['CCCD'],
                "dob": r['Ngày sinh'],
                "phone": r['SĐT'],
                "major": "CNTT" if "CNTT" in str(r['Ngành']).upper() else "QTKD",
                "status": "done"
            })
        try:
            with open('processed_students.json', 'w', encoding='utf-8') as f:
                json.dump(processed_students, f, ensure_ascii=False, indent=4)
        except Exception:
            pass
        return
        
    print(f"\n⚙️ Bắt đầu trộn thư (Mail Merge) cho {len(students_to_process)} học sinh...")
    docx_list = []
    
    for r in students_to_process:
        nganh = r['Ngành'].upper()
        if "CNTT" in nganh:
            template = "Mau_mail_CNTT.docx"
        elif "QTKD" in nganh:
            template = "Mau_mail_QTKDS.docx"
        else:
            template = "Mau_mail_CNTT.docx"
            
        merge_data = {
            "Họ_và_tên": r['Họ và tên'],
            "CCCD": r['CCCD'],
            "Ngày_sinh": r['Ngày sinh'],
            "Nơi_sinh": r['Nơi sinh'],
            "Địa_phương": r['Nơi sinh'],
            "Địa_chỉ_thường_trú": r['Địa chỉ'],
            "Địa_chỉ": r['Địa chỉ'],
            "Địa_chỉ_cụ_thể": r['Địa chỉ'],
            "SĐT": r['SĐT']
        }
        
        file_name = f"{r['Họ và tên']}_{r['CCCD']}"
        docx_path = temp_dir / f"{file_name}.docx"
        
        document = MailMerge(template)
        document.merge(**merge_data)
        document.write(docx_path)
        document.close()
        
        docx_list.append({
            'record_id': r['record_id'],
            'file_name_base': file_name,
            'docx_path': str(docx_path),
            'nganh': r['Ngành']
        })
        
    print("-> Xong phần trộn thư.")
    
    print("\n📄 Bắt đầu chuyển đổi hàng loạt sang PDF (Mất vài phút)...")
    if sys.platform == "win32":
        convert(str(temp_dir))
    else:
        import subprocess
        docx_files = list(temp_dir.glob("*.docx"))
        total_conv = len(docx_files)
        print(f"Đang dùng LibreOffice để chuyển đổi DOCX -> PDF trên Linux ({total_conv} file)...")
        if total_conv > 0:
            batch_size = 50
            print(f"🚀 Đang chia nhỏ thành các nhóm {batch_size} file để chuyển đổi hàng loạt ổn định...")
            for i in range(0, total_conv, batch_size):
                batch = docx_files[i:i+batch_size]
                print(f"-> Đang convert PDF nhóm {i//batch_size + 1} ({len(batch)} file)...", flush=True)
                cmd = [
                    "libreoffice", "--headless", "--convert-to", "pdf", 
                    "--outdir", str(temp_dir)
                ] + [str(f) for f in batch]
                subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
            # Kháng lỗi nâng cao (Self-healing): Tự động phát hiện và convert bổ sung từng file bị thiếu
            missing_files = []
            for f in docx_files:
                pdf_path = temp_dir / f"{f.stem}.pdf"
                if not pdf_path.exists():
                    missing_files.append(f)
            
            if missing_files:
                print(f"⚠ Phát hiện {len(missing_files)} file PDF bị thiếu. Tiến hành convert bổ sung từng file...", flush=True)
                for f in missing_files:
                    subprocess.run([
                        "libreoffice", "--headless", "--convert-to", "pdf", 
                        "--outdir", str(temp_dir), str(f)
                    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    
            print(f"-> Đã convert thành công toàn bộ {total_conv}/{total_conv} file sang PDF!")
    
    print("\n📦 Đang nén các file PDF vào file ZIP theo ngành...")
    zip_cntt = zipfile.ZipFile("PDF_TrungTuyen_CNTT.zip", 'w', zipfile.ZIP_DEFLATED)
    zip_qtkd = zipfile.ZipFile("PDF_TrungTuyen_QTKD.zip", 'w', zipfile.ZIP_DEFLATED)
    
    for item in docx_list:
        pdf_path = temp_dir / f"{item['file_name_base']}.pdf"
        if pdf_path.exists():
            nganh_upper = str(item.get('nganh', '')).upper()
            if "QTKD" in nganh_upper:
                zip_qtkd.write(pdf_path, arcname=pdf_path.name)
            else:
                zip_cntt.write(pdf_path, arcname=pdf_path.name)
                
    zip_cntt.close()
    zip_qtkd.close()
    print("-> Đã tạo 2 file ZIP thành công (CNTT và QTKD) để bạn dự phòng!")
 
    print("\n☁️ Bắt đầu Upload lên Lark Bitable...")
    success_count = 0
    total = len(docx_list)
    
    for idx, item in enumerate(docx_list, 1):
        pdf_path = temp_dir / f"{item['file_name_base']}.pdf"
        print(f"[{idx}/{total}] Đang tải lên {pdf_path.name} ...", end="", flush=True)
        
        # Mặc định status thất bại
        item['status'] = 'failure'
        
        if not pdf_path.exists():
            print(" LỖI: Không tìm thấy file PDF!")
            continue
            
        with open(pdf_path, 'rb') as f:
            file_bytes = f.read()
            
        try:
            token_file = upload_file_to_lark(token, pdf_path.name, file_bytes)
            if token_file:
                if update_bitable_record(token, item['record_id'], token_file):
                    print(" Thành công! ✓")
                    success_count += 1
                    item['status'] = 'success'
                else:
                    print(" LỖI Cập nhật!")
            else:
                print(" LỖI Upload!")
        except Exception as e:
            print(f" LỖI Kết nối: {e}")
            continue
            
    # Xuất thông tin học sinh đã xử lý thực tế ra JSON để hiển thị trên web
    processed_students = []
    for item in docx_list:
        student_info = next((rec for rec in records if rec['record_id'] == item['record_id']), None)
        if student_info:
            processed_students.append({
                "name": student_info['Họ và tên'],
                "cccd": student_info['CCCD'],
                "dob": student_info['Ngày sinh'],
                "phone": student_info['SĐT'],
                "major": "CNTT" if "CNTT" in str(student_info['Ngành']).upper() else "QTKD",
                "status": "done" if item['status'] == 'success' else "error"
            })
            
    # Thêm cả những học sinh cũ đã có sẵn file trên Lark để hiển thị đầy đủ
    for r in records:
        if SKIP_EXISTING and r['Đã có file']:
            processed_students.append({
                "name": r['Họ và tên'],
                "cccd": r['CCCD'],
                "dob": r['Ngày sinh'],
                "phone": r['SĐT'],
                "major": "CNTT" if "CNTT" in str(r['Ngành']).upper() else "QTKD",
                "status": "done"
            })
            
    try:
        with open('processed_students.json', 'w', encoding='utf-8') as f:
            json.dump(processed_students, f, ensure_ascii=False, indent=4)
        print("-> Đã lưu kết quả xử lý thực tế ra file processed_students.json!")
    except Exception as e:
        print(f"Lỗi ghi file processed_students.json: {e}")
                
    print(f"\n🧹 Đang dọn dẹp thư mục tạm...")
    shutil.rmtree(temp_dir, ignore_errors=True)
    
    print(f"✅ HOÀN TẤT ALL-IN-ONE PIPELINE! Đã đẩy thành công {success_count}/{total} file.")

if __name__ == "__main__":
    run_automation()
