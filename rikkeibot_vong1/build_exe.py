import subprocess
import shutil
import os
import sys

# Đảm bảo in tiếng Việt không lỗi trên terminal
if not hasattr(sys.stdout, 'original_stdout'):
    try:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    except Exception:
        pass

def build():
    print(">>> Bat dau qua trinh dong goi ung dung Rikkei Bot Vong 1 bang PyInstaller...")
    
    # Đảm bảo đang chạy đúng thư mục rikkeibot_vong1
    base_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(base_dir)
    
    # 1. Xóa các thư mục build cũ nếu có
    for d in ['build', 'dist', 'rikkeibot_vong1_Deploy']:
        if os.path.exists(d):
            print(f"-> Dang xoa thu muc cu: {d}")
            shutil.rmtree(d, ignore_errors=True)
            
    # 2. Xây dựng câu lệnh PyInstaller
    cmd = [
        "pyinstaller",
        "--noconfirm",
        "--onedir",
        "--name=app",
        "--add-data=templates;templates",
        "--add-data=template_CNTT.docx;.",
        "--add-data=template_QTKD.docx;.",
        "--hidden-import=full_automation_lark_app",
        "--hidden-import=auto_update_new_app",
        "app.py"
    ]
    
    print(f"-> Dang chay lenh: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='replace')
    
    if result.returncode != 0:
        print("X Luan dong goi PyInstaller BI LOI:")
        print(result.stdout)
        print(result.stderr)
        sys.exit(1)
        
    print("-> Hoan tat PyInstaller build.")
    
    # 3. Chuẩn bị thư mục Deploy cuối cùng
    deploy_dir = "rikkeibot_vong1_Deploy"
    print(f"-> Chuon bi thu muc Deploy: {deploy_dir}")
    
    # Copy từ dist/app sang rikkeibot_vong1_Deploy
    shutil.copytree("dist/app", deploy_dir)
    
    # Tạo sẵn file config.json mẫu
    config_sample = {
        "APP_ID": "cli_aa8847bf0f79deef",
        "APP_SECRET": "zBA9rmZ7ScgEtIS1OFFOedGIHaz11V8g",
        "APP_TOKEN": "UWhibaHkCaWLausconRl8krsgZg",
        "TABLE_CNTT": "tblZoFTIttbvejQ7",
        "TABLE_QTKD": "tbl4gHCX7H8eGjaA"
    }
    
    with open(os.path.join(deploy_dir, "config.json"), "w", encoding="utf-8") as f:
        json.dump(config_sample, f, indent=4)
        
    # Tạo hướng dẫn HDSD.txt ngắn gọn
    hdsd_content = """BOT RIKKEI VONG 1 - HUONG DAN SU DUNG
==================================================

1. CAU HINH QUYEN TREN LARK BITABLE (QUAN TRONG):
   - De bot co the doc du lieu va dang tai file PDF len cot "File hoc vien",
     ban bat buoc phai them Custom App lam cong tac vien cua Base.
   - Cac buoc thuc hien:
     + Mo Base tuyen sinh tren trinh duyet.
     + Click nut "Chia se" (Share) o goc tren ben phai Base.
     + Click "Them thanh vien" (Add Collaborator).
     + Tim kiem ung dung Lark co App ID: cli_aa8847bf0f79deef va chon.
     + Phan quyen toi thieu cho App la "Chinh sua" (Edit) hoac "Quan tri" (Manage).

2. CACH KHOI CHAY BOT:
   - Click dup chay file `app.exe` trong thu muc nay.
   - Mot giao dien cua so mau den (Console) se hien len thong bao khoi dong Web Server thanh cong.
   - Mo trinh duyet Web va truy cap dia chi: http://localhost:5000

3. HUONG DAN TREN DASHBOARD WEB:
   - Cac thong tin ket noi API Lark da duoc dien san theo dot tuyen sinh Vong 1.
   - Nhan "Chi cap nhat moi" de bot quet 2 bang CNTT + QTKD, chi lam cac hoc vien chua co file.
   - Nhan "Ghi de toan bo" neu muon chay lai va cap nhat de file PDF moi cho tat ca hoc vien.
   - Nhat ky log thoi gian thuc se hien thi o cuoi trang.
   - Sau khi chay hoan tat, danh sach hoc vien se duoc dong bo. Ban co tthe nhan nut "Tai ZIP"
     o phan tren de tai toan bo file PDF du phong ve may tinh.
"""
    with open(os.path.join(deploy_dir, "HDSD.txt"), "w", encoding="utf-8") as f:
        f.write(hdsd_content)
        
    print(f"-> DA HOAN THANH BO DEPLOY TAI THU MUC: {deploy_dir}")

if __name__ == "__main__":
    import json
    build()
