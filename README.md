# 🤖 Rikkei Bot — Lark Bitable Automation

> Tự động hóa quy trình tạo thư trúng tuyển và upload PDF lên Lark Bitable.

## ✨ Tính Năng

| Tính năng | Mô tả |
|---|---|
| **Fetch & Clean Data** | Lấy dữ liệu học viên từ Lark Bitable, tự động chuẩn hóa tên, CCCD, SĐT, địa chỉ |
| **Mail Merge** | Trộn thư tự động theo mẫu DOCX (CNTT / QTKD) |
| **DOCX → PDF** | Chuyển đổi hàng loạt sang PDF |
| **Upload Lark** | Tải PDF lên cột "File học viên" trên Lark Bitable |
| **Excel Backup** | Xuất dữ liệu đã clean ra file Excel, phân theo ngày |
| **ZIP Archive** | Nén PDF theo ngành (CNTT / QTKD) để lưu trữ |
| **Dashboard UI** | Giao diện web quản lý trực quan (`index.html`) |

## 📁 Cấu Trúc Dự Án

```
Bot_rikkei/
├── auto_update_new.py       # Script chính (chỉ xử lý học viên MỚI)
├── full_automation_lark.py  # Script full (xử lý TẤT CẢ học viên)
├── config.json              # API keys Lark (không commit lên git)
├── config.sample.json       # File mẫu config
├── Mau_mail_CNTT.docx       # Template thư trúng tuyển ngành CNTT
├── Mau_mail_QTKDS.docx      # Template thư trúng tuyển ngành QTKD
├── index.html               # Dashboard UI
├── Dockerfile               # Docker deployment
├── requirements.txt         # Python dependencies
└── README.md
```

## 🚀 Bắt Đầu

### 1. Cài đặt

```bash
git clone https://github.com/bmngu/Bot_rikkei.git
cd Bot_rikkei
pip install -r requirements.txt
```

### 2. Cấu hình API Lark

```bash
cp config.sample.json config.json
```

Mở `config.json` và điền thông tin:

```json
{
  "APP_ID": "cli_xxxxxxxxxxxx",
  "APP_SECRET": "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
  "APP_TOKEN": "XxxxXxxxXxxxXxxxXx",
  "TABLE_ID": "tblXxxxXxxxXxxxXx"
}
```

**Cách lấy credentials:**
1. Truy cập [Lark Open Platform](https://open.larksuite.com/app) → tạo App
2. Lấy **App ID** & **App Secret** từ trang Credentials
3. **APP_TOKEN**: lấy từ URL Bitable (`https://xxx.larksuite.com/base/{APP_TOKEN}`)
4. **TABLE_ID**: lấy từ URL table (`...?table={TABLE_ID}`)
5. Cấp quyền cho App: `bitable:app`, `drive:drive`

### 3. Chạy

```bash
# Chỉ xử lý học viên MỚI (chưa có file)
python auto_update_new.py

# Xử lý LẠI TẤT CẢ (ghi đè file cũ)
python full_automation_lark.py
```

### 4. Xem Dashboard

Mở `index.html` trong trình duyệt.

## 🐳 Deploy với Docker

```bash
docker build -t rikkei-bot .
docker run --rm -v $(pwd)/config.json:/app/config.json rikkei-bot
```

> **Lưu ý**: Trên Docker (Linux), bot sử dụng LibreOffice thay vì Microsoft Word để chuyển DOCX → PDF.

## ⚙️ Pipeline Hoạt Động

```
🔑 Auth Token → 📥 Fetch Lark Data → 🧹 Clean Data → 📝 Mail Merge → 📄 DOCX→PDF → ☁️ Upload Lark
```

1. **Auth**: Lấy `tenant_access_token` từ Lark API
2. **Fetch**: Lấy tất cả records từ Bitable (phân trang tự động)
3. **Clean**: Chuẩn hóa tên (Title Case), CCCD (12 số), SĐT (10 số), địa chỉ (viết tắt → đầy đủ)
4. **Mail Merge**: Chọn template theo ngành → trộn dữ liệu vào file DOCX
5. **Convert**: Chuyển DOCX → PDF hàng loạt
6. **Upload**: Upload PDF lên Lark Drive → gắn vào cột "File học viên"

## 📊 API Response Mẫu

<details>
<summary>Lark Bitable Records Response</summary>

```json
{
  "code": 0,
  "msg": "success",
  "data": {
    "has_more": false,
    "total": 247,
    "items": [
      {
        "record_id": "recXxxXxxXxx",
        "fields": {
          "Họ và tên học viên": [{"text": "Nguyễn Văn An"}],
          "CCCD/ CMT": [{"text": "001204012345"}],
          "Ngày sinh học viên": 1078963200000,
          "SĐT học viên": [{"text": "0912345678"}],
          "Địa chỉ thường trú": [{"text": "123 P. Trần Duy Hưng, Q. Cầu Giấy, HN"}],
          "Loại hình khóa học": [{"text": "CNTT"}],
          "Địa Phương": [{"text": "Hà Nội"}],
          "File học viên": null
        }
      }
    ]
  }
}
```
</details>

## 📋 Yêu Cầu Hệ Thống

- Python 3.8+
- Microsoft Word (Windows) hoặc LibreOffice (Linux/Docker)
- Kết nối Internet (Lark API)

## 📄 License

MIT License
