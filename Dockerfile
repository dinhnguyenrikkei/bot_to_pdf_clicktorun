FROM python:3.10-slim

# Thiết lập thư mục làm việc
WORKDIR /app

# Cài đặt LibreOffice và các font chữ cần thiết
# (Lý do: docx2pdf yêu cầu Microsoft Word, không chạy được trên Linux.
# Nên cài đặt LibreOffice để sau này có thể dùng nó để chuyển đổi docx sang pdf trên Docker)
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    libreoffice \
    fonts-liberation \
    fonts-crosextra-carlito \
    && rm -rf /var/lib/apt/lists/*

# Copy file requirements.txt và cài đặt thư viện
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy toàn bộ mã nguồn vào Docker
COPY . .

# Chạy Flask web app
CMD ["python", "app.py"]
