#!/bin/bash
# Script cài đặt môi trường cho ứng dụng Quản lý xuất nhập kho bằng QR Code
set -e

# Tạo virtualenv nếu chưa có
if [ ! -d ".venv" ]; then
    python3 -m venv .venv
fi

# Kích hoạt virtualenv
source .venv/bin/activate

# Cài đặt các thư viện cần thiết
pip install --upgrade pip
pip install -r requirements.txt

echo "\nCài đặt hoàn tất! Để chạy server:"
echo "source .venv/bin/activate && python app.py"
