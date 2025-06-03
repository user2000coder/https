# Quản lý xuất nhập kho bằng QR Code

Ứng dụng web cho phép điện thoại truy cập, quét QR code, nhập số lượng và lưu kết quả vào file CSV trên máy tính.

## Tính năng
- Webserver Python Flask
- Giao diện web thân thiện cho điện thoại
- Quét QR code bằng camera điện thoại (html5-qrcode)
- Nhập số lượng sau khi quét
- Gửi dữ liệu về server và lưu vào file CSV

## Khởi động
1. Cài đặt Python 3 và pip
2. Cài đặt thư viện:
   ```bash
   pip install flask
   ```
3. Chạy server:
   ```bash
   python app.py
   ```
4. Dùng điện thoại truy cập địa chỉ IP của máy tính (ví dụ: http://192.168.x.x:5000)

## Cấu trúc file
- `app.py`: Flask backend
- `templates/index.html`: Giao diện web
- `static/`: Chứa JS, CSS
- `data.csv`: File lưu kết quả
