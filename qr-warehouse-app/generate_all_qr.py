import qrcode
import csv
import os

os.makedirs('static/qr_codes', exist_ok=True)
with open('/workspaces/codespaces-blank/Sumida_IoT/database/products.csv', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    for row in reader:
        qr_code = row.get('qr_code')
        if qr_code:
            img_path = f'static/qr_codes/{qr_code}.png'
            qrcode.make(qr_code).save(img_path)
print('Đã sinh lại QR code cho tất cả sản phẩm.')
