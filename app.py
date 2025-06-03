from flask import Flask, render_template, request, jsonify, redirect, url_for, flash, send_from_directory, send_file
import os
from datetime import datetime
import pandas as pd
import qrcode
from docx import Document
from docx.shared import Inches
import io
import uuid
import tempfile
import openpyxl
import sqlite3

app = Flask(__name__)

DATABASE_PATH = 'database/warehouse.db'

def get_db_connection():
    conn = sqlite3.connect(DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    return conn

@app.route('/')
def home():
    conn = get_db_connection()
    products = conn.execute('SELECT * FROM MATERIAL').fetchall()
    total_materials = len(products)
    low_stock_count = sum(1 for row in products if int(row['quantity'] if 'quantity' in row.keys() else 0) < 50)
    today = datetime.now().strftime('%Y-%m-%d')
    # Tổng nhập hôm nay
    input_today = conn.execute("""
        SELECT COUNT(*) FROM STOCK_TRANSACTION st
        JOIN MATERIAL m ON st.material_id = m.material_id
        WHERE st.transaction_type = 'input' AND st.transaction_date LIKE ?
    """, (today+'%',)).fetchone()[0]
    # Tổng xuất hôm nay
    output_today = conn.execute("""
        SELECT COUNT(*) FROM STOCK_TRANSACTION st
        JOIN MATERIAL m ON st.material_id = m.material_id
        WHERE st.transaction_type = 'output' AND st.transaction_date LIKE ?
    """, (today+'%',)).fetchone()[0]
    # Hoạt động gần đây (5 giao dịch gần nhất)
    recent_activities = conn.execute("""
        SELECT st.*, m.material_name, m.part_code FROM STOCK_TRANSACTION st
        JOIN MATERIAL m ON st.material_id = m.material_id
        ORDER BY st.created_at DESC LIMIT 5
    """).fetchall()
    # Cảnh báo tồn kho (dưới 50)
    low_stock_alerts = [row for row in products if int(row['quantity'] if 'quantity' in row.keys() else 0) < 50][:5]
    conn.close()
    return render_template('home.html', 
        total_materials=total_materials, 
        low_stock_count=low_stock_count, 
        input_today=input_today, 
        output_today=output_today, 
        recent_activities=recent_activities, 
        low_stock_alerts=low_stock_alerts
    )

@app.route('/nhap-kho')
def nhap_kho():
    conn = get_db_connection()
    products = conn.execute('SELECT * FROM MATERIAL').fetchall()
    # Lấy lịch sử nhập kho từ STOCK_TRANSACTION
    history = conn.execute('''
        SELECT st.*, m.group_name, m.product_code, m.classification, m.part_code, m.material_name, m.specification, m.brand_name, m.unit, m.location, m.imported_by
        FROM STOCK_TRANSACTION st
        JOIN MATERIAL m ON st.material_id = m.material_id
        WHERE st.transaction_type = 'input'
        ORDER BY st.created_at DESC
        LIMIT 100
    ''').fetchall()
    conn.close()
    return render_template('nhap_kho.html', products=products, history=history)

@app.route('/xuat-kho')
def xuat_kho():
    conn = get_db_connection()
    products = conn.execute('SELECT * FROM MATERIAL').fetchall()
    conn.close()
    return render_template('xuat_kho.html', products=products)

@app.route('/bao-cao')
def bao_cao():
    conn = get_db_connection()
    products = conn.execute('SELECT * FROM MATERIAL').fetchall()
    report = []
    for idx, row in enumerate(products, 1):
        material_id = row['material_id']
        # Tổng nhập
        input_qty = conn.execute('SELECT COALESCE(SUM(quantity),0) FROM STOCK_TRANSACTION WHERE material_id=? AND transaction_type="input"', (material_id,)).fetchone()[0]
        # Tổng xuất
        output_qty = conn.execute('SELECT COALESCE(SUM(quantity),0) FROM STOCK_TRANSACTION WHERE material_id=? AND transaction_type="output"', (material_id,)).fetchone()[0]
        # Lần kiểm kê gần nhất
        inventory_row = conn.execute('SELECT quantity, created_at FROM STOCK_TRANSACTION WHERE material_id=? AND transaction_type="inventory" ORDER BY created_at DESC LIMIT 1', (material_id,)).fetchone()
        inventory_qty = inventory_row['quantity'] if inventory_row else ''
        inventory_time = inventory_row['created_at'] if inventory_row else ''
        # Tồn đầu kỳ: kiểm kê gần nhất trước lần nhập/xuất đầu tiên, hoặc 0 nếu chưa có
        opening_row = conn.execute('SELECT quantity FROM STOCK_TRANSACTION WHERE material_id=? AND transaction_type="inventory" ORDER BY created_at ASC LIMIT 1', (material_id,)).fetchone()
        opening_stock = opening_row['quantity'] if opening_row else 0
        # Tồn kho cuối kỳ = tồn đầu kỳ + nhập - xuất
        closing_stock = opening_stock + input_qty - output_qty
        report.append({
            'stt': idx,
            'group_use': row['group_name'] if 'group_name' in row.keys() else '',
            'product_code': row['product_code'] if 'product_code' in row.keys() else '',
            'classify': row['classification'] if 'classification' in row.keys() else '',
            'part_code': row['part_code'] if 'part_code' in row.keys() else '',
            'material_name': row['material_name'] if 'material_name' in row.keys() else '',
            'specification': row['specification'] if 'specification' in row.keys() else '',
            'brand': row['brand_name'] if 'brand_name' in row.keys() else '',
            'unit': row['unit'] if 'unit' in row.keys() else '',
            'opening_stock': opening_stock,
            'input': input_qty,
            'output': output_qty,
            'closing_stock': closing_stock,
            'inventory': inventory_qty,
            'location': row['location'] if 'location' in row.keys() else '',
            'last_update': row['updated_at'] if 'updated_at' in row.keys() else '',
            'last_time': row['created_at'] if 'created_at' in row.keys() else '',
        })
    conn.close()
    return render_template('bao_cao.html', report=report)

@app.route('/bao-cao-xls')
def bao_cao_xls():
    conn = get_db_connection()
    products = conn.execute('SELECT * FROM MATERIAL').fetchall()
    conn.close()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'BaoCaoTonKho'
    headers = ['STT','Mã code','Phân loại','Part code','Tên vật tư','Specification/Drawing','Thương hiệu','Nhà cung cấp','Đơn vị','MH & Loại','Công dụng','Vị trí','Số lượng','Mã QR','Thời gian']
    ws.append(headers)
    for idx, row in enumerate(products, 1):
        ws.append([
            idx,
            row['product_code'],
            row['classification'],
            row['part_code'],
            row['material_name'],
            row['specification'],
            row['brand_name'],
            '', # providers
            row['unit'],
            '', # mh_loai
            '', # cong_dung
            row['location'],
            row['quantity'] if 'quantity' in row.keys() else '',
            row['part_code'],
            row['updated_at']
        ])
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb.save(tmp.name)
    tmp.close()
    return send_file(tmp.name, as_attachment=True, download_name=f"bao_cao_ton_kho_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

@app.route('/danh-sach')
def danh_sach():
    conn = get_db_connection()
    products = conn.execute('SELECT * FROM MATERIAL').fetchall()
    xuat_kho = conn.execute("""
        SELECT st.*, m.material_name, m.part_code FROM STOCK_TRANSACTION st
        JOIN MATERIAL m ON st.material_id = m.material_id
        WHERE st.transaction_type = 'output'
        ORDER BY st.created_at DESC
    """).fetchall()
    conn.close()
    return render_template('danh_sach.html', products=products, xuat_kho=xuat_kho)

@app.route('/xuat-kho-submit', methods=['POST'])
def xuat_kho_submit():
    data = request.get_json()
    qr_code = data.get('qr_code')
    quantity = data.get('quantity')
    if not qr_code or not quantity:
        return jsonify({'status': 'error', 'message': 'Thiếu dữ liệu'})
    try:
        quantity = int(quantity)
    except Exception:
        return jsonify({'status': 'error', 'message': 'Số lượng không hợp lệ'})
    conn = get_db_connection()
    material = conn.execute('SELECT * FROM MATERIAL WHERE part_code=?', (qr_code,)).fetchone()
    if not material:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Không tìm thấy sản phẩm'})
    current_qty = int(material['quantity'] if 'quantity' in material.keys() else 0)
    if current_qty < quantity:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Tồn kho không đủ'})
    new_qty = current_qty - quantity
    conn.execute('UPDATE MATERIAL SET quantity=?, updated_at=? WHERE material_id=?', (new_qty, datetime.now(), material['material_id']))
    conn.execute('''INSERT INTO STOCK_TRANSACTION (material_id, transaction_type, quantity, transaction_date, reference_number, notes, created_at) VALUES (?, 'output', ?, ?, '', '', ?)''',
        (material['material_id'], quantity, datetime.now().date(), datetime.now()))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

@app.route('/xuat-kho-batch', methods=['POST'])
def xuat_kho_batch():
    data = request.get_json()
    items = data.get('items', [])
    if not items:
        return jsonify({'status': 'error', 'message': 'Không có sản phẩm để xuất kho'})
    conn = get_db_connection()
    for item in items:
        qr_code = item.get('qr_code')
        quantity = int(item.get('quantity', 0))
        material = conn.execute('SELECT * FROM MATERIAL WHERE part_code=?', (qr_code,)).fetchone()
        if not material:
            conn.close()
            return jsonify({'status': 'error', 'message': f'Không tìm thấy sản phẩm với mã QR: {qr_code}'})
        current_qty = int(material['quantity'] if 'quantity' in material.keys() else 0)
        if current_qty < quantity:
            conn.close()
            return jsonify({'status': 'error', 'message': f'Tồn kho không đủ cho mã QR: {qr_code}'})
        new_qty = current_qty - quantity
        conn.execute('UPDATE MATERIAL SET quantity=?, updated_at=? WHERE material_id=?', (new_qty, datetime.now(), material['material_id']))
        conn.execute('''INSERT INTO STOCK_TRANSACTION (material_id, transaction_type, quantity, transaction_date, reference_number, notes, created_at) VALUES (?, 'output', ?, ?, '', '', ?)''',
            (material['material_id'], quantity, datetime.now().date(), datetime.now()))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

@app.route('/delete-product', methods=['POST'])
def delete_product():
    data = request.get_json()
    qr_code = data.get('qr_code')
    if not qr_code:
        return jsonify({'status': 'error', 'message': 'Thiếu mã QR'})
    conn = get_db_connection()
    material = conn.execute('SELECT * FROM MATERIAL WHERE part_code=?', (qr_code,)).fetchone()
    if not material:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Không tìm thấy sản phẩm để xóa'})
    conn.execute('DELETE FROM MATERIAL WHERE material_id=?', (material['material_id'],))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

@app.route('/api/product-info/<qr_code>')
def api_product_info(qr_code):
    if not qr_code:
        return jsonify({'status': 'error', 'message': 'Thiếu mã QR'})
    conn = get_db_connection()
    material = conn.execute('SELECT * FROM MATERIAL WHERE part_code=?', (qr_code,)).fetchone()
    conn.close()
    if material:
        return jsonify({'status': 'success', 'product': dict(material)})
    return jsonify({'status': 'error', 'message': 'Không tìm thấy sản phẩm'})

@app.route('/kiem-ke', methods=['GET', 'POST'])
def kiem_ke():
    if request.method == 'GET':
        return render_template('kiem_ke.html')
    data = request.get_json()
    qr_code = data.get('qr_code')
    quantity = data.get('inventory')
    if not qr_code or quantity is None:
        return jsonify({'status': 'error', 'message': 'Thiếu mã QR hoặc số lượng kiểm kê'})
    try:
        quantity = int(quantity)
    except Exception:
        return jsonify({'status': 'error', 'message': 'Số lượng kiểm kê không hợp lệ'})
    conn = get_db_connection()
    material = conn.execute('SELECT * FROM MATERIAL WHERE part_code=?', (qr_code,)).fetchone()
    if not material:
        conn.close()
        return jsonify({'status': 'error', 'message': 'Không tìm thấy sản phẩm'})
    conn.execute('UPDATE MATERIAL SET quantity=?, updated_at=? WHERE material_id=?', (quantity, datetime.now(), material['material_id']))
    # Ghi lịch sử kiểm kê vào bảng riêng (nếu có)
    conn.execute('''INSERT INTO STOCK_TRANSACTION (material_id, transaction_type, quantity, transaction_date, reference_number, notes, created_at) VALUES (?, 'inventory', ?, ?, '', '', ?)''',
        (material['material_id'], quantity, datetime.now().date(), datetime.now()))
    conn.commit()
    conn.close()
    return jsonify({'status': 'success'})

@app.route('/kiem-ke-lich-su')
def kiem_ke_lich_su():
    conn = get_db_connection()
    history = conn.execute("""
        SELECT st.*, m.material_name, m.part_code FROM STOCK_TRANSACTION st
        JOIN MATERIAL m ON st.material_id = m.material_id
        WHERE st.transaction_type = 'inventory'
        ORDER BY st.created_at DESC
    """).fetchall()
    conn.close()
    return render_template('kiem_ke_lich_su.html', history=history)

@app.route('/create-product', methods=['POST'])
def create_product():
    data = request.get_json()
    group_use = data.get('group_use')
    product_code = data.get('product_code')
    classify = data.get('classify')
    part_code = data.get('part_code')
    material_name = data.get('material_name')
    specification = data.get('specification')
    brand = data.get('brand')
    unit = data.get('unit')
    location = data.get('location')
    imported_by = data.get('imported_by')
    quantity = data.get('quantity')
    if not all([group_use, product_code, classify, part_code, material_name, specification, brand, unit, location, imported_by, quantity]):
        return jsonify({'status': 'error', 'message': 'Thiếu thông tin'})
    try:
        quantity = int(quantity)
    except Exception:
        return jsonify({'status': 'error', 'message': 'Số lượng không hợp lệ'})
    conn = get_db_connection()
    now = datetime.now()
    # Luôn thêm mới sản phẩm, không kiểm tra trùng part_code
    conn.execute('''INSERT INTO MATERIAL (group_name, product_code, classification, part_code, material_name, specification, brand_name, unit, location, imported_by, quantity, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
        (group_use, product_code, classify, part_code, material_name, specification, brand, unit, location, imported_by, quantity, now, now))
    material = conn.execute('SELECT material_id FROM MATERIAL WHERE part_code=? AND material_name=? AND created_at=?', (part_code, material_name, now)).fetchone()
    material_id = material['material_id'] if material else None
    if material_id:
        conn.execute('''INSERT INTO STOCK_TRANSACTION (material_id, transaction_type, quantity, transaction_date, reference_number, notes, created_at) VALUES (?, 'input', ?, ?, '', '', ?)''',
            (material_id, quantity, now.date(), now))
    conn.commit()
    conn.close()
    # Sinh mã QR duy nhất cho lần nhập này, tên file: {part_code}_{uuid}.png
    import uuid
    qr_uuid = str(uuid.uuid4())
    qr_img_path = f'static/qr_codes/{part_code}_{qr_uuid}.png'
    qr_dir = os.path.dirname(qr_img_path)
    if not os.path.exists(qr_dir):
        os.makedirs(qr_dir, exist_ok=True)
    qr = qrcode.QRCode(version=1, box_size=10, border=2)
    qr.add_data(part_code)
    qr.make(fit=True)
    img = qr.make_image(fill='black', back_color='white')
    img.save(qr_img_path)
    qr_url = f'/static/qr_codes/{part_code}_{qr_uuid}.png'
    return jsonify({'status': 'success', 'qr_url': qr_url})

@app.route('/download-db')
def download_db():
    db_path = os.path.join(os.path.dirname(__file__), 'database', 'warehouse.db')
    return send_file(db_path, as_attachment=True, download_name='warehouse.db')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
