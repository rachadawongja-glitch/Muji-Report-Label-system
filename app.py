import os
from flask import Flask, request, jsonify
import pandas as pd

app = Flask(__name__)

# กำหนดโฟลเดอร์สำหรับเซฟไฟล์ข้อมูล (ล้อตามดิสก์ /data บน Render)
UPLOAD_FOLDER = 'data'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MUJI Label Report System</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@latest/tabler-icons.min.css">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        body { display: flex; background: #f8fafc; color: #1e293b; min-height: 100vh; }
        .sidebar { width: 240px; background: #1e293b; color: #fff; padding: 24px 16px; position: fixed; height: 100vh; }
        .sidebar-logo { margin-bottom: 32px; padding-left: 8px; }
        .sidebar-logo h1 { font-size: 18px; font-weight: 700; color: #f8fafc; }
        .sidebar-logo p { font-size: 12px; color: #94a3b8; margin-top: 4px; }
        .nav { display: flex; flex-direction: column; gap: 8px; }
        .nav-item { display: flex; align-items: center; gap: 12px; width: 100%; padding: 12px; background: none; border: none; border-left: 4px solid transparent; color: #94a3b8; font-size: 14px; font-weight: 500; text-align: left; cursor: pointer; transition: all 0.2s; border-radius: 0 6px 6px 0; }
        .nav-item:hover { background: rgba(255, 255, 255, .05); color: #fff; }
        .nav-item.active { background: rgba(255, 255, 255, .1); color: #fff; border-left-color: #ef4444; font-weight: 600; }
        .nav-item i { font-size: 18px; }
        .main { margin-left: 240px; flex: 1; padding: 32px; }
        .topbar { margin-bottom: 24px; padding-bottom: 16px; border-bottom: 1px solid #e2e8f0; }
        .topbar h2 { font-size: 22px; font-weight: 600; color: #0f172a; }
        .page { display: none; }
        .page.active { display: block; }
        .card { background: #fff; border-radius: 12px; padding: 24px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border: 1px solid #e2e8f0; max-width: 600px; }
        .file-box { border: 2px dashed #cbd5e1; border-radius: 8px; padding: 32px; text-align: center; margin: 20px 0; background: #f8fafc; cursor: pointer; }
        .file-box:hover { border-color: #94a3b8; }
        .btn-import { background: #ef4444; color: #fff; border: none; padding: 12px 24px; border-radius: 6px; font-size: 15px; font-weight: 600; cursor: pointer; width: 100%; transition: background 0.2s; }
        .btn-import:hover { background: #dc2626; }
        .status-box { margin-top: 20px; padding: 16px; border-radius: 8px; display: none; font-size: 14px; line-height: 1.5; }
        .status-loading { background: #fffbeb; color: #b45309; border: 1px solid #fde68a; }
        .status-success { background: #f0fdf4; color: #15803d; border: 1px solid #bbf7d0; }
        .status-error { background: #fef2f2; color: #b91c1c; border: 1px solid #fca5a5; }
    </style>
</head>
<body>
    <div class="sidebar">
        <div class="sidebar-logo">
            <h1>MUJI Label System</h1>
            <p>ระบบรายงานป้ายสินค้า</p>
        </div>
        <div class="nav">
            <button class="nav-item active" onclick="switchPage('home')"><i class="ti ti-home"></i> หน้าแรก</button>
            <button class="nav-item" onclick="switchPage('report')"><i class="ti ti-file-text"></i> รายงาน</button>
            <button class="nav-item" onclick="switchPage('import')"><i class="ti ti-upload"></i> นำเข้าข้อมูล</button>
        </div>
    </div>
    <div class="main">
        <div id="page-home" class="page active">
            <div class="topbar"><h2>ยินดีต้อนรับสู่ระบบ MUJI Label Report</h2></div>
            <div class="card">
                <p style="color: #64748b;">ระบบกำลังทำงานปกติ หากใช้งานระบบนี้เป็นครั้งแรก กรุณาเลือกเมนู <b>"นำเข้าข้อมูล"</b> ทางด้านซ้าย เพื่ออัปโหลดข้อมูลภาพสินค้าเข้าสู่ฐานข้อมูล</p>
            </div>
        </div>
        <div id="page-report" class="page">
            <div class="topbar"><h2>รายงานป้ายสินค้า</h2></div>
            <div class="card">
                <p style="color: #64748b;">ขณะนี้ยังไม่มีข้อมูลรายงานแสดงผล กรุณานำเข้าข้อมูลก่อนใช้งาน</p>
            </div>
        </div>
        <div id="page-import" class="page">
            <div class="topbar"><h2>นำเข้าข้อมูลสินค้าครั้งแรก</h2></div>
            <div class="card">
                <p style="font-size: 14px; color: #64748b; margin-bottom: 10px;">เมื่อระบบ online แล้ว ให้ import รูปภาพสินค้าเข้า database ครั้งเดียว</p>
                <div class="file-box" onclick="document.getElementById('excelFile').click()">
                    <i class="ti ti-file-spreadsheet" style="font-size: 40px; color: #94a3b8; display:block; margin-bottom: 8px;"></i>
                    <span id="fileNameDisplay" style="font-weight: 500; color: #475569;">คลิกเพื่อเลือกไฟล์ Picture_List.xlsx</span>
                    <input type="file" id="excelFile" accept=".xlsx" style="display: none;" onchange="updateFileName(this)">
                </div>
                <button class="btn-import" onclick="startImport()">เริ่ม Import</button>
                
                <div id="statusMessage" class="status-box status-loading">
                    ⏳ <b>ระบบกำลังนำเข้าข้อมูลและประมวลผลไฟล์...</b> กรุณารอสายสักครู่และห้ามปิดหน้านี้
                </div>
                <div id="successMessage" class="status-box status-success"></div>
                <div id="errorMessage" class="status-box status-error"></div>
            </div>
        </div>
    </div>
    <script>
        function switchPage(pageId) {
            document.querySelectorAll('.page').forEach(page => page.classList.remove('active'));
            document.querySelectorAll('.nav-item').forEach(btn => btn.classList.remove('active'));
            document.getElementById('page-' + pageId).classList.add('active');
            if(event) event.currentTarget.classList.add('active');
        }
        function updateFileName(input) {
            const display = document.getElementById('fileNameDisplay');
            if (input.files && input.files[0]) {
                display.innerText = input.files[0].name;
                display.style.color = "#1e293b";
            }
        }
        // ฟังก์ชันยิง API อัปโหลดไฟล์ไปที่ฝั่ง Python จริง
        function startImport() {
            const fileInput = document.getElementById('excelFile');
            if (!fileInput.files || fileInput.files.length === 0) {
                alert('กรุณาเลือกไฟล์ Picture_List.xlsx ก่อนคลิกเริ่ม Import ครับ');
                return;
            }
            
            const formData = new FormData();
            formData.append('excel_file', fileInput.files[0]);

            document.getElementById('statusMessage').style.display = 'block';
            document.getElementById('successMessage').style.display = 'none';
            document.getElementById('errorMessage').style.display = 'none';

            fetch('/api/import', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                document.getElementById('statusMessage').style.display = 'none';
                if (data.status === 'success') {
                    const successBox = document.getElementById('successMessage');
                    successBox.innerHTML = `✅ <b>นำเข้าข้อมูลสำเร็จ!</b> อ่านได้ทั้งหมด ${data.row_count} แถว เรียบร้อยแล้ว`;
                    successBox.style.display = 'block';
                } else {
                    const errorBox = document.getElementById('errorMessage');
                    errorBox.innerHTML = `❌ <b>เกิดข้อผิดพลาด:</b> ${data.message}`;
                    errorBox.style.display = 'block';
                }
            })
            .catch(error => {
                document.getElementById('statusMessage').style.display = 'none';
                const errorBox = document.getElementById('errorMessage');
                errorBox.innerHTML = `❌ <b>ไม่สามารถเชื่อมต่อเซิร์ฟเวอร์ได้:</b> ${error}`;
                errorBox.style.display = 'block';
            });
        }
    </script>
</body>
</html>
"""

@app.route('/')
def home():
    return HTML_TEMPLATE

# API สำหรับรับไฟล์ อัปโหลด และอ่านข้อมูลจาก Excel
@app.route('/api/import', methods=['POST'])
def import_data():
    if 'excel_file' not in request.files:
        return jsonify({'status': 'error', 'message': 'ไม่พบไฟล์ที่ส่งมา'}), 400
        
    file = request.files['excel_file']
    if file.filename == '':
        return jsonify({'status': 'error', 'message': 'ชื่อไฟล์ว่างเปล่า'}), 400

    if file and file.filename.endswith('.xlsx'):
        try:
            # 1. บันทึกไฟล์ลง Disk ชั่วคราวในโฟลเดอร์ data
            file_path = os.path.join(UPLOAD_FOLDER, 'Picture_List.xlsx')
            file.save(file_path)
            
            # 2. ใช้ pandas อ่านไฟล์ Excel
            df = pd.read_excel(file_path)
            row_count = len(df)
            
            # ตรงนี้คุณสามารถเพิ่ม Logic การ Loop ข้อมูลไปใช้งานต่อได้ เช่น:
            # for index, row in df.iterrows():
            #     print(row['ชื่อคอลัมน์'])  # ตัวอย่างการอ่านค่าคอลัมน์ในแต่ละแถว

            return jsonify({
                'status': 'success',
                'message': 'อัปโหลดและอ่านข้อมูลสำเร็จ',
                'row_count': row_count
            })
            
        except Exception as e:
            return jsonify({'status': 'error', 'message': f'เกิดข้อผิดพลาดขณะอ่านไฟล์: {str(e)}'}), 500
    else:
        return jsonify({'status': 'error', 'message': 'กรุณาอัปโหลดเฉพาะไฟล์นามสกุล .xlsx เท่านั้น'}), 400

if __name__ == '__main__':
    app.run(debug=True)
