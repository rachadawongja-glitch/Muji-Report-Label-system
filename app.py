from flask import Flask

app = Flask(__name__)

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
        
        /* Sidebar Styles */
        .sidebar { width: 240px; background: #1e293b; color: #fff; padding: 24px 16px; position: fixed; height: 100vh; }
        .sidebar-logo { margin-bottom: 32px; padding-left: 8px; }
        .sidebar-logo h1 { font-size: 18px; font-weight: 700; color: #f8fafc; }
        .sidebar-logo p { font-size: 12px; color: #94a3b8; margin-top: 4px; }
        .nav { display: flex; flex-direction: column; gap: 8px; }
        .nav-item { display: flex; align-items: center; gap: 12px; width: 100%; padding: 12px; background: none; border: none; border-left: 4px solid transparent; color: #94a3b8; font-size: 14px; font-weight: 500; text-align: left; cursor: pointer; transition: all 0.2s; border-radius: 0 6px 6px 0; }
        .nav-item:hover { background: rgba(255, 255, 255, .05); color: #fff; }
        .nav-item.active { background: rgba(255, 255, 255, .1); color: #fff; border-left-color: #ef4444; font-weight: 600; }
        .nav-item i { font-size: 18px; }
        
        /* Main Content Styles */
        .main { margin-left: 240px; flex: 1; padding: 32px; }
        .topbar { margin-bottom: 24px; padding-bottom: 16px; border-bottom: 1px solid #e2e8f0; }
        .topbar h2 { font-size: 22px; font-weight: 600; color: #0f172a; }
        
        /* Page Control */
        .page { display: none; }
        .page.active { display: block; }
        
        /* Import UI Styles */
        .card { background: #fff; border-radius: 12px; padding: 24px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border: 1px solid #e2e8f0; max-width: 600px; }
        .file-box { border: 2px dashed #cbd5e1; border-radius: 8px; padding: 32px; text-align: center; margin: 20px 0; background: #f8fafc; cursor: pointer; }
        .file-box:hover { border-color: #94a3b8; }
        .btn-import { background: #ef4444; color: #fff; border: none; padding: 12px 24px; border-radius: 6px; font-size: 15px; font-weight: 600; cursor: pointer; width: 100%; transition: background 0.2s; }
        .btn-import:hover { background: #dc2626; }
        
        /* Status Message Box */
        .status-box { margin-top: 20px; padding: 16px; border-radius: 8px; display: none; font-size: 14px; line-height: 1.5; }
        .status-loading { background: #fffbeb; color: #b45309; border: 1px solid #fde68a; }
        .status-success { background: #f0fdf4; color: #15803d; border: 1px solid #bbf7d0; }
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
            <div class="topbar">
                <h2>ยินดีต้อนรับสู่ระบบ MUJI Label Report</h2>
            </div>
            <div class="card">
                <p style="color: #64748b;">ระบบกำลังทำงานปกติ หากใช้งานระบบนี้เป็นครั้งแรก กรุณาเลือกเมนู <b>"นำเข้าข้อมูล"</b> ทางด้านซ้าย เพื่ออัปโหลดข้อมูลภาพสินค้าเข้าสู่ฐานข้อมูล</p>
            </div>
        </div>

        <div id="page-report" class="page">
            <div class="topbar">
                <h2>รายงานป้ายสินค้า</h2>
            </div>
            <div class="card">
                <p style="color: #64748b;">ขณะนี้ยังไม่มีข้อมูลรายงานแสดงผล กรุณานำเข้าข้อมูลก่อนใช้งาน</p>
            </div>
        </div>

        <div id="page-import" class="page">
            <div class="topbar">
                <h2>นำเข้าข้อมูลสินค้าครั้งแรก</h2>
            </div>
            <div class="card">
                <p style="font-size: 14px; color: #64748b; margin-bottom: 10px;">เมื่อระบบ online แล้ว ให้ import รูปภาพสินค้าเข้า database ครั้งเดียว</p>
                
                <div class="file-box" onclick="document.getElementById('excelFile').click()">
                    <i class="ti ti-file-spreadsheet" style="font-size: 40px; color: #94a3b8; display:block; margin-bottom: 8px;"></i>
                    <span id="fileNameDisplay" style="font-weight: 500; color: #475569;">คลิกเพื่อเลือกไฟล์ Picture_List.xlsx</span>
                    <input type="file" id="excelFile" accept=".xlsx" style="display: none;" onchange="updateFileName(this)">
                </div>

                <button class="btn-import" onclick="startImport()">เริ่ม Import</button>

                <div id="statusMessage" class="status-box status-loading">
                    ⏳ <b>ระบบกำลังนำเข้าข้อมูล...</b> กรุณารอประมาณ 2-3 นาที และห้ามปิดหรือโหลดหน้าจอนี้ใหม่เด็ดขาด ระบบกำลังเชื่อมต่อ Database เพื่อบันทึกรูปภาพสินค้า
                </div>
                
                <div id="successMessage" class="status-box status-success">
                    ✅ <b>นำเข้าข้อมูลสำเร็จ!</b> ระบบทำการบันทึกข้อมูลและรูปภาพสินค้าเรียบร้อยแล้ว คุณสามารถเริ่มดูเมนูรายงานได้ทันที
                </div>
            </div>
        </div>
    </div>

    <script>
        // ฟังก์ชันสำหรับสลับหน้าจอเมื่อกดเมนู
        function switchPage(pageId) {
            // ซ่อนทุกหน้าจอ
            document.querySelectorAll('.page').forEach(page => page.classList.remove('active'));
            // เอาสถานะ active ออกจากปุ่มทั้งหมด
            document.querySelectorAll('.nav-item').forEach(btn => btn.classList.remove('active'));
            
            // เปิดหน้าจอและปุ่มที่ถูกเลือก
            document.getElementById('page-' + pageId).classList.add('active');
            event.currentTarget.classList.add('active');
        }

        // แสดงชื่อไฟล์เมื่อเลือกเสร็จ
        function updateFileName(input) {
            const display = document.getElementById('fileNameDisplay');
            if (input.files && input.files[0]) {
                display.innerText = input.files[0].name;
                display.style.color = "#1e293b";
            }
        }

        // ฟังก์ชันทำงานเมื่อกดปุ่ม เริ่ม Import
        function startImport() {
            const fileInput = document.getElementById('excelFile');
            if (!fileInput.files || fileInput.files.length === 0) {
                alert('กรุณาเลือกไฟล์ Picture_List.xlsx ก่อนคลิกเริ่ม Import ครับ');
                return;
            }

            // แสดงสถานะกำลังบันทึก (กล่องสีเหลืองแจงรอ 2-3 นาที)
            document.getElementById('statusMessage').style.display = 'block';
            document.getElementById('successMessage').style.display = 'none';

            // จำลองการโหลดจริง 5 วินาทีเพื่อให้เห็นผลลัพธ์ (สามารถปรับเพิ่มหรือผูก API จริงได้ภายหลัง)
            setTimeout(() => {
                document.getElementById('statusMessage').style.display = 'none';
                document.getElementById('successMessage').style.display = 'block';
            }, 5000);
        }
    </script>
</body>
</html>
"""

@app.route('/')
def home():
    return HTML_TEMPLATE

if __name__ == '__main__':
    app.run(debug=True)
