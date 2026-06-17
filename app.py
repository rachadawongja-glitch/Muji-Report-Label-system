from flask import Flask

app = Flask(__name__)

# เก็บโค้ดหน้าตาเว็บ (HTML/CSS) ไว้ในตัวแปร HTML_TEMPLATE
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="th">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MUJI Label Report System</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@3.14.0/dist/tabler-icons.min.css">
    <style>
        :root {
            --bg: #f8f9fa;
            --card: #fff;
            --border: #e5e7eb;
            --text: #111827;
            --muted: #6b7280;
            --blue: #1d4ed8;
            --blue-light: #eff6ff;
            --blue-mid: #bfdbfe;
            --green: #15803d;
            --green-light: #f0fdf4;
            --green-mid: #bbf7d0;
            --red: #b91c1c;
            --red-light: #fef2f2;
            --amber: #b45309;
            --amber-light: #fffbeb;
            --hdr: #1A56A0;
            --row-alt: #EFF4FB;
        }
        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }
        body {
            font-family: 'Sarabun', 'Segoe UI', sans-serif;
            background: var(--bg);
            color: var(--text);
            font-size: 14px;
        }
        a {
            color: var(--blue);
            text-decoration: none;
        }
        /* Layout */
        .sidebar {
            position: fixed;
            top: 0;
            left: 0;
            width: 220px;
            height: 100vh;
            background: var(--hdr);
            display: flex;
            flex-direction: column;
            padding: 0;
            z-index: 100;
        }
        .sidebar-logo {
            padding: 20px 18px 16px;
            border-bottom: 1px solid rgba(255, 255, 255, .15);
        }
        .sidebar-logo h1 {
            color: #fff;
            font-size: 15px;
            font-weight: 600;
            line-height: 1.3;
        }
        .sidebar-logo p {
            color: rgba(255, 255, 255, .6);
            font-size: 11px;
            margin-top: 3px;
        }
        .nav {
            flex: 1;
            padding: 12px 0;
            overflow-y: auto;
        }
        .nav-item {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 10px 18px;
            color: rgba(255, 255, 255, .75);
            cursor: pointer;
            border: none;
            background: none;
            width: 100%;
            font-size: 13px;
            font-family: inherit;
            transition: all .15s;
            border-left: 3px solid transparent;
        }
        .nav-item:hover {
            background: rgba(255, 255, 255, .08);
            color: #fff;
        }
        .nav-item.active {
            background: rgba(255, 255, 255, .12);
            color: #fff;
            border-left-color: #fff;
        }
        .nav-item i {
            font-size: 17px;
            width: 20px;
        }
        .main {
            margin-left: 220px;
            min-height: 100vh;
            padding: 24px;
        }
        .topbar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 22px;
        }
        .topbar h2 {
            font-size: 18px;
            font-weight: 600;
        }
        .page {
            display: none;
        }
        .page.active {
            display: block;
        }
    </style>
</head>
<body>

    <div class="sidebar">
        <div class="sidebar-logo">
            <h1>MUJI Label System</h1>
            <p>ระบบรายงานป้ายสินค้า</p>
        </div>
        <div class="nav">
            <button class="nav-item active"><i class="ti ti-home"></i> หน้าแรก</button>
            <button class="nav-item"><i class="ti ti-file-text"></i> รายงาน</button>
        </div>
    </div>

    <div class="main">
        <div class="topbar">
            <h2>ยินดีต้อนรับสู่ระบบ MUJI Label Report</h2>
        </div>
        <div class="page active">
             <p>ระบบกำลังทำงานปกติ</p>
        </div>
    </div>

</body>
</html>
"""

@app.route('/')
def home():
    # ส่งโค้ด HTML ด้านบนออกไปแสดงผลที่หน้าเว็บโดยตรง
    return HTML_TEMPLATE

if __name__ == '__main__':
    app.run(debug=True)
