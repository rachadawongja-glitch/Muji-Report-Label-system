/**
 * generate_docx.js
 * รับ JSON payload จาก Python → สร้าง .docx พร้อมรูปภาพ 3 ภาพต่อรายการ
 * ใช้ docx-js library
 *
 * Usage: node generate_docx.js payload.json output.docx
 */

const fs   = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, AlignmentType, WidthType, BorderStyle, ShadingType,
  VerticalAlign, PageOrientation,
} = require('./node_modules/docx');

// ── Load payload ────────────────────────────────────────────────────────────
const [,, payloadPath, outPath] = process.argv;
if (!payloadPath || !outPath) {
  console.error("Usage: node generate_docx.js payload.json output.docx");
  process.exit(1);
}
const data = JSON.parse(fs.readFileSync(payloadPath, "utf-8"));
const { meta, items } = data;

// ── Helpers ─────────────────────────────────────────────────────────────────
const FONT = "TH Sarabun New";
const pt   = (n) => n * 2; // half-points

const thBorder = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const borders  = { top: thBorder, bottom: thBorder, left: thBorder, right: thBorder };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders= { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function run(text, opts = {}) {
  return new TextRun({
    text,
    font: FONT,
    size: pt(opts.size || 14),
    bold: opts.bold || false,
    color: opts.color || "000000",
    ...opts,
  });
}

function para(children, opts = {}) {
  return new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing:   { before: opts.sb || 0, after: opts.sa || 60 },
    indent:    opts.indent ? { firstLine: opts.indent } : undefined,
    children:  Array.isArray(children) ? children : [children],
  });
}

function headerCell(text, opts = {}) {
  return new TableCell({
    borders,
    width:   { size: opts.w || 1500, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    shading: { fill: "1A56A0", type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    children: [para(run(text, { bold: true, color: "FFFFFF", size: 13 }),
                    { align: AlignmentType.CENTER })],
  });
}

function dataCell(content, opts = {}) {
  return new TableCell({
    borders,
    width:   { size: opts.w || 1500, type: WidthType.DXA },
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    shading: { fill: opts.bg || "FFFFFF", type: ShadingType.CLEAR },
    verticalAlign: VerticalAlign.CENTER,
    children: Array.isArray(content) ? content : [content],
  });
}

function imgCell(base64Data, w, opts = {}) {
  if (!base64Data) {
    return dataCell(para(run("—", { color: "AAAAAA", size: 11 }),
                         { align: AlignmentType.CENTER }), opts);
  }
  const buf  = Buffer.from(base64Data, "base64");
  const imgRun = new ImageRun({
    type: "png",
    data: buf,
    transformation: { width: 75, height: 62 },
    altText: { title: "Product image", description: "Product image", name: "img" },
  });
  return dataCell(
    para([imgRun], { align: AlignmentType.CENTER }),
    opts
  );
}

// ── Thai date ────────────────────────────────────────────────────────────────
const TH_MONTHS = ["","มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                   "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];
function thDate(str) {
  const m = (str || "").match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${parseInt(m[1])} ${TH_MONTHS[parseInt(m[2])]} พ.ศ. ${m[3]}`;
  return str || "";
}
const today = new Date();
const todayStr = `วันที่ ${today.getDate()} ${TH_MONTHS[today.getMonth()+1]} พ.ศ. ${today.getFullYear()+543}`;

// ── Column widths (DXA). A4 content = 11906 - 1800 - 1800 = 8306 ────────────
// cols: no(700) | reg(2600) | name(3400) | qty(800) | img1(950) | img2(950) | img3(950) = 10350 → adjust
// Use: 600 | 2300 | 3300 | 700 | 900 | 900 | 900 = 9600 ≈ A4 landscape-margin-safe
const COL_W = [600, 2200, 3400, 700, 900, 900, 900];
const TBL_W  = COL_W.reduce((a, b) => a + b, 0); // 9600

// ── Build document ──────────────────────────────────────────────────────────
const children = [];

// Date
children.push(para(run(todayStr, { size: 14 }), { align: AlignmentType.RIGHT, sa: 40 }));

// เรื่อง / เรียน
children.push(para([
  run("เรื่อง  ", { bold: true, size: 14 }),
  run("การจัดทำรายงานชี้แจงการติดฉลากสินค้า ตามประกาศกระทรวงสาธารณสุข ฉบับที่ ๔๕๐", { size: 14 }),
], { sa: 30 }));

children.push(para([
  run("เรียน  ", { bold: true, size: 14 }),
  run("ผู้อำนวยการด่านอาหารและยา", { size: 14 }),
], { sa: 30 }));

// สิ่งที่แนบ
children.push(para([
  run("สิ่งที่แนบมาด้วย  1. ใบขนสินค้าขาเข้าเลขที่ ", { bold: true, size: 14 }),
  run(meta.customs_no, { size: 14 }),
], { sa: 10 }));
children.push(para([
  run("                              2. Invoice no. ", { bold: true, size: 14 }),
  run(meta.invoice_no, { size: 14 }),
], { sa: 10 }));
children.push(para(run("                              3. รายงานผลการชี้แจงการติดฉลากสินค้า", { bold: true, size: 14 }), { sa: 40 }));

// Body paragraph 1
children.push(para(
  run(
    `\t\tบริษัท มูจิ รีเทล (ประเทศไทย) จำกัด ได้นำเข้าสินค้าที่เป็นอาหาร ` +
    `จำนวน ${meta.total_cartons} CARTONS โดยเรือ ${meta.ship_name} ` +
    `ตามใบขนสินค้าขาเข้าเลขที่ ${meta.customs_no} ` +
    `ตามเอกสารใบแจ้งหนี้เลขที่ ${meta.invoice_no} ` +
    `นำเข้ามาจำนวนทั้งหมด ${meta.total_items} รายการ ` +
    `ซึ่งได้รับการตรวจปล่อยวันที่ ${meta.import_date} สถานที่ที่ได้รับการตรวจปล่อยท่าเรือกรุงเทพ สทก.`,
    { size: 14 }
  ), { sa: 30 }
));

// Body paragraph 2
children.push(para(
  run(
    `\t\tโดยทางบริษัทฯ ใช้สถานที่จัดเก็บสินค้า CENTRAL DEPARTMENT STORE LTD (WAREHOUSE) ` +
    `ที่อยู่ 105-106 หมู่ที่ 1 ตำบล ศีรษะจรเข้ใหญ่ อำเภอ บางเสาธง จังหวัด สมุทรปราการ 10540`,
    { size: 14 }
  ), { sa: 30 }
));

// Body paragraph 3
children.push(para(
  run(
    `\t\tทั้งนี้ทางบริษัทฯ จึงได้จัดทำรายงานการชี้แจงการติดฉลากของสินค้า ` +
    `ตามประกาศกระทรวงสาธารณสุข ฉบับที่ ๔๕๐ โดยมีรายละเอียดตามเลขที่ใบแจ้งหนี้ ` +
    `${meta.invoice_no} ดังนี้`,
    { size: 14 }
  ), { sa: 60 }
));

// ── Table ─────────────────────────────────────────────────────────────────────
const tableRows = [
  // Header row
  new TableRow({
    tableHeader: true,
    children: [
      headerCell("รายการที่",                   { w: COL_W[0] }),
      headerCell("เลขทะเบียน/เลขสารบบ",        { w: COL_W[1] }),
      headerCell("ชื่อสินค้า",                  { w: COL_W[2] }),
      headerCell("จำนวน\nที่นำเข้า",           { w: COL_W[3] }),
      headerCell("รูปภาพ 1",                    { w: COL_W[4] }),
      headerCell("รูปภาพ 2",                    { w: COL_W[5] }),
      headerCell("รูปภาพ 3",                    { w: COL_W[6] }),
    ],
  }),
];

// Data rows
items.forEach((item, idx) => {
  const bg = idx % 2 === 0 ? "FFFFFF" : "EFF4FB";
  tableRows.push(
    new TableRow({
      height: { value: 1440, rule: "atLeast" }, // ~1 inch min
      children: [
        // รายการที่
        dataCell(para(run(String(item.no), { size: 12 }), { align: AlignmentType.CENTER }), { w: COL_W[0], bg }),
        // เลขทะเบียน
        dataCell(para(run(item.reg || "—", { size: 10, color: "555555" })), { w: COL_W[1], bg }),
        // ชื่อสินค้า
        dataCell(para(run(item.name, { size: 12 })), { w: COL_W[2], bg }),
        // จำนวน
        dataCell(para(run(String(item.qty), { size: 12 }), { align: AlignmentType.CENTER }), { w: COL_W[3], bg }),
        // รูป 1, 2, 3
        imgCell(item.imgs[0], COL_W[4], { w: COL_W[4], bg }),
        imgCell(item.imgs[1], COL_W[5], { w: COL_W[5], bg }),
        imgCell(item.imgs[2], COL_W[6], { w: COL_W[6], bg }),
      ],
    })
  );
});

children.push(
  new Table({
    width:        { size: TBL_W, type: WidthType.DXA },
    columnWidths: COL_W,
    rows:         tableRows,
  })
);

// Footer
children.push(para(run(""), { sa: 60 }));
children.push(para(
  run("จึงเรียนมาเพื่อทราบและโปรดพิจารณา ให้แก่ทางบริษัทฯ ด้วยจักเป็นพระคุณยิ่ง", { size: 14 }),
  { sa: 20 }
));
children.push(para(run("ขอแสดงความนับถืออย่างสูง", { size: 14 }), { sa: 200 }));
children.push(para(run("( Mr. Ippei Ota )", { size: 14 }), { align: AlignmentType.CENTER, sa: 10 }));
children.push(para(run("กรรมการผู้จัดการ", { size: 14 }), { align: AlignmentType.CENTER }));

// ── Assemble & write ─────────────────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: pt(14) } } },
  },
  sections: [{
    properties: {
      page: {
        size:   { width: 11906, height: 16838 }, // A4
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }, // ~2cm
      },
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outPath, buf);
  console.log(`✓ Written: ${outPath} (${items.length} items)`);
}).catch((err) => {
  console.error("DOCX error:", err);
  process.exit(1);
});
