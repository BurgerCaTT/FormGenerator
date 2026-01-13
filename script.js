// ================== BIẾN TOÀN CỤC ==================
let cachedWorkbook = null;
let STT_ROWS = [];
let EXCEL_SO_QD = "";

// ================== HELPER ==================
function v(id) {
  return document.getElementById(id)?.value?.trim() || "";
}

// ================== SET TODAY ==================
function setToday() {
  const d = new Date();
  document.getElementById("NGAY").value = String(d.getDate()).padStart(2, "0");
  document.getElementById("THANG").value = String(d.getMonth() + 1).padStart(2, "0");
  document.getElementById("NAM").value = d.getFullYear();
}

// ================== EXPORT DOCX (LOGIC CŨ – GIỮ NGUYÊN) ==================
async function exportDOCX(filename = "Quyet_dinh.docx") {
  const res = await fetch("VanBan.docx");
  const content = await res.arrayBuffer();

  const zip = new PizZip(content);
  const doc = new window.docxtemplater(zip, {
    delimiters: { start: "[[", end: "]]" },
    paragraphLoop: true,
    linebreaks: true
  });

  doc.setData({
    SO_QD: v("SO_QD"),
    NGAY: v("NGAY"),
    THANG: v("THANG"),
    NAM: v("NAM"),
    CHI_NHANH: v("CHI_NHANH"),
    STT_ROWS
  });

  doc.render();

  const out = doc.getZip().generate({
    type: "blob",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });

  saveAs(out, filename);
}

// ================== TẠO STT_ROWS TỪ NHẬP TAY ==================
function buildRowsFromManual() {
  STT_ROWS = [];
  let stt = 1;

  for (let i = 1; i <= 6; i++) {
    const name = v(`f${i}0`);
    const role = v(`f${i}1`);
    const extra = v(`f${i}2`);

    if (name) {
      STT_ROWS.push({
        stt: stt++,
        name,
        role,
        extra
      });
    }
  }
}

// ================== IMPORT EXCEL ==================
async function importExcelToForm() {
  const file = document.getElementById("excelFile").files[0];
  if (!file) return alert("❌ Chưa chọn Excel");

  setToday();

  const data = await file.arrayBuffer();
  cachedWorkbook = XLSX.read(data, { type: "array" });

  const select = document.getElementById("sheetSelect");
  select.innerHTML = `<option value="">-- Chọn sheet --</option>`;

  cachedWorkbook.SheetNames.forEach(name => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    select.appendChild(opt);
  });

  select.style.display = "inline-block";
  document.getElementById("btnExportBranch").style.display = "inline-block";
}

// ================== EXPORT THEO CHI NHÁNH (EXCEL) ==================
async function exportByBranch() {
  const sheetName = document.getElementById("sheetSelect").value;
  if (!sheetName) return alert("❌ Chưa chọn sheet");

  const sheet = cachedWorkbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: ""
  });

  let currentBranch = "";
  let members = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];

    const colC = row[2]; // Chi nhánh
    const colD = row[3]; // Họ tên
    const colE = row[4]; // Chức vụ
    const colF = row[5]; // Ghi chú
    const colG = row[6]; // SO_QD ✅

    if (colC) {
      if (currentBranch && members.length) {
        await generateBranchDoc(currentBranch, members);
      }

      currentBranch = colC;
      members = [];

      if (colG) {
        EXCEL_SO_QD = colG; // ✅ LƯU SO_QD THEO CHI NHÁNH
      }
    }

    if (colD) {
      members.push({
        name: colD,
        role: colE,
        extra: colF
      });
    }
  }

  if (currentBranch && members.length) {
    await generateBranchDoc(currentBranch, members);
  }

  alert("✅ Xuất xong tất cả file theo chi nhánh");
}

// ================== GEN FILE DOCX (EXCEL) ==================
async function generateBranchDoc(branch, members) {
  document.getElementById("CHI_NHANH").value = branch;
  document.getElementById("SO_QD").value = EXCEL_SO_QD;

  STT_ROWS = members.map((m, i) => ({
    stt: i + 1,
    name: m.name,
    role: m.role,
    extra: m.extra
  }));

  const safeName = branch.replace(/[\\/:*?"<>|]/g, "_");
  await exportDOCX(`QD_${safeName}.docx`);
}
