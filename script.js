// ================== BIẾN TOÀN CỤC ==================
let cachedWorkbook = null;
let STT_ROWS = [];
let EXCEL_SO_QD = "";

// ================== HELPER ==================
function v(id) {
  return document.getElementById(id)?.value?.trim() || "";
}

const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// ================== SET TODAY ==================
function setToday() {
  const d = new Date();
  document.getElementById("NGAY").value = String(d.getDate()).padStart(2, "0");
  document.getElementById("THANG").value = String(d.getMonth() + 1).padStart(2, "0");
  document.getElementById("NAM").value = d.getFullYear();
}

// ================== EXPORT DOCX ==================
async function exportDOCX(filename = "Quyet_dinh.docx") {
  try {
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
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });

    saveAs(out, filename);
  } catch (error) {
    console.error("Lỗi xuất file:", filename, error);
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

// ================== EXPORT THEO CHI NHÁNH (KHÔNG ZIP) ==================
async function exportByBranch() {
  const sheetName = document.getElementById("sheetSelect").value;
  if (!sheetName) return alert("❌ Chưa chọn sheet");

  const sheet = cachedWorkbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  let currentBranch = "";
  let currentSoQD = "";
  let members = [];
  let fileCount = 0;

  // Duyệt từ dòng thứ 2 (i=1) để bỏ qua Header
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    
    const colC = row[2] ? row[2].toString().trim() : ""; // Tên Chi nhánh
    const colD = row[3] ? row[3].toString().trim() : ""; // Họ tên
    const colE = row[4] ? row[4].toString().trim() : ""; // Chức vụ CQ
    const colF = row[5] ? row[5].toString().trim() : ""; // Chức vụ Đoàn
    const colG = row[6] ? row[6].toString().trim() : ""; // Số QD

    // Nếu gặp tên Chi nhánh mới
    if (colC !== "") {
      // Xuất file cho chi nhánh cũ trước khi reset sang chi nhánh mới
      if (currentBranch !== "" && members.length > 0) {
        await runExport(currentBranch, currentSoQD, members);
        fileCount++;
        await delay(400); // ⏳ Quan trọng: Nghỉ để trình duyệt không chặn download
      }

      // Khởi tạo dữ liệu cho chi nhánh mới
      currentBranch = colC;
      currentSoQD = colG;
      members = [];
    }

    // Nếu có tên thành viên ở cột D thì thêm vào danh sách
    if (colD !== "") {
      members.push({
        name: colD,
        role: colE,
        extra: colF
      });
    }
  }

  // Xuất chi nhánh cuối cùng sau khi kết thúc vòng lặp
  if (currentBranch !== "" && members.length > 0) {
    await runExport(currentBranch, currentSoQD, members);
    fileCount++;
  }

  alert(`✅ Hoàn tất! Đã yêu cầu tải xuống ${fileCount} file.`);
}

// Hàm bổ trợ điền dữ liệu và gọi hàm xuất Word
async function runExport(branch, soQD, membersList) {
  document.getElementById("CHI_NHANH").value = branch;
  document.getElementById("SO_QD").value = soQD;

  STT_ROWS = membersList.map((m, i) => ({
    stt: i + 1,
    name: m.name,
    role: m.role,
    extra: m.extra
  }));

  const safeName = branch.replace(/[\\/:*?"<>|]/g, "_");
  await exportDOCX(`QD_${safeName}.docx`);
}