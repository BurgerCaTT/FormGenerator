async function exportDOCX() {

  const res = await fetch("Chuany.docx");
  const content = await res.arrayBuffer();

  const zip = new PizZip(content);
  const doc = new window.docxtemplater(zip, {
     delimiters: { start: '[[', end: ']]' },
    paragraphLoop: true,
    linebreaks: true
  });

  doc.setData({
    SO_QD: v("SO_QD"),
    NGAY: v("NGAY"),
    THANG: v("THANG"),
    NAM: v("NAM"),
    CHI_NHANH: v("CHI_NHANH"),

    F10: v("f10"),
    F11: v("f11"),
    F20: v("f20"),
    F21: v("f21"),
    F30: v("f30"),
    F31: v("f31"),
    F40: v("f40"),
    F41: v("f41"),
    F50: v("f50"),
    F51: v("f51")
  });

  try {
    doc.render();
  } catch (e) {
    console.error(e);
    alert("❌ Lỗi template – kiểm tra placeholder");
    return;
  }

  const out = doc.getZip().generate({
    type: "blob",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });

  saveAs(out, "Quyet_dinh_da_dien.docx");
}

function v(id) {
  return document.getElementById(id)?.value || "";
}
