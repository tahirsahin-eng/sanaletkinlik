const SAYMAN_SIFRE = "fatma7952";
const SHEET_NAME = "VR_Etkinlik_Kayitlari";
const DRIVE_BACKUP_FOLDER = "Saffet_Cam_VR_Etkinlik_2026";

const EXPECTED_HEADERS = [
  "Kayıt ID",
  "Tarih / Saat",
  "Sınıf/Şube",
  "Okul No",
  "Öğrenci Adı Soyadı",
  "VR Seçim 1",
  "VR Seçim 2",
  "VR Seçim 3",
  "VR Seçim 4",
  "VR Seçim 5",
  "Ödeme Durumu",
  "Tutar (TL)",
  "Veli Adı Soyadı",
  "Veli Telefon",
  "Ödeme Notu",
  "VR_ID_JSON"
];

function getTargetSheet_(forceRecreate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet && !forceRecreate) {
    const activeSheet = ss.getActiveSheet();
    const firstCell = activeSheet && activeSheet.getLastRow() > 0
      ? String(activeSheet.getRange(1, 1).getValue() || "").trim()
      : "";
    if (firstCell === "Kayıt ID") {
      activeSheet.setName(SHEET_NAME);
      sheet = activeSheet;
    }
  }

  if (forceRecreate && sheet) {
    ss.deleteSheet(sheet);
    sheet = null;
  }

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(EXPECTED_HEADERS);
    sheet.getRange(1, 1, 1, EXPECTED_HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#1e1b4b")
      .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, EXPECTED_HEADERS.length);
  }

  return sheet;
}

function getFormattedRecords_(sheet) {
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) {
    return { records: [], registeredKeys: [] };
  }

  const records = [];
  const registeredKeys = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    const className = String(r[2] || "");
    const studentNo = String(r[3] || "");
    registeredKeys.push(className + "_" + studentNo);

    let vrIds = [];
    try {
      vrIds = JSON.parse(r[15] || "[]");
    } catch (_) {
      vrIds = [];
    }

    records.push({
      id: String(r[0]),
      updatedAt: String(r[1] || ""),
      className: className,
      studentNo: studentNo,
      studentName: String(r[4] || ""),
      vrTitles: [r[5], r[6], r[7], r[8], r[9]].filter(Boolean).map(String),
      vrIds: vrIds,
      paymentStatus: String(r[10]).indexOf("Ödendi") !== -1 ? "paid" : "pending",
      amount: Number(r[11] || 0),
      parentName: String(r[12] || ""),
      parentPhone: String(r[13] || ""),
      paymentNote: String(r[14] || "")
    });
  }

  return { records: records.reverse(), registeredKeys: registeredKeys };
}

function parseIncomingPayload_(e) {
  if (e && e.parameter && e.parameter.payload) {
    return JSON.parse(e.parameter.payload);
  }
  const rawBody = (e && e.postData && typeof e.postData.contents === "string")
    ? e.postData.contents.trim()
    : "";
  if (rawBody) {
    return JSON.parse(rawBody);
  }
  return {};
}

function processPayload_(payload) {
  const action = payload.action || "upsert";

  if (
    action === "list" ||
    action === "delete" ||
    action === "toggle_payment" ||
    action === "backup_drive" ||
    action === "rebuild_sheet"
  ) {
    if (String(payload.sifre || "").trim() !== SAYMAN_SIFRE) {
      return { ok: false, mesaj: "Yetkisiz işlem: Hatalı sayman şifresi." };
    }
  }

  if (action === "rebuild_sheet") {
    getTargetSheet_(true);
    return { ok: true, mesaj: "VR_Etkinlik_Kayitlari tablosu başlıklarla yeniden oluşturuldu." };
  }

  const sheet = getTargetSheet_(false);

  if (action === "list") {
    const data = getFormattedRecords_(sheet);
    return { ok: true, records: data.records };
  }

  const rows = sheet.getDataRange().getValues();

  if (action === "delete") {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(payload.id)) {
        sheet.deleteRow(i + 1);
        return { ok: true, mesaj: "Kayıt Google E-Tablodan silindi." };
      }
    }
    return { ok: false, mesaj: "Silinecek kayıt bulunamadı." };
  }

  if (action === "toggle_payment") {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(payload.id)) {
        const isPaid = payload.paymentStatus === "paid";
        sheet.getRange(i + 1, 2).setValue(new Date().toLocaleString("tr-TR"));
        sheet.getRange(i + 1, 11).setValue(isPaid ? "Ödendi (150 TL)" : "Ödeme Bekliyor");
        sheet.getRange(i + 1, 12).setValue(isPaid ? 150 : 0);
        return { ok: true, mesaj: "Ödeme durumu güncellendi." };
      }
    }
    return { ok: false, mesaj: "Güncellenecek kayıt bulunamadı." };
  }

  if (action === "backup_drive") {
    const folder = getOrCreateDriveFolder_(DRIVE_BACKUP_FOLDER);
    const csvLines = rows
      .map(row => row.map(c => '"' + String(c).replace(/"/g, '""') + '"').join(";"))
      .join("\n");
    const fileName =
      "Saffet_Cam_VR_Yedek_" +
      Utilities.formatDate(new Date(), "Europe/Istanbul", "yyyy-MM-dd_HH-mm") +
      ".csv";
    const file = folder.createFile(fileName, "\uFEFF" + csvLines, MimeType.CSV);
    return {
      ok: true,
      mesaj: "Google Drive klasörüne yedeklendi: " + file.getName(),
      url: file.getUrl()
    };
  }

  const rec = payload.record || payload;
  if (!rec || !rec.className || !rec.studentNo || !rec.studentName) {
    return { ok: false, mesaj: "Eksik öğrenci verisi gönderildi." };
  }

  const isAdminRequest = String(payload.sifre || "").trim() === SAYMAN_SIFRE;
  const paymentStatus = isAdminRequest && rec.paymentStatus === "paid" ? "paid" : "pending";

  let existingRowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(rec.id)) {
      existingRowIndex = i + 1;
      break;
    }
    if (
      !payload.isEdit &&
      String(rows[i][2]) === String(rec.className) &&
      String(rows[i][3]) === String(rec.studentNo)
    ) {
      return { ok: false, mesaj: "Bu öğrenci için daha önce kayıt oluşturulmuştur." };
    }
  }

  const vrTitles = Array.isArray(rec.vrTitles) ? rec.vrTitles : [];
  const vrIds = Array.isArray(rec.vrIds) ? rec.vrIds : [];

  const rowValues = [
    rec.id || ("rec_" + Date.now()),
    new Date().toLocaleString("tr-TR"),
    String(rec.className).trim(),
    String(rec.studentNo).trim(),
    String(rec.studentName).trim(),
    vrTitles[0] || "",
    vrTitles[1] || "",
    vrTitles[2] || "",
    vrTitles[3] || "",
    vrTitles[4] || "",
    paymentStatus === "paid" ? "Ödendi (150 TL)" : "Ödeme Bekliyor",
    paymentStatus === "paid" ? 150 : 0,
    String(rec.parentName || "").trim(),
    String(rec.parentPhone || "").trim(),
    String(rec.paymentNote || "").trim(),
    JSON.stringify(vrIds)
  ];

  if (existingRowIndex > 0) {
    sheet.getRange(existingRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return { ok: true, mesaj: "Başvuru Google E-Tabloya başarıyla kaydedildi." };
}

function doGet(e) {
  const callback = (e && e.parameter && e.parameter.callback)
    ? String(e.parameter.callback).replace(/[^a-zA-Z0-9_]/g, "")
    : "";
  try {
    if (e && e.parameter && e.parameter.payload) {
      const payload = JSON.parse(e.parameter.payload);
      return outputResponse_(processPayload_(payload), callback);
    }

    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "list";
    const sheet = getTargetSheet_(false);
    const data = getFormattedRecords_(sheet);

    if (action === "public_check") {
      return outputResponse_({ ok: true, registeredKeys: data.registeredKeys }, callback);
    }

    const sifre = (e && e.parameter && e.parameter.sifre) ? String(e.parameter.sifre).trim() : "";
    if (sifre !== SAYMAN_SIFRE) {
      return outputResponse_({ ok: false, mesaj: "Yetkisiz erişim: Sayman şifresi doğrulanamadı." }, callback);
    }

    return outputResponse_({ ok: true, records: data.records }, callback);
  } catch (err) {
    return outputResponse_({ ok: false, mesaj: err.message }, callback);
  }
}

function doPost(e) {
  try {
    const payload = parseIncomingPayload_(e);
    return jsonResponse_(processPayload_(payload));
  } catch (err) {
    return jsonResponse_({ ok: false, mesaj: "Sunucu hatası: " + err.message });
  }
}

function getOrCreateDriveFolder_(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function outputResponse_(obj, callback) {
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + JSON.stringify(obj) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonResponse_(obj);
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}