const SAYMAN_SIFRE = "fatma7952";
const SHEET_NAME = "VR_Etkinlik_Kayitlari";
const DRIVE_BACKUP_FOLDER = "Saffet_Cam_VR_Etkinlik_2026";
const EVENT_FEE = 150;

const EXPECTED_HEADERS = [
  "Kayıt ID",
  "Tarih / Saat",
  "Katılımcı Türü",
  "Sınıf/Şube",
  "Okul No",
  "Öğrenci / Katılımcı Adı Soyadı",
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

function normalizeRecordType_(rawType) {
  const val = String(rawType || "ogrenci").trim().toLowerCase();
  if (val === "veli" || val === "parent") return "veli";
  if (val === "misafir" || val === "guest") return "misafir";
  return "ogrenci";
}

function getRecordTypeLabel_(recordType) {
  if (recordType === "veli") return "Veli";
  if (recordType === "misafir") return "Misafir Öğrenci";
  return "Öğrenci";
}

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

  const headers = rows[0].map(h => String(h || "").trim());
  const hasTypeCol = headers[2] === "Katılımcı Türü";

  const records = [];
  const registeredKeys = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;

    const recordType = hasTypeCol ? normalizeRecordType_(r[2]) : "ogrenci";
    const className = hasTypeCol ? String(r[3] || "") : String(r[2] || "");
    const studentNo = hasTypeCol ? String(r[4] || "") : String(r[3] || "");
    const participantName = hasTypeCol ? String(r[5] || "") : String(r[4] || "");

    const vrCols = hasTypeCol ? [r[6], r[7], r[8], r[9], r[10]] : [r[5], r[6], r[7], r[8], r[9]];
    const paymentCol = hasTypeCol ? r[11] : r[10];
    const amountCol = hasTypeCol ? r[12] : r[11];
    const parentNameCol = hasTypeCol ? r[13] : r[12];
    const parentPhoneCol = hasTypeCol ? r[14] : r[13];
    const paymentNoteCol = hasTypeCol ? r[15] : r[14];
    const vrJsonCol = hasTypeCol ? r[16] : r[15];

    if (recordType === "ogrenci" && className && studentNo) {
      registeredKeys.push(className + "_" + studentNo);
    }

    let vrIds = [];
    try {
      vrIds = JSON.parse(vrJsonCol || "[]");
    } catch (_) {
      vrIds = [];
    }

    records.push({
      id: String(r[0]),
      updatedAt: String(r[1] || ""),
      recordType: recordType,
      recordTypeLabel: getRecordTypeLabel_(recordType),
      className: className || (recordType === "veli" ? "Veli" : "Misafir"),
      studentNo: studentNo || "-",
      studentName: participantName,
      vrTitles: vrCols.filter(Boolean).map(String),
      vrIds: vrIds,
      paymentStatus: String(paymentCol).indexOf("Ödendi") !== -1 ? "paid" : "pending",
      amount: Number(amountCol || 0),
      parentName: String(parentNameCol || ""),
      parentPhone: String(parentPhoneCol || ""),
      paymentNote: String(paymentNoteCol || "")
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
    return { ok: true, mesaj: "VR_Etkinlik_Kayitlari tablosu güncel başlıklarla yeniden oluşturuldu." };
  }

  const sheet = getTargetSheet_(false);

  if (action === "list") {
    const data = getFormattedRecords_(sheet);
    return { ok: true, records: data.records };
  }

  const rows = sheet.getDataRange().getValues();
  const headers = rows.length > 0 ? rows[0].map(h => String(h || "").trim()) : EXPECTED_HEADERS;
  const hasTypeCol = headers[2] === "Katılımcı Türü";

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
    const dateColIdx = 2;
    const statusColIdx = hasTypeCol ? 12 : 11;
    const amountColIdx = hasTypeCol ? 13 : 12;

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(payload.id)) {
        const isPaid = payload.paymentStatus === "paid";
        sheet.getRange(i + 1, dateColIdx).setValue(new Date().toLocaleString("tr-TR"));
        sheet.getRange(i + 1, statusColIdx).setValue(isPaid ? "Ödendi (" + EVENT_FEE + " TL)" : "Ödeme Bekliyor");
        sheet.getRange(i + 1, amountColIdx).setValue(isPaid ? EVENT_FEE : 0);
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
  const recordType = normalizeRecordType_(rec.recordType || rec.participantType || rec.type);
  const participantName = String(rec.studentName || rec.participantName || rec.parentName || "").trim();
  const className = String(rec.className || "").trim();
  const studentNo = String(rec.studentNo || "").trim();
  const parentName = String(rec.parentName || "").trim();
  const parentPhone = String(rec.parentPhone || "").trim();
  const paymentNote = String(rec.paymentNote || "").trim();

  if (recordType === "ogrenci") {
    if (!className || !studentNo || !participantName) {
      return { ok: false, mesaj: "Öğrenci kaydı için sınıf/şube, okul no ve ad soyad zorunludur." };
    }
  } else if (recordType === "veli") {
    if (!participantName && !parentName) {
      return { ok: false, mesaj: "Veli kaydı için ad soyad bilgisi zorunludur." };
    }
  } else if (recordType === "misafir") {
    if (!participantName) {
      return { ok: false, mesaj: "Misafir öğrenci kaydı için ad soyad bilgisi zorunludur." };
    }
  }

  const finalName = participantName || parentName;
  const finalClass = className || (recordType === "veli" ? "Veli" : "Misafir");
  const finalNo = studentNo || "-";

  const isAdminRequest = String(payload.sifre || "").trim() === SAYMAN_SIFRE;
  const paymentStatus = isAdminRequest && rec.paymentStatus === "paid" ? "paid" : "pending";

  let existingRowIndex = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(rec.id)) {
      existingRowIndex = i + 1;
      break;
    }
    const rowType = hasTypeCol ? normalizeRecordType_(rows[i][2]) : "ogrenci";
    const rowClass = hasTypeCol ? String(rows[i][3] || "") : String(rows[i][2] || "");
    const rowNo = hasTypeCol ? String(rows[i][4] || "") : String(rows[i][3] || "");

    if (
      !payload.isEdit &&
      recordType === "ogrenci" &&
      rowType === "ogrenci" &&
      rowClass === className &&
      rowNo === studentNo
    ) {
      return { ok: false, mesaj: "Bu öğrenci için daha önce kayıt oluşturulmuştur." };
    }
  }

  const vrTitles = Array.isArray(rec.vrTitles) ? rec.vrTitles : [];
  const vrIds = Array.isArray(rec.vrIds) ? rec.vrIds : [];

  const rowValues = hasTypeCol
    ? [
        rec.id || ("rec_" + Date.now()),
        new Date().toLocaleString("tr-TR"),
        getRecordTypeLabel_(recordType),
        finalClass,
        finalNo,
        finalName,
        vrTitles[0] || "",
        vrTitles[1] || "",
        vrTitles[2] || "",
        vrTitles[3] || "",
        vrTitles[4] || "",
        paymentStatus === "paid" ? "Ödendi (" + EVENT_FEE + " TL)" : "Ödeme Bekliyor",
        paymentStatus === "paid" ? EVENT_FEE : 0,
        parentName || (recordType === "veli" ? finalName : ""),
        parentPhone,
        paymentNote,
        JSON.stringify(vrIds)
      ]
    : [
        rec.id || ("rec_" + Date.now()),
        new Date().toLocaleString("tr-TR"),
        finalClass,
        finalNo,
        finalName,
        vrTitles[0] || "",
        vrTitles[1] || "",
        vrTitles[2] || "",
        vrTitles[3] || "",
        vrTitles[4] || "",
        paymentStatus === "paid" ? "Ödendi (" + EVENT_FEE + " TL)" : "Ödeme Bekliyor",
        paymentStatus === "paid" ? EVENT_FEE : 0,
        parentName || (recordType === "veli" ? finalName : ""),
        parentPhone,
        paymentNote,
        JSON.stringify(vrIds)
      ];

  if (existingRowIndex > 0) {
    sheet.getRange(existingRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return { ok: true, mesaj: getRecordTypeLabel_(recordType) + " başvurusu Google E-Tabloya başarıyla kaydedildi." };
}

function addManualRecordForHtml(sifre, manualRec) {
  return processPayload_({
    action: "upsert",
    sifre: sifre,
    record: manualRec
  });
}

function doGet(e) {
  const isSaymanView = e && e.parameter && String(e.parameter.sayman || "") === "1";
  if (isSaymanView) {
    return HtmlService.createHtmlOutputFromFile("SaymanPanel")
      .setTitle("Saffet Çam Ortaokulu - VR Sayman Paneli")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
  }

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

function getSaymanRecordsForHtml(sifre) {
  return processPayload_({ action: "list", sifre: sifre });
}

function togglePaymentForHtml(sifre, id, paymentStatus) {
  return processPayload_({ action: "toggle_payment", sifre: sifre, id: id, paymentStatus: paymentStatus });
}

function deleteRecordForHtml(sifre, id) {
  return processPayload_({ action: "delete", sifre: sifre, id: id });
}

function backupDriveForHtml(sifre) {
  return processPayload_({ action: "backup_drive", sifre: sifre });
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