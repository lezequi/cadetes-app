const SPREADSHEET_ID = "13UpbfB0tqwyZ4y4FKnh1nPto2CuXIKmKe-1FMZjp0DA";
const FOLDER_ID = "16BDn64N1bQ56GLBLcwfMaHgj9-Z9tWI1";
const SHEET_NAME = "App";

const HEADERS = [
  "Fecha de carga",
  "Fecha y hora comprobante",
  "Cadete",
  "Afiliado",
  "Remito",
  "Factura",
  "Documento",
  "Error Foto"
];

function doGet() {
  return ContentService.createTextOutput("ok");
}

function doPost(e) {
  let response = {
    type: "cadetes-submission",
    ok: false,
    saved: false,
    error: ""
  };

  try {
    const sheet = getSheet_();
    const data = e && e.parameter ? e.parameter : {};

    ensureSheetSetup_(sheet);

    const submittedAt = clean_(data.submittedAt) || new Date().toISOString();
    const fechaHoraComprobante = clean_(data.fechaHoraComprobante);
    const cadete = clean_(data.cadete);
    const afiliado = clean_(data.afiliado);
    const remito = clean_(data.remito);
    const factura = clean_(data.factura);
    const photoBase64 = clean_(data.photoBase64);
    const photoName = clean_(data.photoName) || "comprobante.jpg";
    const photoType = clean_(data.photoType) || "image/jpeg";

    let photoUrl = "";
    let photoError = "";

    if (photoBase64) {
      try {
        const folder = DriveApp.getFolderById(FOLDER_ID);
        const bytes = Utilities.base64Decode(photoBase64);
        const blob = Utilities.newBlob(bytes, photoType, buildFileName_(remito, photoName));
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrl = file.getUrl();
      } catch (error) {
        photoError = error && error.message ? error.message : "Error al guardar comprobante";
      }
    } else {
      photoError = "No llego archivo adjunto";
    }

    sheet.appendRow([
      formatDate_(submittedAt),
      formatLocalInputDate_(fechaHoraComprobante),
      cadete,
      afiliado,
      remito,
      factura,
      photoUrl,
      photoError
    ]);

    const lastRow = sheet.getLastRow();
    if (photoUrl) {
      const richText = SpreadsheetApp.newRichTextValue()
        .setText("Ver comprobante")
        .setLinkUrl(photoUrl)
        .build();
      sheet.getRange(lastRow, 7).setRichTextValue(richText);
    }

    styleDataRow_(sheet, lastRow);

    response = {
      type: "cadetes-submission",
      ok: !photoError,
      saved: true,
      remito,
      photoUrl,
      error: photoError
    };
  } catch (error) {
    response.error = error && error.message ? error.message : "No se pudo guardar la entrega";
  }

  return buildPostMessageResponse_(response);
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  return sheet;
}

function ensureSheetSetup_(sheet) {
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  formatSheet_(sheet);
}

function formatSheet_(sheet) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight("bold")
    .setBackground("#5b3f8c")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 210);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 170);
  sheet.setColumnWidth(5, 130);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 170);
  sheet.setColumnWidth(8, 240);

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  const lastRow = Math.max(sheet.getLastRow(), 2);
  sheet.getRange(1, 1, lastRow, HEADERS.length).createFilter();
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), HEADERS.length)
    .setVerticalAlignment("middle");
  sheet.setRowHeights(1, Math.max(sheet.getMaxRows(), 2), 28);
}

function styleDataRow_(sheet, row) {
  sheet.getRange(row, 1, 1, HEADERS.length)
    .setBackground(row % 2 === 0 ? "#f7f4fb" : "#ffffff");
}

function formatDate_(isoString) {
  const date = new Date(isoString);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
}

function formatLocalInputDate_(value) {
  if (!value) {
    return "";
  }

  const date = new Date(value);
  if (isNaN(date.getTime())) {
    return value;
  }

  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
}

function buildFileName_(remito, photoName) {
  const safeRemito = clean_(remito) || "entrega";
  return safeRemito + "-" + new Date().getTime() + "-" + photoName;
}

function buildPostMessageResponse_(payload) {
  const json = JSON.stringify(payload).replace(/</g, "\\u003c");
  return HtmlService.createHtmlOutput(
    '<!doctype html><html><body><script>' +
    'var payload=' + json + ';' +
    'try{window.parent.postMessage(payload,"*");}catch(e){}' +
    'try{window.top.postMessage(payload,"*");}catch(e){}' +
    '</script></body></html>'
  );
}

function clean_(value) {
  return String(value || "").trim();
}

function resetAppSheet() {
  const sheet = getSheet_();
  sheet.clear();
  sheet.clearFormats();
  ensureSheetSetup_(sheet);
}

function testWrite() {
  const sheet = getSheet_();
  ensureSheetSetup_(sheet);
  sheet.appendRow([
    formatDate_(new Date().toISOString()),
    formatLocalInputDate_("2026-04-24T12:00"),
    "Codex",
    "Farmacia",
    "REM-MANUAL",
    "FAC-MANUAL",
    "",
    ""
  ]);
  styleDataRow_(sheet, sheet.getLastRow());
}
