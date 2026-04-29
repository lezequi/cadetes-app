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
  "Error Foto",
  "Tipo de carga"
];

function doGet() {
  return ContentService.createTextOutput("ok");
}

function doPost(e) {
  try {
    const sheet = getSheet_();
    const data = e && e.parameter ? e.parameter : {};
    setupSheet_(sheet);

    const remito = clean_(data.remito);
    const paymentType = clean_(data.paymentType) === "efectivo" ? "Pago en efectivo" : "Con comprobante";
    const cashPayment = paymentType === "Pago en efectivo";
    let photoUrl = "";
    let photoError = "";

    if (clean_(data.photoBase64)) {
      try {
        const folder = DriveApp.getFolderById(FOLDER_ID);
        const bytes = Utilities.base64Decode(clean_(data.photoBase64));
        const fileName = fileName_(remito, clean_(data.photoName) || "comprobante.jpg");
        const blob = Utilities.newBlob(bytes, clean_(data.photoType) || "image/jpeg", fileName);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrl = file.getUrl();
      } catch (error) {
        photoError = message_(error, "Error al guardar comprobante");
      }
    } else if (!cashPayment) {
      photoError = "No llego archivo adjunto";
    }

    sheet.appendRow([
      formatDate_(clean_(data.submittedAt) || new Date().toISOString()),
      formatInputDate_(data.fechaHoraComprobante),
      clean_(data.cadete),
      clean_(data.afiliado),
      remito,
      clean_(data.factura),
      "",
      photoError,
      paymentType
    ]);

    const row = sheet.getLastRow();
    setLink_(sheet, row, 7, photoUrl);
    styleRow_(sheet, row);

    return response_({
      type: "cadetes-submission",
      ok: !photoError,
      saved: true,
      remito,
      photoUrl,
      error: photoError
    });
  } catch (error) {
    return response_({
      type: "cadetes-submission",
      ok: false,
      saved: false,
      error: message_(error, "No se pudo guardar la entrega")
    });
  }
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function setupSheet_(sheet) {
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setFontWeight("bold")
    .setBackground("#5b3f8c")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  [170, 210, 150, 170, 130, 130, 170, 240, 160].forEach(function (width, index) {
    sheet.setColumnWidth(index + 1, width);
  });

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 2), HEADERS.length).createFilter();
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), HEADERS.length).setVerticalAlignment("middle");
  sheet.setRowHeights(1, Math.max(sheet.getMaxRows(), 2), 28);
}

function setLink_(sheet, row, column, url) {
  if (!url) {
    return;
  }

  sheet.getRange(row, column).setRichTextValue(
    SpreadsheetApp.newRichTextValue().setText("Ver comprobante").setLinkUrl(url).build()
  );
}

function styleRow_(sheet, row) {
  sheet.getRange(row, 1, 1, HEADERS.length).setBackground(row % 2 === 0 ? "#f7f4fb" : "#ffffff");
}

function repararLinksDocumento() {
  const sheet = getSheet_();
  const column = findColumn_(sheet, "Documento");
  if (!column || sheet.getLastRow() < 2) {
    return;
  }

  const formulas = sheet.getRange(2, column, sheet.getLastRow() - 1, 1).getFormulas();
  let repaired = 0;
  formulas.forEach(function (row, index) {
    const url = extractUrl_(row[0]);
    if (url) {
      setLink_(sheet, index + 2, column, url);
      repaired += 1;
    }
  });
  Logger.log("Links reparados: " + repaired);
}

function resetAppSheet() {
  const sheet = getSheet_();
  sheet.clear();
  sheet.clearFormats();
  setupSheet_(sheet);
}

function testWrite() {
  const sheet = getSheet_();
  setupSheet_(sheet);
  sheet.appendRow([
    formatDate_(new Date().toISOString()),
    "29/04/2026 10:22",
    "Codex",
    "Farmacia",
    "REM-MANUAL",
    "FAC-MANUAL",
    "",
    "",
    "Pago en efectivo"
  ]);
  styleRow_(sheet, sheet.getLastRow());
}

function findColumn_(sheet, header) {
  const values = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (let index = 0; index < values.length; index += 1) {
    if (clean_(values[index]) === header) {
      return index + 1;
    }
  }
  return 0;
}

function extractUrl_(formula) {
  const match = String(formula || "").match(/https:\/\/[^"',;)]+/);
  return match ? match[0] : "";
}

function fileName_(remito, name) {
  return (clean_(remito) || "entrega") + "-" + new Date().getTime() + "-" + name;
}

function formatDate_(value) {
  return Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
}

function formatInputDate_(value) {
  const text = clean_(value);
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
  return match ? match[3] + "/" + match[2] + "/" + match[1] + " " + match[4] + ":" + match[5] : text;
}

function response_(payload) {
  const json = JSON.stringify(payload).replace(/</g, "\\u003c");
  return HtmlService.createHtmlOutput(
    "<!doctype html><html><body><script>" +
    "var payload=" + json + ";" +
    "try{window.parent.postMessage(payload,'*');}catch(e){}" +
    "try{window.top.postMessage(payload,'*');}catch(e){}" +
    "</script></body></html>"
  );
}

function message_(error, fallback) {
  return error && error.message ? error.message : fallback;
}

function clean_(value) {
  return String(value || "").trim();
}
