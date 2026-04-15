import fs from "node:fs";
import { google } from "googleapis";
import formidable from "formidable";

export const config = {
  api: {
    bodyParser: false
  }
};

const REQUIRED_ENV = [
  "GOOGLE_CLIENT_EMAIL",
  "GOOGLE_PRIVATE_KEY",
  "GOOGLE_SPREADSHEET_ID",
  "GOOGLE_DRIVE_FOLDER_ID"
];

function parseForm(req) {
  const form = formidable({
    multiples: false,
    maxFiles: 1,
    maxFileSize: 10 * 1024 * 1024,
    keepExtensions: true
  });

  return new Promise((resolve, reject) => {
    form.parse(req, (error, fields, files) => {
      if (error) {
        reject(error);
        return;
      }

      resolve({ fields, files });
    });
  });
}

function pickFirst(value) {
  return Array.isArray(value) ? value[0] : value;
}

function assertEnv() {
  const missing = REQUIRED_ENV.filter((name) => !process.env[name]);
  if (missing.length > 0) {
    throw new Error("Faltan variables en Vercel: " + missing.join(", "));
  }
}

function buildGoogleClients() {
  const auth = new google.auth.JWT({
    email: process.env.GOOGLE_CLIENT_EMAIL,
    key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),
    scopes: [
      "https://www.googleapis.com/auth/drive",
      "https://www.googleapis.com/auth/spreadsheets"
    ]
  });

  return {
    drive: google.drive({ version: "v3", auth }),
    sheets: google.sheets({ version: "v4", auth })
  };
}

async function uploadPhoto(drive, photoFile, remito) {
  const originalName = photoFile.originalFilename || "foto";
  const fileName = `${remito || "entrega"}-${Date.now()}-${originalName}`;

  const createdFile = await drive.files.create({
    requestBody: {
      name: fileName,
      parents: [process.env.GOOGLE_DRIVE_FOLDER_ID]
    },
    media: {
      mimeType: photoFile.mimetype || "application/octet-stream",
      body: fs.createReadStream(photoFile.filepath)
    },
    fields: "id, webViewLink"
  });

  const fileId = createdFile.data.id;
  await drive.permissions.create({
    fileId,
    requestBody: {
      role: "reader",
      type: "anyone"
    }
  });

  const fileData = await drive.files.get({
    fileId,
    fields: "id, webViewLink, webContentLink"
  });

  return fileData.data.webViewLink || fileData.data.webContentLink || "";
}

async function appendRow(sheets, row) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: process.env.GOOGLE_SPREADSHEET_ID,
    range: "A1",
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: [row]
    }
  });
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Metodo no permitido." });
    return;
  }

  let photoFile;

  try {
    assertEnv();

    const { fields, files } = await parseForm(req);
    const cadete = String(pickFirst(fields.cadete) || "").trim();
    const afiliado = String(pickFirst(fields.afiliado) || "").trim();
    const remito = String(pickFirst(fields.remito) || "").trim();
    const factura = String(pickFirst(fields.factura) || "").trim();
    const submittedAt = String(pickFirst(fields.submittedAt) || new Date().toISOString()).trim();
    photoFile = pickFirst(files.foto);

    if (!cadete || !afiliado || !remito) {
      res.status(400).json({ error: "Faltan datos obligatorios." });
      return;
    }

    if (!photoFile) {
      res.status(400).json({ error: "Falta la foto." });
      return;
    }

    const { drive, sheets } = buildGoogleClients();
    const photoUrl = await uploadPhoto(drive, photoFile, remito);

    await appendRow(sheets, [
      submittedAt,
      cadete,
      afiliado,
      remito,
      factura,
      photoUrl,
      process.env.PHARMACY_NAME || "Farmacia"
    ]);

    res.status(200).json({
      ok: true,
      remito,
      photoUrl
    });
  } catch (error) {
    res.status(500).json({
      error: error.message || "No se pudo guardar la entrega."
    });
  } finally {
    if (photoFile && photoFile.filepath) {
      fs.promises.unlink(photoFile.filepath).catch(() => {});
    }
  }
}
