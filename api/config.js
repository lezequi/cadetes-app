export default function handler(_req, res) {
  res.status(200).json({
    pharmacyName: process.env.PHARMACY_NAME || "Farmacia",
    configured: Boolean(
      process.env.GOOGLE_CLIENT_EMAIL &&
      process.env.GOOGLE_PRIVATE_KEY &&
      process.env.GOOGLE_SPREADSHEET_ID &&
      process.env.GOOGLE_DRIVE_FOLDER_ID
    )
  });
}
