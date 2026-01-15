import fs from "fs";
import path from "path";
import { google } from "googleapis";
import dotenv from "dotenv";

dotenv.config();

const credentialsPath = path.join(path.dirname(new URL(import.meta.url).pathname), "../credentials.json");
const credentials = JSON.parse(fs.readFileSync(credentialsPath));

const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: ["https://www.googleapis.com/auth/drive.file"],
});

const FOLDER_ID = process.env.UPLOAD_FOLDER_ID;

if (!FOLDER_ID) {
  console.error("Error: UPLOAD_FOLDER_ID is not set in your environment variables.");
  console.log("Please add UPLOAD_FOLDER_ID=YOUR_GOOGLE_DRIVE_FOLDER_ID to your .env file.");
  process.exit(1);
}

async function uploadExcelAsSheet(localFilePath, sheetTitle) {
  const drive = google.drive({ version: "v3", auth: await auth.getClient() });
  const fileMetadata = {
    name: sheetTitle,
    mimeType: "application/vnd.google-apps.spreadsheet",
    parents: [FOLDER_ID], // Add this line
  };
  const media = {
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    body: fs.createReadStream(localFilePath),
  };
  const res = await drive.files.create({
    resource: fileMetadata,
    media: media,
    fields: "id, webViewLink",
  });
  console.log(`Google Sheet created for ${localFilePath}!`);
  console.log("Sheet ID:", res.data.id);
  console.log("Open in browser:", res.data.webViewLink);
}

// Convert both Excel files in data/
(async () => {
  await uploadExcelAsSheet("./data/MasterData with test data.xlsx", "MasterData with test data (Converted)");
  // await uploadExcelAsSheet("./data/MasterDataa.xlsx", "MasterDataa (Converted)");
  await uploadExcelAsSheet("./data/Week39_Template.xlsx", "Sycamore Week 39 Data");
})();