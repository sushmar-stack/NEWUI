import fs from "fs";
import path from "path";
import { google } from "googleapis";
import dotenv from "dotenv";

dotenv.config();

const credentialsPath = path.join(
  path.dirname(new URL(import.meta.url).pathname),
  "../credentials.json"
);

if (!fs.existsSync(credentialsPath)) {
  throw new Error(`Credentials file not found at: ${credentialsPath}`);
}

const credentials = JSON.parse(fs.readFileSync(credentialsPath));

const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
  ],
});

export const drive = google.drive({ version: "v3", auth });
export const sheets = google.sheets({ version: "v4", auth });