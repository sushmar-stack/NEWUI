import { sheets, drive } from "./google.js";
import dotenv from "dotenv";
import { google } from "googleapis";
import path from 'path';
import { fileURLToPath } from 'url';
import { promises as fsp } from "fs";
import XLSX from 'xlsx';
import { updateGoogleSheetCell } from './excel.js';

dotenv.config();

/**
 * Updates data in a Google Sheet.
 * @param {string} spreadsheetId The ID of the spreadsheet.
 * @param {string} customerName The name of the customer (column) to update.
 * @param {object} data The data to write.
 */
export async function updateGoogleSheetData(spreadsheetId, customerName, data) {
  console.log(`Updating data for ${customerName} in Google Sheet ${spreadsheetId}...`);

  const updatePromises = [];

  // Iterate through the new data and create a promise for each cell update.
  for (const [category, items] of Object.entries(data)) {
    if (!Array.isArray(items)) continue; // Skip non-array properties

    for (const item of items) {
      const [newSubcategory, newValue] = item.split(/:(.*)/s);

      // Use the robust updateGoogleSheetCell function for each update.
      // This function handles row creation if the subcategory doesn't exist.
      updatePromises.push(
        updateGoogleSheetCell(spreadsheetId, customerName, newSubcategory.trim(), (newValue || '').trim())
      );
    }
  }

  // Execute all updates concurrently.
  await Promise.all(updatePromises);
  console.log(`Successfully processed ${updatePromises.length} updates for ${customerName}.`);
}

/**
 * Public export for adding a new client.
 */
export async function addGoogleSheetClient(spreadsheetId, customerName, customerData) {
  const __filename = fileURLToPath(import.meta.url);
  const __dirname = path.dirname(__filename);
  const credPath = path.join(__dirname, '..', 'credentials.json'); // Corrected path import
  const credsContent = await fsp.readFile(credPath, 'utf8');
  const creds = JSON.parse(credsContent);

  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: creds.client_email,
      private_key: (creds.private_key || creds.privateKey).replace(/\\n/g, '\n'),
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const sheets = google.sheets({ version: 'v4', auth });
  const sheetName = 'Sheet1'; // Assuming the main data is in Sheet1

  // 1. Get the current sheet details to find the sheetId
  const sheetDetailsResponse = await sheets.spreadsheets.get({
    spreadsheetId: spreadsheetId,
    fields: 'sheets.properties',
  });
  const sheetId = sheetDetailsResponse.data.sheets.find(s => s.properties.title === sheetName)?.properties.sheetId;

  if (sheetId === undefined) {
      throw new Error(`Sheet with name "${sheetName}" not found in spreadsheet ID: ${spreadsheetId}`);
  }

  // 2. Prepare the values for the new column
  const getResponse = await sheets.spreadsheets.values.get({
    spreadsheetId: spreadsheetId,
    range: `${sheetName}!A:Z`, // Read more columns to find the last data column
  });

  const rows = getResponse.data.values || [];
  if (rows.length === 0) throw new Error("Sheet is empty or missing data rows.");

  const headerRow = rows[0];
  const subcategoryColIndex = headerRow.findIndex(h => (h || "").toString().toLowerCase().includes('subcategory'));

  if (subcategoryColIndex === -1) throw new Error("Subcategory column not found in sheet.");

  const newColumnValues = [customerName]; // The first value is the customer name (header)

  // Map the new customer data to a key-value object for quick lookup
  const allSubcategories = {};
  for (const category of Object.keys(customerData)) {
    (customerData[category] || []).forEach(item => {
      const [key, value] = item.split(/:(.*)/s).map(s => s.trim());
      if (key) allSubcategories[key] = value || '';
    });
  }

  // Fill the rest of the column with corresponding data
  rows.slice(1).forEach(row => {
    const subcategory = (row[subcategoryColIndex] || "").toString().trim();
    newColumnValues.push(subcategory ? (allSubcategories[subcategory] || '') : '');
  });

  // 3. Find the column index for insertion
  // Find the last column that has a value in the header row.
  let lastHeaderIndex = headerRow.length - 1;
  while (lastHeaderIndex > 0 && !headerRow[lastHeaderIndex]) {
    lastHeaderIndex--;
  }
  const insertIndex = lastHeaderIndex + 1;

  // 4. Insert the new column
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: spreadsheetId,
    requestBody: {
      requests: [{
        insertDimension: {
          range: {
            sheetId: sheetId,
            dimension: 'COLUMNS',
            startIndex: insertIndex,
            endIndex: insertIndex + 1,
          }
        }
      }]
    }
  });

  // 5. Update values in the new column
  await sheets.spreadsheets.values.update({
    spreadsheetId: spreadsheetId,
    range: `${sheetName}!${XLSX.utils.encode_col(insertIndex)}1`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: newColumnValues.map(v => [v]) }
  });
}

/**
 * Deletes a customer from all sheets in a Google Spreadsheet.
 * @param {string} dataSourceId The ID of the Google Sheet to update.
 * @param {string} customerName The name of the customer to delete.
 */
export async function deleteGoogleSheetClient(dataSourceId, customerName) {
  const auth = new google.auth.GoogleAuth({
    keyFile: "credentials.json",
    scopes: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive",
    ],
  });

  const authClient = await auth.getClient();
  const sheets = google.sheets({
    version: "v4",
    auth: authClient,
  });

  console.log(`Deleting customer ${customerName} from Google Sheet ${dataSourceId}...`);

  // 1. Fetch the current sheet data to find customer column
  const sheetData = await sheets.spreadsheets.get({
    spreadsheetId: dataSourceId,
    includeGridData: true,
  });

  const gridData = sheetData.data.sheets[0].data[0];
  const rows = gridData.rowData;

  // Find the header row and customer column index
  let headerRowIndex = -1;
  let customerColIndex = -1;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i].values || [];
    const customerIndex = row.findIndex(cell => cell.formattedValue === customerName);
    if (customerIndex !== -1) {
      headerRowIndex = i;
      customerColIndex = customerIndex;
      break;
    }
  }

  if (headerRowIndex === -1 || customerColIndex === -1) {
    throw new Error(`Could not find customer "${customerName}" in sheet ${dataSourceId}`);
  }

  // 2. Delete the column
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: dataSourceId,
    requestBody: {
      requests: [{
        deleteDimension: {
          range: {
            sheetId: sheetData.data.sheets[0].properties.sheetId,
            dimension: 'COLUMNS',
            startIndex: customerColIndex,
            endIndex: customerColIndex + 1,
          }
        }
      }]
    }
  });

  console.log(`Successfully deleted customer ${customerName} from sheet ${dataSourceId}.`);
}

export { drive };