import XLSX from "xlsx";
import fetch from "node-fetch";
import fs, { promises as fsp } from "fs";
import { google } from "googleapis";
import path from 'path';
import { fileURLToPath } from 'url';

// --- Constants ---
const PRODUCT_UPDATE_KEYS = ['Customer Name', 'Current State', 'Next Up', 'Top 3 items in upcoming Release(s)', 'Tech Stack/Infra Upgrades (As Needed)'];
const PRODUCT_UPDATE_SHEET_NAME = 'ProductUpdates';
const FRONTEND_TAB_KEYS = ['currentState', 'nextUp', 'top3', 'techStack'];

const CLIENT_SPECIFIC_DETAILS_SHEET_NAME = 'ClientSpecificDetails';
const CLIENT_SPECIFIC_DETAILS_KEYS = ['Customer Name', 'Deployment Details', 'Scheduled Activities/ Backlog', 'Product Development & Services Alignment', 'Performance Metrics from last week'];
const CLIENT_SPECIFIC_FRONTEND_KEYS = ['deploymentDetails', 'scheduledActivities', 'productAlignment', 'performanceMetrics'];

const TRACKER_SHEET_PREFIX = 'Tracker ';
const PL_SHEET_PREFIX = 'PL ';

// List of subcategories that are expected to contain document links
const DOCUMENT_SUBCATEGORIES = [
  'sow',
  'qm and certification',
  'product documents',
  'deployment documents',
  'consolidated document'
];
// ============================================================================
//  HELPER: Google Auth & Sheets Client
// ============================================================================
async function getGoogleSheetsClient() {
  const __filename = fileURLToPath(import.meta.url);
  const __dirname = path.dirname(__filename);
  const credPath = path.join(__dirname, '..', 'credentials.json');
  
  if (!fs.existsSync(credPath)) {
    throw new Error('credentials.json not found at ' + credPath);
  }

  const credsContent = await fsp.readFile(credPath, 'utf8');
  const creds = JSON.parse(credsContent);

  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: creds.client_email,
      private_key: (creds.private_key || creds.privateKey).replace(/\\n/g, '\n'),
    },
    scopes: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive.readonly' 
    ],
  });

  return google.sheets({ version: 'v4', auth });
}

// ============================================================================
//  CORE: Load Data Logic
// ============================================================================

export async function loadData(sources, selectedWeek) {
  if (typeof sources === 'string') sources = [sources];

  const masterSource = sources[0];
  const allData = {};

  if (masterSource) {
    const masterData = await loadSingleSource(masterSource);
    mergeSheetData(allData, masterData);

    const weeklyUpdates = await loadWeeklyUpdates(masterSource);
    if (weeklyUpdates) {
      allData._weeklyUpdates = weeklyUpdates;
    }
  }

  const weeklySource = getWeeklySource(sources, selectedWeek);
  if (weeklySource) {
    const weeklyData = await loadSingleSource(weeklySource);
      const previousWeek = getPreviousWeek(selectedWeek);
      const previousWeeklySource = getWeeklySource(sources, previousWeek);

      if (previousWeeklySource) {
        const previousWeekData = await loadSingleSource(previousWeeklySource);
        mergeSheetData(weeklyData, previousWeekData, true); // Overwrite with previous week's data
      }
    mergeSheetData(allData, weeklyData);
  }

  return allData;
}

async function loadWeeklyUpdates(source) {
  try {
    const wb = await fetchGoogleSheetWorkbook(source);
    const sheet = wb.Sheets['WeeklyUpdates'];
    if (!sheet) return null;

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const updates = {};
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (row[0]) updates[row[0]] = row[1] || '';
    }
    return updates;
  } catch (error) {
    console.warn(`Could not load 'WeeklyUpdates' sheet from ${source}:`, error.message);
    return null;
  }
}

async function loadSingleSource(source) {
  try {
    const workbook = await fetchGoogleSheetWorkbook(source);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    return processSheetData(sheet, rawRows, source);
  } catch (error) {
    console.error(`Error loading source ${source}:`, error.message);
    return {};
  }
}

// ============================================================================
//  CORE: Update Data Logic (HYBRID: BATCH UPDATE + APPEND)
// ============================================================================

/**
 * Updates data in Google Sheets.
 * 1. Checks if row exists -> Batches an update.
 * 2. If row MISSING -> Appends a new row.
 */
export async function updateGoogleSheetData(spreadsheetId, customerName, data, masterSheetId = null) {
  const sheets = await getGoogleSheetsClient();

  // 1. Get the correct sheet name (Don't assume 'Sheet1')
  let sheetName = 'Sheet1';
  try {
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    sheetName = meta.data.sheets[0].properties.title;
  } catch (e) {
    console.warn("Could not fetch sheet metadata, defaulting to Sheet1");
  }

  console.log(`[Update] Processing ${customerName} in "${sheetName}" (ID: ${spreadsheetId})...`);

  // 2. Fetch Data to map coordinates
  let rows = [];
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:Z`, 
    });
    rows = res.data.values || [];
  } catch (e) {
    console.warn(`[Update] Failed to fetch data: ${e.message}`);
    return;
  }

  if (rows.length === 0) { console.warn("Sheet is empty!"); return; }

  const headerRow = rows[0];
  const cIdx = headerRow.findIndex(h => (h || '').toString().trim() === customerName);
  const scColIdx = headerRow.findIndex(h => (h || '').toString().toLowerCase().includes('subcategory'));

  if (cIdx === -1 || scColIdx === -1) {
    console.warn(`[Update] Headers not found for Customer: ${customerName}`);
    return;
  }

  const batchUpdates = [];
  const rowsToAppend = [];

  // 3. Categorize changes: Update vs Append
  for (const [category, items] of Object.entries(data)) {
    if (!Array.isArray(items)) continue;

    for (const item of items) {
      const [newSubcategory, newValue] = item.split(/:(.*)/s);
      const scName = newSubcategory.trim();
      const val = (newValue || '').trim();

      // Find Row Index in current data
      let rIdx = -1;
      for (let i = 1; i < rows.length; i++) {
        const currentRow = rows[i];
        if (!currentRow) continue; // Skip if the row is undefined/empty
        if ((currentRow[scColIdx] || '').toString().trim() === scName) {
          rIdx = i; // 0-based index
          break;
        }
      }

      if (rIdx !== -1) {
        // Exists: Add to Batch Update
        batchUpdates.push({
          range: `${sheetName}!${XLSX.utils.encode_cell({ c: cIdx, r: rIdx })}`,
          values: [[val]]
        });
      } else {
        // Missing: Add to Append List
        console.log(`[Update] Row missing for "${scName}". Appending...`);
        const newRow = new Array(headerRow.length).fill('');
        newRow[scColIdx] = scName;
        newRow[cIdx] = val;
        
        // Helper to guess category if needed (simple logic)
        const catColIdx = headerRow.findIndex(h => (h || '').toString().toLowerCase().includes('sycamore'));
        if (catColIdx !== -1) newRow[catColIdx] = category.includes('Sycamore') ? 'Sycamore' : 'Client';

        rowsToAppend.push(newRow);
      }
    }
  }

  // 4. Execute Batch Updates (Existing Rows)
  if (batchUpdates.length > 0) {
    console.log(`[Update] Updating ${batchUpdates.length} existing cells...`);
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: { valueInputOption: 'USER_ENTERED', data: batchUpdates }
    });
  }

  // 5. Execute Appends (New Rows)
  if (rowsToAppend.length > 0) {
    console.log(`[Update] Appending ${rowsToAppend.length} new rows...`);
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetName}!A:A`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: rowsToAppend }
    });
  }

  console.log(`[Update] Complete.`);
}

/**
 * DEPRECATED: Retained only to prevent import errors.
 */
export async function updateGoogleSheetCell(dataSourceId, customerName, subcategoryName, newValue) {
   console.warn("Using deprecated updateGoogleSheetCell. Please migrate to updateGoogleSheetData.");
}

// ============================================================================
//  FEATURE: Product Updates
// ============================================================================

export async function loadProductUpdateData(dataSourceId, customerName) {
  try {
    const workbook = await fetchGoogleSheetWorkbook(dataSourceId);
    const sheet = workbook.Sheets[PRODUCT_UPDATE_SHEET_NAME];
    if (!sheet) throw new Error(`Sheet '${PRODUCT_UPDATE_SHEET_NAME}' not found`);

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (rows.length < 2) return {};

    const headerRow = rows[0];
    const dataRow = rows.find(row => (row[0] || '').toString().trim() === customerName);
    if (!dataRow) return {}; 
    
    const parsedData = {};
    const columnIndices = {};
    headerRow.forEach((title, index) => { columnIndices[title.toString().trim()] = index; });

    PRODUCT_UPDATE_KEYS.slice(1).forEach((title, index) => {
        const colIndex = columnIndices[title];
        if (colIndex !== undefined && dataRow[colIndex] !== undefined) {
            const frontendKey = FRONTEND_TAB_KEYS[index];
            parsedData[frontendKey] = dataRow[colIndex].toString().trim();
        }
    });
    return parsedData;
  } catch (error) {
    if (error.message.includes(`Sheet '${PRODUCT_UPDATE_SHEET_NAME}' not found`)) throw error;
    console.error(`Error in loadProductUpdateData:`, error.message);
    throw new Error("Failed to process product update sheet data.");
  }
}

export async function updateProductUpdateData(dataSourceId, customerName, data) {
  const sheets = await getGoogleSheetsClient();
  let sheetId;

  try {
    const sheetDetails = await sheets.spreadsheets.get({ spreadsheetId: dataSourceId, fields: 'sheets.properties' });
    const sheet = sheetDetails.data.sheets.find(s => s.properties.title === PRODUCT_UPDATE_SHEET_NAME);
    sheetId = sheet?.properties.sheetId;

    if (sheetId === undefined) {
      const addSheetRes = await sheets.spreadsheets.batchUpdate({
        spreadsheetId: dataSourceId,
        requestBody: { requests: [{ addSheet: { properties: { title: PRODUCT_UPDATE_SHEET_NAME } } }] }
      });
      sheetId = addSheetRes.data.replies[0].addSheet.properties.sheetId;
      await sheets.spreadsheets.values.update({
        spreadsheetId: dataSourceId,
        range: `${PRODUCT_UPDATE_SHEET_NAME}!A1:E1`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [PRODUCT_UPDATE_KEYS] },
      });
    }
  } catch (error) {
    throw new Error('Failed to access/create ProductUpdates sheet.');
  }

  const getRes = await sheets.spreadsheets.values.get({ spreadsheetId: dataSourceId, range: `${PRODUCT_UPDATE_SHEET_NAME}!A:Z` });
  const rows = getRes.data.values || [];
  const customerRowIndex = rows.findIndex(row => (row[0] || '').toString().trim() === customerName);
  
  const newRowValues = [
    customerName,
    data.currentState || '',
    data.nextUp || '',
    data.top3 || '',
    data.techStack || '',
  ];
  
  const headerRow = rows[0] || PRODUCT_UPDATE_KEYS;
  const headerMap = {};
  headerRow.forEach((title, index) => { headerMap[title] = index; });

  if (customerRowIndex !== -1) {
    const startCol = XLSX.utils.encode_col(headerMap[PRODUCT_UPDATE_KEYS[1]]);
    const endCol = XLSX.utils.encode_col(headerMap[PRODUCT_UPDATE_KEYS.at(-1)]);
    const range = `${PRODUCT_UPDATE_SHEET_NAME}!${startCol}${customerRowIndex + 1}:${endCol}${customerRowIndex + 1}`;
    await sheets.spreadsheets.values.update({
      spreadsheetId: dataSourceId,
      range: range,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [newRowValues.slice(1)] }, 
    });
  } else {
    await sheets.spreadsheets.values.append({
      spreadsheetId: dataSourceId,
      range: `${PRODUCT_UPDATE_SHEET_NAME}!A:A`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [newRowValues] },
    });
  }
}

// ============================================================================
//  FEATURE: Client Specific Details
// ============================================================================

export async function loadClientSpecificDetailsData(dataSourceId, customerName) {
  try {
    const workbook = await fetchGoogleSheetWorkbook(dataSourceId);
    const sheet = workbook.Sheets[CLIENT_SPECIFIC_DETAILS_SHEET_NAME];
    if (!sheet) throw new Error(`Sheet '${CLIENT_SPECIFIC_DETAILS_SHEET_NAME}' not found`);

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (rows.length < 2) return {};

    const headerRow = rows[0];
    const dataRow = rows.find(row => (row[0] || '').toString().trim() === customerName);
    if (!dataRow) return {}; 
    
    const parsedData = {};
    const columnIndices = {};
    headerRow.forEach((title, index) => { columnIndices[title.toString().trim()] = index; });

    CLIENT_SPECIFIC_DETAILS_KEYS.slice(1).forEach((title, index) => {
        const colIndex = columnIndices[title];
        if (colIndex !== undefined && dataRow[colIndex] !== undefined) {
            const frontendKey = CLIENT_SPECIFIC_FRONTEND_KEYS[index];
            parsedData[frontendKey] = dataRow[colIndex].toString().trim();
        }
    });
    return parsedData;
  } catch (error) {
    if (error.message.includes(`Sheet '${CLIENT_SPECIFIC_DETAILS_SHEET_NAME}' not found`)) throw error;
    console.error(`Error in loadClientSpecificDetailsData:`, error.message);
    throw new Error("Failed to process client specific details sheet data.");
  }
}

export async function updateClientSpecificDetailsData(dataSourceId, customerName, data) {
  const sheets = await getGoogleSheetsClient();
  let sheetId;

  try {
    const sheetDetails = await sheets.spreadsheets.get({ spreadsheetId: dataSourceId, fields: 'sheets.properties' });
    const sheet = sheetDetails.data.sheets.find(s => s.properties.title === CLIENT_SPECIFIC_DETAILS_SHEET_NAME);
    sheetId = sheet?.properties.sheetId;

    if (sheetId === undefined) {
      const addSheetRes = await sheets.spreadsheets.batchUpdate({
        spreadsheetId: dataSourceId,
        requestBody: { requests: [{ addSheet: { properties: { title: CLIENT_SPECIFIC_DETAILS_SHEET_NAME } } }] }
      });
      sheetId = addSheetRes.data.replies[0].addSheet.properties.sheetId;
      await sheets.spreadsheets.values.update({
        spreadsheetId: dataSourceId,
        range: `${CLIENT_SPECIFIC_DETAILS_SHEET_NAME}!A1:E1`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [CLIENT_SPECIFIC_DETAILS_KEYS] },
      });
    }
  } catch (error) {
    throw new Error('Failed to access or create ClientSpecificDetails sheet.');
  }

  const getRes = await sheets.spreadsheets.values.get({ spreadsheetId: dataSourceId, range: `${CLIENT_SPECIFIC_DETAILS_SHEET_NAME}!A:A` });
  const rows = getRes.data.values || [];
  const customerRowIndex = rows.findIndex(row => (row[0] || '').toString().trim() === customerName);
  const newRowValues = [customerName, ...CLIENT_SPECIFIC_FRONTEND_KEYS.map(key => data[key] || '')];
  
  if (customerRowIndex !== -1) {
    const range = `${CLIENT_SPECIFIC_DETAILS_SHEET_NAME}!B${customerRowIndex + 1}:E${customerRowIndex + 1}`;
    await sheets.spreadsheets.values.update({
      spreadsheetId: dataSourceId, 
      range: range, 
      valueInputOption: 'USER_ENTERED', 
      requestBody: { values: [newRowValues.slice(1)] },
    });
  } else {
    await sheets.spreadsheets.values.append({
      spreadsheetId: dataSourceId, 
      range: `${CLIENT_SPECIFIC_DETAILS_SHEET_NAME}!A:A`, 
      valueInputOption: 'USER_ENTERED', 
      requestBody: { values: [newRowValues] },
    });
  }
}

// ============================================================================
//  FEATURE: Trackers & Project Lists
// ============================================================================

export async function loadTrackerData(masterSheetId, customerName, year) {
  const sheetName = `${TRACKER_SHEET_PREFIX}${year}`;
  try {
    const workbook = await fetchGoogleSheetWorkbook(masterSheetId);
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return {}; 

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (rows.length < 2) return {};

    const headerRow = rows[0];
    const dateColIndex = headerRow.findIndex(h => (h || '').toString().toLowerCase().includes('date'));
    const clientColIndex = headerRow.findIndex(h => (h || '').toString().trim() === customerName);

    if (dateColIndex === -1 || clientColIndex === -1) return {};
    
    const trackerData = {};
    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const date = (row[dateColIndex] || '').toString().trim();
        const content = (row[clientColIndex] || '').toString().trim();
        if (date && content) trackerData[date] = content;
    }
    return trackerData;
  } catch (error) {
    console.error(`Error in loadTrackerData:`, error.message);
    return {};
  }
}

export async function updateTrackerData(masterSheetId, customerName, date, content) {
  const sheets = await getGoogleSheetsClient();
  const year = new Date(date).getFullYear();
  const sheetName = `${TRACKER_SHEET_PREFIX}${year}`;

  let rows;
  try {
    const sheetDetails = await sheets.spreadsheets.get({ spreadsheetId: masterSheetId, fields: 'sheets.properties' });
    const sheet = sheetDetails.data.sheets.find(s => s.properties.title === sheetName);
    
    if (sheet) {
        const valRes = await sheets.spreadsheets.values.get({ spreadsheetId: masterSheetId, range: `${sheetName}!A:Z` });
        rows = valRes.data.values || [];
    } else {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: masterSheetId,
        requestBody: { requests: [{ addSheet: { properties: { title: sheetName } } }] }
      });
      await sheets.spreadsheets.values.update({
        spreadsheetId: masterSheetId,
        range: `${sheetName}!A1:B1`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [['Date', customerName]] },
      });
      rows = [['Date', customerName]];
    }
  } catch (error) {
    throw new Error('Failed to Tracker sheet: ' + error.message);
  }

  const headerRow = rows[0];
  let dateColIndex = headerRow.findIndex(h => (h || '').toString().toLowerCase().includes('date'));
  let clientColIndex = headerRow.findIndex(h => (h || '').toString().trim() === customerName);
  let dateRowIndex = -1;

  if (clientColIndex === -1) {
    clientColIndex = headerRow.length;
    await sheets.spreadsheets.values.update({
        spreadsheetId: masterSheetId,
        range: `${sheetName}!${XLSX.utils.encode_col(clientColIndex)}1`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [[customerName]] },
    });
  }

  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][dateColIndex] || '').toString().trim() === date) {
      dateRowIndex = i + 1;
      break;
    }
  }

  if (dateRowIndex !== -1) {
    const range = `${sheetName}!${XLSX.utils.encode_cell({ c: clientColIndex, r: dateRowIndex - 1 })}`;
    await sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: range,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[content]] },
    });
  } else {
    const newRow = new Array(clientColIndex + 1).fill('');
    newRow[dateColIndex] = date;
    newRow[clientColIndex] = content;
    await sheets.spreadsheets.values.append({
      spreadsheetId: masterSheetId,
      range: `${sheetName}!A:A`, 
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [newRow] },
    });
  }
}

export async function loadProjectListData(masterSheetId, customerName) {
  try {
    const workbook = await fetchGoogleSheetWorkbook(masterSheetId);
    const plData = {};
    const plSheetNames = workbook.SheetNames.filter(name => name.startsWith(PL_SHEET_PREFIX));

    for (const sheetName of plSheetNames) {
      const year = sheetName.replace(PL_SHEET_PREFIX, '').trim();
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) continue;

      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      if (rows.length < 2) continue;

      const headerRow = rows[0];
      const clientColIndex = headerRow.findIndex(h => (h || '').toString().trim() === customerName);

      if (clientColIndex !== -1) {
        const content = (rows[1] && rows[1][clientColIndex]) ? rows[1][clientColIndex].toString().trim() : '';
        if (content) plData[year] = content;
      }
    }
    return plData;
  } catch (error) {
    console.error(`Error in loadProjectListData:`, error.message);
    throw new Error("Failed to process project list sheets.");
  }
}

export async function updateProjectListData(masterSheetId, customerName, year, content) {
  const sheets = await getGoogleSheetsClient();
  const sheetName = `${PL_SHEET_PREFIX}${year}`;

  let rows;
  try {
    const sheetDetails = await sheets.spreadsheets.get({ spreadsheetId: masterSheetId, fields: 'sheets.properties' });
    const sheet = sheetDetails.data.sheets.find(s => s.properties.title === sheetName);

    if (sheet) {
      const valRes = await sheets.spreadsheets.values.get({ spreadsheetId: masterSheetId, range: `${sheetName}!A:Z` });
      rows = valRes.data.values || [];
    } else {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: masterSheetId,
        requestBody: { requests: [{ addSheet: { properties: { title: sheetName } } }] }
      });
      rows = []; 
    }
  } catch (error) {
    throw new Error('Failed to Project List sheet: ' + error.message);
  }

  const headerRow = rows.length > 0 ? rows[0] : [];
  let clientColIndex = headerRow.findIndex(h => (h || '').toString().trim() === customerName);

  if (clientColIndex === -1) {
    clientColIndex = headerRow.length;
    await sheets.spreadsheets.values.update({
      spreadsheetId: masterSheetId,
      range: `${sheetName}!${XLSX.utils.encode_col(clientColIndex)}1`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[customerName]] },
    });
  }

  const range = `${sheetName}!${XLSX.utils.encode_cell({ c: clientColIndex, r: 1 })}`;
  await sheets.spreadsheets.values.update({
    spreadsheetId: masterSheetId,
    range: range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [[content]] },
  });
}

// ============================================================================
//  UTILITIES & HELPERS
// ============================================================================

export function getWeeklySource(sources, selectedWeek) {
    if (!selectedWeek || selectedWeek === 'master') return null;
    const weekMatch = selectedWeek.match(/^(\d{4})-W(\d{1,2})$/);
    if (!weekMatch) return null;

    const year = parseInt(weekMatch[1], 10);
    const weekNum = parseInt(weekMatch[2], 10);

    const yearSourcesEnv = `WEEKLY_SOURCES_${year}`;
    const yearSources = process.env[yearSourcesEnv];

    if (!yearSources) return null;

    const sourceIds = yearSources.split(',');
    // Assuming week numbers in env map directly to array indices (e.g., week 1 is at index 0)
    return sourceIds[weekNum - 1] || null;
}

function getPreviousWeek(selectedWeek) {
    const weekMatch = selectedWeek.match(/^(\d{4})-W(\d{1,2})$/);
    if (!weekMatch) return null;
    const year = parseInt(weekMatch[1], 10);
    const week = parseInt(weekMatch[2], 10);
    if (week === 1) return `${year - 1}-W52`;
    return `${year}-W${week - 1}`;
}
export function getAvailableWeeks(sources) {
  const weeks = [];
  const baseWeek = parseInt(process.env.WEEK_START || '1', 10);
  for (let i = 1; i < sources.length; i++) {
    weeks.push(`week${baseWeek + (i - 1)}`);
  }
  return weeks;
}

function processSheetData(sheet, rawRows, sourceId = '<unknown>') {
  const normalize = (s) => (s || '').toString().toLowerCase().trim().replace(/[^a-z0-9]/g, '');

  const headerRowIndex = rawRows.findIndex(row =>
    row.some(cell => normalize(cell).includes('subcategory'))
  );

  if (headerRowIndex === -1) {
    console.warn(`Could not find a header row with 'Subcategory' in sheet ${sourceId}.`);
    return {};
  }

  const headerRow = rawRows[headerRowIndex];
  const subcategoryColIndex = headerRow.findIndex(h => normalize(h).includes('subcategory'));
  const sycamoreColIdx = headerRow.findIndex(h => {
    const n = normalize(h);
    return (n.includes('sycamore') && n.includes('client')) || n === 'sycamoreclient' || n === 'clienttype';
  });

  const mapSycamoreValue = (v) => {
    const n = (v || '').toString().toLowerCase();
    if (n.includes('sycamore') && n.includes('client')) return 'Sycamore and Client';
    if (n.includes('sycamore') && !n.includes('client')) return 'Sycamore';
    if (n.includes('client') && !n.includes('sycamore')) return 'Client';
    if (n.includes('both') || n.includes('and')) return 'Sycamore and Client';
    return 'Client';
  };

  const custCols = [];
  const custNames = [];
  for (let i = 0; i < headerRow.length; i++) {
    const n = normalize(headerRow[i]);
    if (n !== 'category' && n !== 'subcategory' && n !== 'staticdynamic' && n !== 'sycamoreclient') {
      custCols.push(i);
      custNames.push(headerRow[i]);
    }
  }

  const result = {};

  for (let r = headerRowIndex + 1; r < rawRows.length; r++) {
    const row = rawRows[r];
    const subcategory = (row[subcategoryColIndex] || '').toString().trim();
    if (!subcategory) continue; // Skip empty subcategories

    const scValue = sycamoreColIdx !== -1 ? (row[sycamoreColIdx] || '').toString().trim() : '';
    let targetCategory = mapSycamoreValue(scValue);

    custCols.forEach((colIndex, ci) => {
      const cust = custNames[ci];
      const value = (row[colIndex] || '').toString().trim() || 'No Data';
      const normalizedSubcategory = normalize(subcategory);
      
      if (!result[cust]) {
        result[cust] = { 'Client': [], 'Sycamore': [], 'Sycamore and Client': [], _logoUrl: '' }; // Removed _sowUrl
      }

      let finalValueToPush = value;
      let extractedHyperlink = '';

      // Apply robust hyperlink extraction for document-related subcategories
      if (DOCUMENT_SUBCATEGORIES.includes(normalizedSubcategory)) {
        const cellAddress = XLSX.utils.encode_cell({ c: colIndex, r: r });
        const cell = sheet[cellAddress];
        
        // 1. Try standard hyperlink
        extractedHyperlink = cell?.l?.Target;

        // 2. If no standard link, try to extract from HYPERLINK() formula
        if (!extractedHyperlink && cell?.f && cell.f.includes('HYPERLINK')) {
           // Regex to grab the URL between the first set of quotes
           const matches = cell.f.match(/HYPERLINK\s*\(\s*"([^"]+)"/i);
           if (matches && matches[1]) {
             extractedHyperlink = matches[1];
           }
        }
        
        // 3. Fallback: If cell value acts like a URL (starts with http)
        if (!extractedHyperlink && cell?.v && typeof cell.v === 'string' && cell.v.toLowerCase().startsWith('http')) {
            extractedHyperlink = cell.v;
        }
        if (extractedHyperlink && !value.includes(extractedHyperlink)) {
            finalValueToPush = `${value} [LINK: ${extractedHyperlink}]`; // Append URL in a parsable format
        }
      } else if (normalizedSubcategory === 'logo') { // Keep logo special handling
        const displayValue = value.toLowerCase().startsWith('http') ? 'Link Available' : value;
        result[cust][targetCategory].push(`${subcategory}: ${displayValue}`);
        return;
      }

      if (subcategory === 'Customer Location' || subcategory === 'Customer Description' || subcategory === 'Customer Name') {
        targetCategory = 'Client';
      }
      
      result[cust][targetCategory].push(`${subcategory}: ${finalValueToPush}`);
    });
  }

  return result;
}

function mergeSheetData(allData, sheetData, overwrite = false) {
  for (const [customerName, categories] of Object.entries(sheetData)) {
    if (!allData[customerName]) {
      allData[customerName] = {
        "Client": [],
        "Sycamore": [],
        "Sycamore and Client": [],
        _logoUrl: ''
      };
    }
    
    for (const [category, items] of Object.entries(categories)) {
      const existingItems = allData[customerName][category];
      if (!existingItems) {
        allData[customerName][category] = [];
      }

      if (category.startsWith('_')) {
        if (items) allData[customerName][category] = items;
        continue;
      }

      items.forEach(item => {
        const subcategory = item.split(':')[0].trim();
        const existingIndex = existingItems.findIndex(
          existingItem => existingItem.split(':')[0].trim() === subcategory
        );
        if (existingIndex !== -1) {
          existingItems[existingIndex] = item;
        } else if (overwrite) {
          existingItems[existingIndex] = item;
        } else {
          existingItems.push(item);
        }
      });
    }
  }
}

async function fetchGoogleSheetWorkbook(sheetId) {
  const sheets = await getGoogleSheetsClient();
  const drive = google.drive({ version: 'v3', auth: sheets.context._options.auth });

  try {
    const res = await drive.files.export({
      fileId: sheetId,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      supportsAllDrives: true
    }, { responseType: 'arraybuffer' });

    const buffer = Buffer.from(res.data);
    const wb = XLSX.read(buffer, { type: 'buffer' });
    return wb;
  } catch (err) {
    console.error('Drive export error for', sheetId, { message: err.message });
    throw err;
  }
}

export async function addGoogleSheetClient(spreadsheetId, customerName, customerData) {
  const sheets = await getGoogleSheetsClient();
  const sheetName = 'Sheet1';

  // 1. Get Sheet ID
  const sheetDetails = await sheets.spreadsheets.get({ spreadsheetId: spreadsheetId, fields: 'sheets.properties' });
  const sheetId = sheetDetails.data.sheets.find(s => s.properties.title === sheetName)?.properties.sheetId;

  if (sheetId === undefined) throw new Error(`Sheet "${sheetName}" not found.`);

  // 2. Prepare Data
  const valRes = await sheets.spreadsheets.values.get({ spreadsheetId: spreadsheetId, range: `${sheetName}!A:C` });
  const rows = valRes.data.values || [];
  if (rows.length === 0) throw new Error("Sheet is empty.");

  const headerRow = rows[0];
  const subcategoryColIndex = headerRow.findIndex(h => (h || "").toString().toLowerCase().includes('subcategory'));

  const newColumnValues = [customerName];
  const allSubcategories = [
    ...customerData.Client, 
    ...customerData.Sycamore, 
    ...customerData["Sycamore and Client"]
  ].reduce((acc, item) => {
    const [key, value] = item.split(/:(.*)/s).map(s => s.trim());
    if (key) acc[key] = value || '';
    return acc;
  }, {});

  rows.slice(1).forEach(row => {
    const subcategory = (row[subcategoryColIndex] || "").toString().trim();
    newColumnValues.push(subcategory ? (allSubcategories[subcategory] || '') : '');
  });

  // 3. Insert Column & Add Data
  const insertIndex = rows.reduce((max, row) => Math.max(max, row.length), 0);
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: spreadsheetId,
    requestBody: { requests: [{ insertDimension: { range: { sheetId: sheetId, dimension: 'COLUMNS', startIndex: insertIndex, endIndex: insertIndex + 1 } } }] }
  });

  const range = `${sheetName}!${XLSX.utils.encode_col(insertIndex)}:${XLSX.utils.encode_col(insertIndex)}`;
  await sheets.spreadsheets.values.update({
    spreadsheetId: spreadsheetId,
    range: range,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: newColumnValues.map(v => [v]) }
  });
}

export async function deleteGoogleSheetClient(dataSourceId, customerName) {
  const sheets = await getGoogleSheetsClient();
  
  // 1. Find Column
  const sheetData = await sheets.spreadsheets.get({ spreadsheetId: dataSourceId, includeGridData: true });
  const gridData = sheetData.data.sheets[0].data[0];
  const rows = gridData.rowData;

  let customerColIndex = -1;
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i].values || [];
    const idx = row.findIndex(cell => cell.formattedValue === customerName);
    if (idx !== -1) {
      customerColIndex = idx;
      break;
    }
  }

  if (customerColIndex === -1) throw new Error(`Customer "${customerName}" not found`);

  // 2. Delete Column
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
}

export function listCustomers(db) {
  return Object.keys(db).filter(key => !key.startsWith('_')).sort();
}

export async function updateWeeklyUpdate(spreadsheetId, week, text) {
  const sheets = await getGoogleSheetsClient();
  const updateSheetName = 'WeeklyUpdates';

  try {
    // 1. Check if the sheet exists, and create it if it doesn't.
    const sheetDetails = await sheets.spreadsheets.get({ spreadsheetId, fields: 'sheets.properties' });
    const sheet = sheetDetails.data.sheets.find(s => s.properties.title === updateSheetName);

    if (!sheet) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: [{ addSheet: { properties: { title: updateSheetName } } }] }
      });
      // Add headers to the new sheet
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${updateSheetName}!A1:B1`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [['Week', 'Update']] },
      });
    }

    // 2. Get existing data to find the row to update or append.
    const valRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: `${updateSheetName}!A:B` });
    const rows = valRes.data.values || [];
    const rowIndex = rows.findIndex(row => row && row[0] === week);

    if (rowIndex !== -1) {
      // Update existing row
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${updateSheetName}!B${rowIndex + 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [[text]] },
      });
    } else {
      // Append new row
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: `${updateSheetName}!A:A`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [[week, text]] },
      });
    }
  } catch (error) {
    console.error(`Failed to update weekly update for week "${week}":`, error);
    throw new Error(`Failed to save weekly update: ${error.message}`);
  }
}