import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import multer from 'multer';
import fs from 'fs';
import path from 'path';

// Import necessary functions from excel.js and sheetScheduler
import { 
  loadData, 
  listCustomers, 
  updateWeeklyUpdate, 
  getWeeklySource, 
  updateGoogleSheetCell, 
  updateGoogleSheetData, 
  loadProductUpdateData, 
  updateProductUpdateData, 
  loadClientSpecificDetailsData, 
  updateClientSpecificDetailsData, 
  loadTrackerData, 
  updateTrackerData, 
  loadProjectListData, 
  updateProjectListData 
} from "./excel.js";

import { drive as googleDrive, addGoogleSheetClient, deleteGoogleSheetClient } from "./googleSheetWrite.js";
import { startScheduler } from "./sheetScheduler.js";

// --- DATE-FNS IMPORTS ---
import { 
  getISOWeek, 
  startOfISOWeek, 
  endOfISOWeek, 
  format, 
  setISOWeek, 
  setISOWeekYear, 
  isSameWeek, 
  addWeeks, 
  subWeeks, 
  subMonths, 
  isAfter, 
  isValid
} from "date-fns";

dotenv.config();
const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

const PORT = process.env.PORT || 4000;

// Configure Multer storage to save files to the public directory
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = path.join(process.cwd(), 'client', 'public');
    fs.mkdirSync(uploadDir, { recursive: true });
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    const nameFromUrl = req.params.name;
    const clientName = (nameFromUrl || (req.body.clientName || 'unknown').replace(/\s/g, '-'));
    cb(null, `${clientName}.png`);
  }
});
const upload = multer({ storage: storage });
app.use(express.static(path.join(process.cwd(), 'client', 'public'))); 

// Check for Master Sheet ID
const MASTER_SHEET_ID = process.env.SOURCES ? process.env.SOURCES.split(',')[0] : null;
if (!MASTER_SHEET_ID) {
  throw new Error("No master sheet configured. Set SOURCES in .env with your master sheet ID.");
}

const DATA_SOURCES_ARRAY = process.env.SOURCES ? process.env.SOURCES.split(',') : [];

// Cache for loaded data by week
const dataCache = new Map();
const productUpdateCache = new Map();
const clientSpecificDetailsCache = new Map();
const trackerCache = new Map();
const plCache = new Map();

// --- UTILITY: Load Data for Specific Week ---
async function loadWeekData(week) {
  const cacheKey = week;
  if (dataCache.has(cacheKey)) {
    return dataCache.get(cacheKey);
  }

  try {
    const data = await loadData(DATA_SOURCES_ARRAY, week);
    dataCache.set(cacheKey, data);
    // Cache for 5 minutes
    setTimeout(() => {
      dataCache.delete(cacheKey);
    }, 5 * 60 * 1000);
    
    return data;
  } catch (e) {
    console.error(`Error loading data for week ${week}:`, e);
    throw e;
  }
}

// --- MIDDLEWARE: Add Week Data to Request ---
app.use(async (req, res, next) => {
  try {
    const today = new Date();
    const currentWeekNum = getISOWeek(today);
    const currentYear = today.getFullYear();
    
    // Default to ISO format: YYYY-WXX
    const defaultWeek = `${currentYear}-W${String(currentWeekNum).padStart(2, '0')}`;
    const week = req.query.week || req.body.week || defaultWeek;

    if (req.path !== '/api/weeks' && !req.path.startsWith('/api/upload-logo')) {
      req.db = await loadWeekData(week);
    }
    req.selectedWeek = week;
    next();
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: "Failed to load data: " + e.message });
  }
});

// ============================================================================
//  API: GET AVAILABLE WEEKS (With Labels & Filtering)
// ============================================================================
app.get("/api/weeks", (req, res) => {
  const today = new Date();
  const threeMonthsAgo = subMonths(today, 3);
  let allWeeks = [];

  // Helper to process a specific year's variables from .env
  const processYear = (year) => {
    const sourceVar = process.env[`WEEKLY_SOURCES_${year}`];
    if (!sourceVar) return;

    const ids = sourceVar.split(',');
    ids.forEach((id, index) => {
      const weekNum = index + 1; // Assuming first ID is Week 1
      const weekValue = `${year}-W${String(weekNum).padStart(2, '0')}`;
      
      const dateHelper = setISOWeek(setISOWeekYear(new Date(), year), weekNum);
      const start = startOfISOWeek(dateHelper);
      const end = endOfISOWeek(dateHelper);

      if (!isValid(start) || !isValid(end)) return;

      allWeeks.push({
        value: weekValue,
        weekNum: weekNum,
        year: year,
        start: start,
        end: end,
        originalLabel: `Week ${weekNum}`
      });
    });
  };

  processYear(2025);
  processYear(2026);

  // --- CHANGED SORT ORDER: ASCENDING (Oldest First -> Upcoming Last) ---
  allWeeks.sort((a, b) => a.start - b.start); 

  // Apply Logic: Previous / Current / Upcoming / Filter (3 Months)
  const finalOptions = allWeeks.reduce((acc, weekObj) => {
    let label = "";
    let isCurrent = false;

    // 1. Determine Special Status
    if (isSameWeek(weekObj.start, today, { weekStartsOn: 1 })) {
      label = "Current";
      isCurrent = true;
    } else if (isSameWeek(weekObj.start, subWeeks(today, 1), { weekStartsOn: 1 })) {
      label = "Previous";
    } else if (isSameWeek(weekObj.start, addWeeks(today, 1), { weekStartsOn: 1 })) {
      label = "Upcoming";
    } else {
      label = weekObj.originalLabel; // "Week 52", etc.
    }

    // 2. Format Date Range
    const dateRange = `(${format(weekObj.start, 'MMM d, yyyy')} - ${format(weekObj.end, 'MMM d, yyyy')})`;
    const fullLabel = `${label} ${dateRange}`;

    // 3. Filter Logic
    const isSpecial = label === "Current" || label === "Previous" || label === "Upcoming";
    const isInWindow = isAfter(weekObj.end, threeMonthsAgo);

    if (isSpecial || isInWindow) {
      acc.push({
        value: weekObj.value,
        label: fullLabel,
        isCurrent: isCurrent
      });
    }

    return acc;
  }, []);

  res.json(finalOptions);
});

// List customers
app.get("/api/customers", (req, res) => {
  res.json(listCustomers(req.db));
});

// Get all data for a week
app.get("/api/data", (req, res) => {
  res.json(req.db);
});

// API endpoint to handle the logo upload
app.post("/api/upload-logo", upload.single('logo'), (req, res) => {
  if (req.file) {
    const publicPath = `/${req.file.filename}`;
    res.json({ success: true, logoUrl: publicPath });
  } else {
    res.status(400).json({ success: false, error: "No file uploaded." });
  }
});

// ADD NEW CUSTOMER
app.post("/api/customers", async (req, res) => {
  const { customerName, customerData, color } = req.body;
  const week = req.query.week;

  const trimmedName = (customerName || "").toString().trim();
  if (!trimmedName || trimmedName === 'undefined') {
    return res.status(400).json({ error: "Invalid customerName" });
  }

  if (req.db[trimmedName]) {
    return res.status(409).json({ error: `Client "${trimmedName}" already exists.` });
  }

  const weeklySource = getWeeklySource(DATA_SOURCES_ARRAY, week);
  const targetSheets = [MASTER_SHEET_ID];
  if (weeklySource) targetSheets.push(weeklySource);

  console.log(`Adding client ${trimmedName} to Master and Week: ${week}`);

  try {
    const allowedCategories = ['Client', 'Sycamore', 'Sycamore and Client'];
    const filteredData = {};
    for (const cat of allowedCategories) {
      if (Array.isArray(customerData[cat])) filteredData[cat] = customerData[cat];
    }

    const addPromises = targetSheets.map(sheetId =>
      addGoogleSheetClient(sheetId, trimmedName, filteredData)
    );
    await Promise.all(addPromises);

    if (color) {
      await updateGoogleSheetCell(MASTER_SHEET_ID, trimmedName, 'Background', color);
    }
    
    dataCache.clear();
    productUpdateCache.clear();
    clientSpecificDetailsCache.clear();
    trackerCache.clear();

    res.json({ ok: true, message: `Client ${trimmedName} added.` });
  } catch (e) {
    console.error("Error adding new client:", e);
    res.status(500).json({ error: "Failed to add new client: " + e.message });
  }
});

app.put("/api/customers/:name/logo", upload.single('logo'), async (req, res) => {
  const { name } = req.params;
  if (!req.file) return res.status(400).json({ error: "No logo file." });

  const logoUrl = `/${req.file.filename}`;

  try {
    await updateGoogleSheetCell(MASTER_SHEET_ID, name, 'Logo', logoUrl);
    dataCache.clear(); 
    res.json({ ok: true, message: `Logo updated.`, logoUrl: logoUrl });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.put("/api/customers/:name/background", async (req, res) => {
  const { name } = req.params;
  const { color } = req.body; 

  if (!color) return res.status(400).json({ error: "Missing color" });

  try {
    await updateGoogleSheetCell(MASTER_SHEET_ID, name, 'Background', color);
    dataCache.clear();
    res.json({ ok: true });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// --- SUB-DATA ENDPOINTS ---

app.get("/api/customers/:name/product-update", async (req, res) => {
  const name = req.params.name;
  const week = req.selectedWeek;

  if (!req.db[name]) return res.status(404).json({ error: "Customer not found in main sheet." });

  const cacheKey = `${week}:${name}`;
  if (productUpdateCache.has(cacheKey)) return res.json({ data: productUpdateCache.get(cacheKey) });

  try {
    const weeklySource = getWeeklySource(DATA_SOURCES_ARRAY, week);
    if (!weeklySource) return res.status(404).json({ error: `No source for ${week}` });

    const data = await loadProductUpdateData(weeklySource, name);
    productUpdateCache.set(cacheKey, data);
    setTimeout(() => productUpdateCache.delete(cacheKey), 5 * 60 * 1000);

    res.json({ data });
  } catch (e) {
    if (e.message.includes("not found")) return res.status(404).json({ error: "No data." });
    return res.status(500).json({ error: e.message });
  }
});

app.put("/api/customers/:name/product-update", async (req, res) => {
  const name = req.params.name;
  const week = req.selectedWeek;
  
  if (!req.db[name]) return res.status(404).json({ error: "Customer not found" });
  
  const weeklySource = getWeeklySource(DATA_SOURCES_ARRAY, week);
  if (!weeklySource) return res.status(400).json({ error: `No source for ${week}` });

  try {
    await updateProductUpdateData(weeklySource, name, req.body);
    productUpdateCache.delete(`${week}:${name}`);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/customers/:name/client-specific-details", async (req, res) => {
  const name = req.params.name;
  const week = req.selectedWeek;
  
  if (!req.db[name]) return res.status(404).json({ error: "Customer not found" });

  const cacheKey = `${week}:${name}`;
  if (clientSpecificDetailsCache.has(cacheKey)) return res.json({ data: clientSpecificDetailsCache.get(cacheKey) });

  try {
    const weeklySource = getWeeklySource(DATA_SOURCES_ARRAY, week);
    if (!weeklySource) return res.status(404).json({ error: `No source` });

    const data = await loadClientSpecificDetailsData(weeklySource, name);
    clientSpecificDetailsCache.set(cacheKey, data);
    setTimeout(() => clientSpecificDetailsCache.delete(cacheKey), 5 * 60 * 1000);
    res.json({ data });
  } catch (e) {
    if (e.message.includes("not found")) return res.status(404).json({ error: "No data." });
    return res.status(500).json({ error: e.message });
  }
});

app.put("/api/customers/:name/client-specific-details", async (req, res) => {
  const name = req.params.name;
  const week = req.selectedWeek;
  if (!req.db[name]) return res.status(404).json({ error: "Customer not found" });

  const weeklySource = getWeeklySource(DATA_SOURCES_ARRAY, week);
  if (!weeklySource) return res.status(400).json({ error: `No source` });

  try {
    await updateClientSpecificDetailsData(weeklySource, name, req.body);
    clientSpecificDetailsCache.delete(`${week}:${name}`);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/customers/:name/tracker", async (req, res) => {
  const name = req.params.name;
  const year = req.query.year || new Date().getFullYear();
  const cacheKey = `${name}:${year}`;
  if (trackerCache.has(cacheKey)) return res.json({ data: trackerCache.get(cacheKey) });

  try {
    const data = await loadTrackerData(MASTER_SHEET_ID, name, year);
    trackerCache.set(cacheKey, data);
    setTimeout(() => trackerCache.delete(cacheKey), 5 * 60 * 1000);
    res.json({ data });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.put("/api/customers/:name/tracker", async (req, res) => {
  const name = req.params.name;
  const { date, content } = req.body;
  if (!date || content === undefined) return res.status(400).json({ error: "Missing date/content" });

  try {
    // Prepending an apostrophe to the date string forces Google Sheets
    // to treat it as a literal string and not convert it to a serial date number.
    const dateAsString = `'${date}`;
    await updateTrackerData(MASTER_SHEET_ID, name, dateAsString, content);
    trackerCache.clear(); 
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/customers/:name/project-list", async (req, res) => {
  const name = req.params.name;
  const cacheKey = `${name}:all-years`;
  if (plCache.has(cacheKey)) return res.json({ data: plCache.get(cacheKey) });

  try {
    const data = await loadProjectListData(MASTER_SHEET_ID, name);
    plCache.set(cacheKey, data);
    setTimeout(() => plCache.delete(cacheKey), 5 * 60 * 1000);
    res.json({ data });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.put("/api/customers/:name/project-list", async (req, res) => {
  const name = req.params.name;
  const { year, content } = req.body;
  if (!year || content === undefined) return res.status(400).json({ error: "Missing year/content" });

  try {
    await updateProjectListData(MASTER_SHEET_ID, name, year, content);
    plCache.delete(`${name}:all-years`);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/customers/:name", async (req, res) => {
  const name = req.params.name;
  const weeklyData = req.db[name];
  if (!weeklyData) return res.status(404).json({ error: "Customer not found" });

  const masterData = await loadData([MASTER_SHEET_ID], 'master');
  const masterCustomerData = masterData[name];

  const documentUrls = {};
  const extractUrls = (customerData) => {
    if (!customerData) return;
    for (const category in customerData) {
      if (Array.isArray(customerData[category])) {
        customerData[category].forEach(item => {
          const parts = item.split(':');
          if (parts.length > 1) {
            const key = parts[0].trim();
            let value = parts.slice(1).join(':').trim();
            let extractedUrl = '';
            const linkMatch = value.match(/\[LINK:\s*(https?:\/\/[^\]]+)\]/i);
            if (linkMatch && linkMatch[1]) extractedUrl = linkMatch[1];
            else if (value.startsWith('http')) extractedUrl = value;
            if (extractedUrl) documentUrls[key] = extractedUrl;
          }
        });
      }
    }
  };

  extractUrls(weeklyData);
  extractUrls(masterCustomerData);
  res.json({
    ...weeklyData,
    _meta: {
      week: req.selectedWeek,
      timestamp: new Date().toISOString(),
      documentUrls: documentUrls
    }
  });
});

app.get("/api/search", (req, res) => {
  const q = (req.query.q || "").toString().trim();
  if (!q) return res.json({});

  const lowerCaseQuery = q.toLowerCase();
  const searchResults = {};
  const wholeWordRegex = new RegExp(`\\b${q}\\b`, 'i');

  const dataToSearch = req.db._weeklyUpdates ? { ...req.db } : req.db;
  delete dataToSearch._weeklyUpdates;

  for (const [cust, categories] of Object.entries(dataToSearch)) {
    const customerMatches = {};
    for (const [categoryName, items] of Object.entries(categories)) {
      if (categoryName.toLowerCase().includes(lowerCaseQuery)) {
        customerMatches[categoryName] = items;
        continue;
      }
      if (!Array.isArray(items)) continue;
      const matchingItems = items.filter(item => 
        wholeWordRegex.test(item) || item.toLowerCase().includes(lowerCaseQuery)
      );
      if (matchingItems.length > 0) customerMatches[categoryName] = matchingItems;
    }
    if (Object.keys(customerMatches).length > 0) searchResults[cust] = customerMatches;
  }

  res.json({ results: searchResults, customers: Object.keys(searchResults) });
});

app.put("/api/customers/:name/data", async (req, res) => {
  const name = req.params.name;
  const newData = req.body;
  const week = req.selectedWeek;
  if (!req.db[name]) return res.status(404).json({ error: "Customer not found" });

  let dataSource;
  if (week === 'master') dataSource = MASTER_SHEET_ID;
  else dataSource = getWeeklySource(DATA_SOURCES_ARRAY, week);

  if (!dataSource) return res.status(400).json({ error: `No editable data source` });

  try {
    await updateGoogleSheetData(dataSource, name, newData);
    dataCache.delete(week); 
    if (week === 'master') dataCache.clear();
    return res.json({ ok: true });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

app.post("/api/cache/clear", (req, res) => {
  dataCache.clear();
  productUpdateCache.clear();
  clientSpecificDetailsCache.clear();
  trackerCache.clear();
  plCache.clear();
  res.json({ ok: true });
});

app.get("/api/weekly-update", async (req, res) => {
  const week = req.query.week;
  // Bypass the main middleware cache to ensure we always get the latest weekly update.
  const db = await loadWeekData(week);
  const updates = db._weeklyUpdates || {};
  res.json({ text: updates[week] || '' });
});

app.post("/api/weekly-update", async (req, res) => {
  const { text } = req.body;
  const week = req.query.week;
  const weeklySourceId = getWeeklySource(DATA_SOURCES_ARRAY, week);

  if (!weeklySourceId) return res.status(400).json({ error: `No source` });

  try {
    await updateWeeklyUpdate(weeklySourceId, week, text);
    dataCache.clear();
    res.json({ ok: true });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.delete("/api/customers/:name", async (req, res) => {
  const { name } = req.params;
  try {
    const deletePromises = DATA_SOURCES_ARRAY.map(sheetId => deleteGoogleSheetClient(sheetId, name));
    await Promise.all(deletePromises);
    dataCache.clear();
    productUpdateCache.clear();
    clientSpecificDetailsCache.clear();
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ============================================================================
//  EXPORT ENDPOINT - Get all client data for export
// ============================================================================
app.get("/api/customers/:name/export", async (req, res) => {
  const name = req.params.name;
  const week = req.query.week || req.selectedWeek;

  try {
    // Get main customer data
    const weeklyData = req.db[name];
    if (!weeklyData) return res.status(404).json({ error: "Customer not found" });

    const masterData = await loadData([MASTER_SHEET_ID], 'master');
    const masterCustomerData = masterData[name];

    // Merge weekly and master data (weekly takes precedence)
    const mergedData = {
      "Client": [],
      "Sycamore": [],
      "Sycamore and Client": []
    };

    // Helper to merge categories
    const mergeCategory = (target, source) => {
      if (!source) return;
      source.forEach(item => {
        const key = item.split(':')[0].trim();
        const exists = target.find(t => t.split(':')[0].trim() === key);
        if (exists) {
          // Replace with newer data (weekly) if available
          const idx = target.indexOf(exists);
          target[idx] = item;
        } else {
          target.push(item);
        }
      });
    };

    if (weeklyData) {
      mergeCategory(mergedData["Client"], weeklyData["Client"]);
      mergeCategory(mergedData["Sycamore"], weeklyData["Sycamore"]);
      mergeCategory(mergedData["Sycamore and Client"], weeklyData["Sycamore and Client"]);
    }

    if (masterCustomerData) {
      mergeCategory(mergedData["Client"], masterCustomerData["Client"]);
      mergeCategory(mergedData["Sycamore"], masterCustomerData["Sycamore"]);
      mergeCategory(mergedData["Sycamore and Client"], masterCustomerData["Sycamore and Client"]);
    }

    // Get product update data
    let productUpdateData = {};
    try {
      const weeklySource = getWeeklySource(DATA_SOURCES_ARRAY, week);
      if (weeklySource) {
        productUpdateData = await loadProductUpdateData(weeklySource, name);
      }
    } catch (e) {
      console.warn(`Could not load product update data for ${name}:`, e.message);
    }

    // Get client specific details
    let clientSpecificData = {};
    try {
      const weeklySource = getWeeklySource(DATA_SOURCES_ARRAY, week);
      if (weeklySource) {
        clientSpecificData = await loadClientSpecificDetailsData(weeklySource, name);
      }
    } catch (e) {
      console.warn(`Could not load client specific data for ${name}:`, e.message);
    }

    // Get tracker data (all years)
    let trackerData = {};
    try {
      trackerData = await loadTrackerData(MASTER_SHEET_ID, name, new Date().getFullYear());
    } catch (e) {
      console.warn(`Could not load tracker data for ${name}:`, e.message);
    }

    // Get project list data (all years)
    let projectListData = {};
    try {
      projectListData = await loadProjectListData(MASTER_SHEET_ID, name);
    } catch (e) {
      console.warn(`Could not load project list data for ${name}:`, e.message);
    }

    // Extract document URLs
    const documentUrls = {};
    const extractUrls = (customerData) => {
      if (!customerData) return;
      for (const category in customerData) {
        if (Array.isArray(customerData[category])) {
          customerData[category].forEach(item => {
            const parts = item.split(':');
            if (parts.length > 1) {
              const key = parts[0].trim();
              let value = parts.slice(1).join(':').trim();
              let extractedUrl = '';
              const linkMatch = value.match(/\[LINK:\s*(https?:\/\/[^\]]+)\]/i);
              if (linkMatch && linkMatch[1]) extractedUrl = linkMatch[1];
              else if (value.startsWith('http')) extractedUrl = value;
              if (extractedUrl) documentUrls[key] = extractedUrl;
            }
          });
        }
      }
    };

    extractUrls(mergedData);

    res.json({
      customerName: name,
      week: week,
      exportedAt: new Date().toISOString(),
      home: {
        clientInfo: mergedData["Client"],
        stakeholders: mergedData["Sycamore"].filter(item => {
          const key = item.split(':')[0].toLowerCase();
          return ['csm', 'lead ba', 'production operation poc', 'support lead', 'technical lead', 'specialist in sycamore informatics', 'sme', 'support team', 'escalation matrix'].some(k => key.includes(k));
        }),
        versions: mergedData["Sycamore"].filter(item => {
          const key = item.split(':')[0].toLowerCase();
          return ['sycamore informatics product', 'add-on modules'].some(k => key.includes(k));
        }),
        users: mergedData["Client"].filter(item => {
          const key = item.split(':')[0].toLowerCase();
          return ['number of active users', 'number of full users', 'number of read only users', 'number of tlf users'].some(k => key.includes(k));
        }),
        sentiment: mergedData["Client"].find(item => item.toLowerCase().includes('customer sentiment score'))
      },
      sycamore: mergedData["Sycamore"],
      sycamoreAndClient: mergedData["Sycamore and Client"],
      cft: {
        productUpdates: productUpdateData,
        clientSpecificDetails: clientSpecificData
      },
      tracker: trackerData,
      projectList: projectListData,
      documents: documentUrls
    });
  } catch (e) {
    console.error(`Export error for ${name}:`, e);
    res.status(500).json({ error: "Failed to export client data: " + e.message });
  }
});

app.get("*", (req, res) => {
  res.sendFile(path.resolve(process.cwd(), 'client', 'public', 'index.html'));
});

app.listen(PORT, async () => {
  console.log(`\n🚀 Server listening on http://localhost:${PORT}`);
  startScheduler();
});