// googleClient.js
import { google } from "googleapis";
import fs, { promises as fsp } from "fs";
import path from 'path';
import { fileURLToPath } from 'url';

// Utility to get __dirname in ES Modules
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const credPath = path.join(__dirname, '..', 'credentials.json');

// Cache for the authorized client instances (based on required scopes)
const clientCache = new Map();

/**
 * Loads credentials and returns an authorized Google client object for a given scope.
 * This is the central authentication point for the backend.
 * @param {string[]} scopes - Array of required scopes (e.g., ['https://www.googleapis.com/auth/spreadsheets']).
 * @returns {Promise<google.auth.GoogleAuth | google.auth.JWT>}
 */
async function getAuthClient(scopes) {
    const cacheKey = `auth-${scopes.sort().join(',')}`;
    if (clientCache.has(cacheKey)) {
        return clientCache.get(cacheKey);
    }

    if (!fs.existsSync(credPath)) {
        throw new Error('credentials.json not found at ' + credPath);
    }
    const credsContent = await fsp.readFile(credPath, 'utf8');
    const creds = JSON.parse(credsContent);
    const rawKey = creds.private_key || creds.privateKey || '';
    const normalizedKey = rawKey.replace(/\\n/g, '\n').trim();

    if (!normalizedKey) {
        throw new Error('Service account private key is empty.');
    }

    const auth = new google.auth.JWT({
        email: creds.client_email,
        key: normalizedKey,
        scopes: scopes,
    });

    try {
        await auth.authorize();
    } catch (err) {
        console.error('Service account authorization failed:', err.message || err);
        throw new Error('Service account authorization failed: ' + (err.message || err));
    }

    clientCache.set(cacheKey, auth);
    return auth;
}


/**
 * Returns an authorized Google Sheets client (v4).
 * @param {string[]} scopes - Array of required scopes. Default: ['https://www.googleapis.com/auth/spreadsheets'].
 * @returns {Promise<google.sheets_v4.Sheets>}
 */
export async function getGoogleSheetsClient(scopes = ['https://www.googleapis.com/auth/spreadsheets']) {
    const auth = await getAuthClient(scopes);
    return google.sheets({ version: 'v4', auth });
}

/**
 * Returns an authorized Google Drive client (v3).
 * @param {string[]} scopes - Array of required scopes. Default: ['https://www.googleapis.com/auth/drive.readonly', 'https://www.googleapis.com/auth/drive.file'].
 * @returns {Promise<google.drive_v3.Drive>}
 */
export async function getGoogleDriveClient(scopes = ['https://www.googleapis.com/auth/drive.readonly', 'https://www.googleapis.com/auth/drive.file']) {
    const auth = await getAuthClient(scopes);
    return google.drive({ version: 'v3', auth });
}