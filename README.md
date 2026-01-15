# Sycamore Customer Dashboard (SCD)

A full-stack customer dashboard application that acts as a **"Menu-Card"** for client data. 
It features a **React** frontend for visualization and a **Node.js/Express** backend that syncs data in real-time with **Google Sheets**.

## 📋 Prerequisites

Before running this project, ensure you have the following installed on your computer:

1.  **Visual Studio Code (VS Code):** A code editor. [Download here](https://code.visualstudio.com/).
2.  **Node.js (v18 or higher):** Required to run the server and client. [Download here](https://nodejs.org/).
3.  **Git:** To clone the repository. [Download here](https://git-scm.com/).

---

## 🚀 Setup Guide (From Scratch)

Since this repository was uploaded without the heavy dependency folders (`node_modules`) to keep the project lightweight, **you must reinstall them** before the app will work.

### 1. Clone the Repository
Open your terminal (Command Prompt or Terminal in VS Code) and run this command to download the code:

```bash
git clone <YOUR_REPO_URL>
```
Then, move into the project folder:

```bash
cd sycamore-menu-dashboard
```

2. Install Server Dependencies
You need to install the libraries for the backend (Server). Run this command to enter the server folder:

```bash
cd server
```
Then run the install command:

```bash
npm install
```

3. Install Client Dependencies
Now you need to install the libraries for the frontend (Client). First, move to the client folder:

```bash
cd client
```

Then run the install command:

```bash
npm install
```

### 🔑 Google Cloud Configuration (Critical)

Since this app reads/writes to Google Sheets, you must create your own credentials. These are secret keys and are not included in this repository.

#### Step A: Create a Service Account

1. Go to the Google Cloud Console.

2. Create a New Project (e.g., "Sycamore Dashboard").

3. Go to APIs & Services > Library and enable:

- Google Sheets API

- Google Drive API

4. Go to APIs & Services > Credentials and click Create Credentials > Service Account.

- Name it (e.g., dashboard-backend) and click Done.

#### Step B: Generate the Key (credentials.json)

1. Click on the newly created Service Account email (e.g., dashboard-backend@...).

2. Go to the Keys tab > Add Key > Create new key.

3. Select JSON and create. A file will download to your computer.

4. Rename this file to credentials.json.

5. Move this file into the server/ folder of this project.

#### Step C: Share Your Sheets

1. Open the Google Sheet you want to use as your database.

2. Click the Share button.

3. Paste the Service Account Email (found in your Google Cloud Console) and give it Editor access.

4. Copy the Sheet ID from the URL (the long string between /d/ and /edit).

### ⚙️ Environment Variables

You need to tell the code where your servers are and which Sheets to load.

1. Server Configuration
Create a file named .env inside the server/ folder and paste this content:

```Code snippet
PORT=4000
# Comma-separated list of Sheet IDs. First ID is the MASTER sheet.
SOURCES=your_master_sheet_id_here,your_weekly_sheet_id_here

# For Weekly Update Logic (Year-specific IDs)
# Format: Week1_ID,Week2_ID,Week3_ID...
WEEKLY_SOURCES_2025=id_for_week_1,id_for_week_2,...

# Google Drive Folder ID (for uploading Excel conversions/images)
UPLOAD_FOLDER_ID=your_drive_folder_id_here
```

2. Client Configuration
Create a file named .env inside the client/ folder and paste this content:

```Code snippet
VITE_API_BASE=http://localhost:4000
# Secure password for the login screen
VITE_ADMIN_PASSWORD=password@24
```

## 🏃‍♂️ How to Run

You need to run the Server and Client at the same time. Open two separate terminals in VS Code.

Terminal 1: Start the Backend
Navigate to the server folder:

```bash
cd server
```
Start the server:

```bash
npm run dev
```
You should see: 🚀 Server listening on http://localhost:4000

Terminal 2: Start the Frontend
Navigate to the client folder:

```bash
cd client
```
Start the client:

```bash
npm run dev
```
You should see: Local: http://localhost:5173

---

## 🛠️ Usage Notes

1. Login: Open http://localhost:5173 in your browser. Log in with admin and the password you set in your .env file.

2. Excel vs. Google Sheets: The logic currently pulls data exclusively from the Google Sheets connected in your .env.

3. Search: The global search bar filters customers based on data in your active Google Sheets.

4. Images: Place client logo images (PNG) in client/public/. Name them exactly as the client appears in the sheet (e.g., Novo Nordisk.png).

---

## 🛑 Troubleshooting

1. "Credentials file not found": Ensure server/credentials.json exists and is named correctly.

2. "Sheet not found": Ensure you shared the Google Sheet with the Service Account email address.

3. "Connection Refused": Ensure the server is running on port 4000 before starting the client.

4. "Module not found": If you see errors about missing modules, make sure you ran npm install in both folders.
