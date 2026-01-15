# Sycamore Menu-Card Dashboard

A full-stack customer dashboard application that acts as a **"Menu-Card"** for client data. 
It features a **React** frontend for visualization and a **Node.js/Express** backend that syncs data in real-time with **Google Sheets**.

![Dashboard Screenshot](client/public/sycamore-logo.png)

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

git clone <YOUR_REPO_URL>
Then, move into the project folder:

'''Bash

cd sycamore-menu-dashboard
2. Install Server Dependencies
You need to install the libraries for the backend (Server). Run this command to enter the server folder:

'''Bash

cd server
Then run the install command:

'''Bash

npm install
3. Install Client Dependencies
Now you need to install the libraries for the frontend (Client). First, move to the client folder:

'''Bash

cd ../client
Then run the install command:

'''Bash

npm install
🔑 Google Cloud Configuration (Critical)
Since this app reads/writes to Google Sheets, you must create your own credentials. These are secret keys and are not included in this repository.

Step A: Create a Service Account
Go to the Google Cloud Console.

Create a New Project (e.g., "Sycamore Dashboard").

Go to APIs & Services > Library and enable:

Google Sheets API

Google Drive API

Go to APIs & Services > Credentials and click Create Credentials > Service Account.

Name it (e.g., dashboard-backend) and click Done.

Step B: Generate the Key (credentials.json)
Click on the newly created Service Account email (e.g., dashboard-backend@...).

Go to the Keys tab > Add Key > Create new key.

Select JSON and create. A file will download to your computer.

Rename this file to credentials.json.

Move this file into the server/ folder of this project.

Step C: Share Your Sheets
Open the Google Sheet you want to use as your database.

Click the Share button.

Paste the Service Account Email (found in your Google Cloud Console) and give it Editor access.

Copy the Sheet ID from the URL (the long string between /d/ and /edit).

⚙️ Environment Variables
You need to tell the code where your servers are and which Sheets to load.

1. Server Configuration
Create a file named .env inside the server/ folder and paste this content:

'''Code snippet

PORT=4000
# Comma-separated list of Sheet IDs. First ID is the MASTER sheet.
SOURCES=your_master_sheet_id_here,your_weekly_sheet_id_here

# For Weekly Update Logic (Year-specific IDs)
# Format: Week1_ID,Week2_ID,Week3_ID...
WEEKLY_SOURCES_2025=id_for_week_1,id_for_week_2,...

# Google Drive Folder ID (for uploading Excel conversions/images)
UPLOAD_FOLDER_ID=your_drive_folder_id_here
2. Client Configuration
Create a file named .env inside the client/ folder and paste this content:

'''Code snippet

VITE_API_BASE=http://localhost:4000
# Secure password for the login screen
VITE_ADMIN_PASSWORD=password@24
🏃‍♂️ How to Run
You need to run the Server and Client at the same time. Open two separate terminals in VS Code.

Terminal 1: Start the Backend
Navigate to the server folder:

'''Bash

cd server
Start the server:

'''Bash

npm run dev
You should see: 🚀 Server listening on http://localhost:4000

Terminal 2: Start the Frontend
Navigate to the client folder:

'''Bash

cd client
Start the client:

'''Bash

npm run dev
You should see: Local: http://localhost:5173

🛠️ Usage Notes
Login: Open http://localhost:5173 in your browser. Log in with admin and the password you set in your .env file.

Excel vs. Google Sheets: The logic currently pulls data exclusively from the Google Sheets connected in your .env.

Search: The global search bar filters customers based on data in your active Google Sheets.

Images: Place client logo images (PNG) in client/public/. Name them exactly as the client appears in the sheet (e.g., Novo Nordisk.png).

🛑 Troubleshooting
"Credentials file not found": Ensure server/credentials.json exists and is named correctly.

"Sheet not found": Ensure you shared the Google Sheet with the Service Account email address.

"Connection Refused": Ensure the server is running on port 4000 before starting the client.

"Module not found": If you see errors about missing modules, make sure you ran npm install in both folders.
