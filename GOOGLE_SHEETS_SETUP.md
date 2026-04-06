# 🚀 Comprehensive Google Sheets Setup Guide

Follow these detailed steps to connect your Admission Portal to Google Sheets.

---

### Step 1: Create a Google Cloud Project
1.  Go to the [Google Cloud Console](https://console.cloud.google.com/).
2.  Click the **Project Dropdown** at the top left.
3.  Click **New Project** and name it `AdmissionPortalDB`.
4.  Click **Create**.

### Step 2: Enable Google Sheets & Drive APIs
1.  In the search bar at the top, type **"Google Sheets API"**.
2.  Select it from the results and click **Enable**.
3.  Go back to the search bar and type **"Google Drive API"**.
4.  Select it and click **Enable**.
    > [!IMPORTANT]
    > Both APIs must be enabled for the connection to work.

### Step 3: Create a Service Account
1.  Go to **IAM & Admin > Service Accounts** in the left sidebar.
2.  Click **Create Service Account** at the top.
3.  Name it `admission-bot` and click **Create and Continue**.
4.  (Optional) Skip the roles setup and click **Done**.

### Step 4: Generate your Credentials Key
1.  In the Service Accounts list, click on the **email address** of the account you just created.
2.  Go to the **Keys** tab at the top.
3.  Click **Add Key > Create New Key**.
4.  Choose **JSON** and click **Create**.
5.  A file will download to your computer. **Rename it to `credentials.json`**.
6.  Move this `credentials.json` file into your project folder: `D:\adportal\`.

### Step 5: Share your Google Sheet
1.  Create or open a Google Sheet named **`AdmissionPortal_DB`**.
2.  Open your `credentials.json` file and look for the line `"client_email": "..."`.
3.  Copy that email address (it looks like `admission-bot@...gserviceaccount.com`).
4.  In your Google Sheet, click the blue **Share** button.
5.  Paste the email address, ensure the role is set to **Editor**, and click **Send**.

### Step 6: Start the Application
1.  Ensure your `credentials.json` is in the `D:\adportal` folder.
2.  Run the application:
    -   **PowerShell**: Right-click `run.ps1` and select "Run with PowerShell" (or run `.\run.ps1` in terminal).
    -   **CMD**: Double-click `run.bat`.
3.  The app will automatically create the header row if the sheet is empty!

---

### 🛠️ Troubleshooting
-   **Error: Credentials not found**: Ensure the file is named exactly `credentials.json` and is in the root folder.
-   **Error: SpreadsheetNotFound**: Double-check that the sheet name in `app.py` (line 19) matches your Google Sheet name exactly.
-   **Error: PermissionDenied**: Ensure you Shared the sheet with the Service Account email and gave it "Editor" access.
