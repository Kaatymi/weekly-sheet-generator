# 📁 Automated File Generator for Weekly Tasks

This Google Sheets automation script generates weekly folders and files based on user-defined Monday–Friday dates. It creates 4 Google Sheets files inside a Drive folder:

```
📂  6/24/2025 - 6/28/2025
├ 📄 File 1 - Main Input (e.g. "File 1 - Main Input - 6/24/2025 - 6/28/2025")
├ 📄 File 2 - Report (e.g. "File 2 - Report - 6/28/2025") → sheet: "Report WE 6/28/2025 Dest"
├ 📄 File 3 - Main Input (e.g. "File 3 - Main Input - 6/24/2025 - 6/28/2025")
├ 📄 File 4 - Report (e.g. "File 4 - Report - 6/28/2025") → sheet: "Report WE 6/28/2025 Dest"
```

---

## ✅ Features

- 📁 Auto-creates a Drive folder with the selected weekly range (Monday–Friday)
- 📄 Automatically copies templates for 2 input files and 2 report files
- 🧠 **Self-Healing Logic**: if files are missing or deleted, only the missing ones are recreated
- 🔗 Adds working links (with emojis 📄 📂) back to the source sheet
- 🔁 Dynamically inserts `IMPORTRANGE` formulas into report files
- 📝 Renames tabs automatically for clarity and linkage
- ⏱️ Visual toast alerts for progress and success

---

## ⚙️ How it works

The automation is triggered when a checkbox in column **H** of the `CreatingTable` sheet is marked `TRUE`. It uses the dates in columns **A (Monday)** and **B (Friday)** to:

1. Create a folder in Drive (if not existing)
2. Copy template files and rename based on the week
3. Rename the sheets
4. Add hyperlinks with emojis back to the tracker
5. Insert `IMPORTRANGE()` formulas to link data between Input and Report files

---

## 🔐 Authorization & Triggers

### 📎 Required Permissions
This script uses Google Drive and Spreadsheet APIs. On first run, Google will ask for access.

### 🔓 One-time Setup
You **must run this function manually once** to authorize access:
```js
function authorizeDriveAccess() {
  DriveApp.getRootFolder();
  Logger.log("Drive access authorized.");
}
```

### ⏰ Installable Trigger
To make the script respond automatically when checkboxes are edited:
1. Open your Apps Script project
2. Go to **Triggers** (⏰ icon in toolbar)
3. Add new trigger:
   - Function: `onEdit`
   - Event: `On edit`

This will ensure the automation runs any time a checkbox is clicked.

---

## 💡 Use Cases

- Weekly data report generation
- Repetitive task automation
- Template-based spreadsheet duplication
- Google Workspace optimization

---

## 📄 License
MIT License © 2025 — Kaatymi

## 🤖 AI-Assisted Notice
This project was developed with the assistance of GitHub Copilot, OpenAI Codex, and ChatGPT to accelerate scripting, improve code structure, and refine automation logic.  
All final logic, decisions, and validations were made manually by the developer.
