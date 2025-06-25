// === CONFIGURATION ===
// Replace placeholder values with your actual folder and template IDs
const MAIN_FOLDER_ID = "PASTE_YOUR_MAIN_FOLDER_ID_HERE";
const TEMPLATE_SOURCE_1 = "PASTE_TEMPLATE_ID_FOR_FILE_1";
const TEMPLATE_DEST_1 = "PASTE_TEMPLATE_ID_FOR_FILE_2";
const TEMPLATE_SOURCE_2 = "PASTE_TEMPLATE_ID_FOR_FILE_3";
const TEMPLATE_DEST_2 = "PASTE_TEMPLATE_ID_FOR_FILE_4";

function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== "WeeklyAutomation") return;

    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();

    // Trigger when checkbox in column H is checked
    if (col === 8 && range.getValue() === true) {
      showToast("Creating folder and files...", "In progress", 30);

      const monday = sheet.getRange(row, 1).getValue();
      const friday = sheet.getRange(row, 2).getValue();

      if (!(monday instanceof Date) || !(friday instanceof Date)) {
        showToast("Invalid MONDAY or FRIDAY date", "Error");
        return;
      }

      const folderName = `${formatDate(monday)} - ${formatDate(friday)}`;
      const mainFolder = DriveApp.getFolderById(MAIN_FOLDER_ID);
      let weekFolder = getFolderByName(mainFolder, folderName);
      if (!weekFolder) weekFolder = mainFolder.createFolder(folderName);

      const fileLinks = {};
      const filesToCreate = [
        { id: TEMPLATE_SOURCE_1, name: `File 1 - Main Input - ${folderName}`, col: 3, type: "source", key: "file1", sheetName: `Report WE ${formatDate(friday)}` },
        { id: TEMPLATE_DEST_1,   name: `File 2 - Report - ${formatDate(friday)}`,     col: 4, type: "dest",   key: "file1", sheetName: `Report WE ${formatDate(friday)} Dest` },
        { id: TEMPLATE_SOURCE_2, name: `File 3 - Main Input - ${folderName}`,         col: 5, type: "source", key: "file2", sheetName: `Report WE ${formatDate(friday)}` },
        { id: TEMPLATE_DEST_2,   name: `File 4 - Report - ${formatDate(friday)}`,     col: 6, type: "dest",   key: "file2", sheetName: `Report WE ${formatDate(friday)} Dest` }
      ];

      for (const file of filesToCreate) {
        let f = getFileByName(weekFolder, file.name);
        if (!f) {
          f = DriveApp.getFileById(file.id).makeCopy(file.name, weekFolder);
        }

        const newSpreadsheet = SpreadsheetApp.openById(f.getId());
        newSpreadsheet.getSheets()[0].setName(file.sheetName);

        const url = `https://docs.google.com/spreadsheets/d/${f.getId()}/edit`;
        sheet.getRange(row, file.col).setFormula(`=HYPERLINK("${url}", "Link")`);

        if (file.type === "source") fileLinks[file.key] = { url, sheetName: file.sheetName };
      }

      const destFile1 = getFileByName(weekFolder, `File 2 - Report - ${formatDate(friday)}`);
      const destFile2 = getFileByName(weekFolder, `File 4 - Report - ${formatDate(friday)}`);

      const destSheet1 = SpreadsheetApp.openById(destFile1.getId()).getSheets()[0];
      const destSheet2 = SpreadsheetApp.openById(destFile2.getId()).getSheets()[0];

      destSheet1.getRange("A8").setFormula(`=IMPORTRANGE("${fileLinks.file1.url}", "'${fileLinks.file1.sheetName}'!A8:Z")`);
      destSheet2.getRange("A8").setFormula(`=IMPORTRANGE("${fileLinks.file2.url}", "'${fileLinks.file2.sheetName}'!A8:Z")`);

      const folderUrl = `https://drive.google.com/drive/folders/${weekFolder.getId()}`;
      sheet.getRange(row, 7).setFormula(`=HYPERLINK("${folderUrl}", "Folder")`);

      showToast(`Week folder and files created: ${folderName}`, "Done", 3);
    }
  } catch (error) {
    console.warn("Silent error: " + error.message);
  }
}

// === HELPER FUNCTIONS ===

// Format date as M/D/YYYY
function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "M/d/yyyy");
}

// Display toast message in sheet UI
function showToast(message, title = "Info", timeoutSeconds = 5) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title, timeoutSeconds);
}

// Check if folder exists by name under parent
function getFolderByName(parentFolder, name) {
  const folders = parentFolder.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : null;
}

// Check if file exists by name in folder
function getFileByName(folder, name) {
  const files = folder.getFilesByName(name);
  return files.hasNext() ? files.next() : null;
}

// === ONE-TIME AUTHORIZATION FUNCTION ===
// Run this manually once to authorize Google Drive access
function authorizeDriveAccess() {
  DriveApp.getRootFolder();
  Logger.log("Drive access authorized.");
}
