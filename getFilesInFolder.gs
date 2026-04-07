function getFilesInFolder() {
  // =============================================
  // CONFIGURATION - Change this to your folder ID
  // =============================================
  const FOLDER_ID = "ลิ้งโฟลเดอร์ที่ต้องการจะใช้งาน"; // Replace with your Google Drive folder ID
  const OUTPUT_SHEET_NAME = "ชื่อชีทย่อย // เมื่อใช้งานแล้ว แนะนำให้เปลี่ยนชื่อ มิฉะนั้นโปรแกรมจะเขียนทับชื่อชีทเดิม"; 
  // =============================================

  // --- Clean the folder ID (removes spaces or accidental URL parts) ---
  const cleanId = extractFolderId(FOLDER_ID);
  if (!cleanId) {
    SpreadsheetApp.getUi().alert("❌ Invalid Folder ID. Please check your FOLDER_ID value.");
    return;
  }

  let folder;
  try {
    folder = DriveApp.getFolderById(cleanId);
  } catch (e) {
    SpreadsheetApp.getUi().alert(
      "❌ Cannot access folder.\n\n" +
      "Reason: " + e.message + "\n\n" +
      "Tips:\n" +
      "1. Make sure the Folder ID is correct.\n" +
      "2. Make sure you have access to the folder.\n" +
      "3. Folder ID used: " + cleanId
    );
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Create or clear the output sheet
  let sheet = ss.getSheetByName(OUTPUT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(OUTPUT_SHEET_NAME);
  } else {
    sheet.clearContents();
  }

  // Set headers
  const headers = ["File Name", "File URL", "File Type", "Size (KB)", "Last Modified", "Folder Path"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#4285F4")
    .setFontColor("#FFFFFF");

  // Start recursive scraping
  const results = [];
  scrapeFolder(folder, folder.getName(), results);

  // Write all results to sheet
  if (results.length > 0) {
    sheet.getRange(2, 1, results.length, headers.length).setValues(results);

    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }

    // Make URLs clickable
    for (let i = 0; i < results.length; i++) {
      const url = results[i][1];
      const cell = sheet.getRange(i + 2, 2);
      cell.setFormula(`=HYPERLINK("${url}", "Open File")`);
    }
  }

  SpreadsheetApp.getUi().alert(`✅ Done! Found ${results.length} files in folder: "${folder.getName()}"`);
}


/**
 * Extracts a clean folder ID from either a raw ID or a full Drive URL
 * Handles formats like:
 *   - 1A2B3C4D5E6F (raw ID)
 *   - https://drive.google.com/drive/folders/1A2B3C4D5E6F
 *   - https://drive.google.com/drive/u/0/folders/1A2B3C4D5E6F
 */
function extractFolderId(input) {
  if (!input || input.trim() === "" || input === "YOUR_FOLDER_ID_HERE") return null;

  input = input.trim();

  // If it's a full URL, extract the ID
  const urlMatch = input.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (urlMatch) return urlMatch[1];

  // If it looks like a raw ID (alphanumeric + dash + underscore)
  const rawMatch = input.match(/^[a-zA-Z0-9_-]+$/);
  if (rawMatch) return input;

  return null;
}


/**
 * Recursively scrape files from a folder and its subfolders
 */
function scrapeFolder(folder, folderPath, results) {
  try {
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      try {
        results.push([
          file.getName(),
          file.getUrl(),
          file.getMimeType().split(".").pop().split("/").pop(),
          (file.getSize() / 1024).toFixed(2),
          Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),
          folderPath
        ]);
      } catch (fileErr) {
        // Skip unreadable files silently
        results.push(["⚠️ Unreadable file", "", "", "", "", folderPath]);
      }
    }

    // Recurse into subfolders
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      const subFolder = subFolders.next();
      scrapeFolder(subFolder, folderPath + " > " + subFolder.getName(), results);
    }
  } catch (folderErr) {
    results.push(["⚠️ Cannot read folder: " + folderPath, "", "", "", "", folderPath]);
  }
}


/**
 * Adds a custom menu to the spreadsheet
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📁 Drive Scraper")
    .addItem("Scrape File Links", "getFilesInFolder")
    .addToUi();
}
