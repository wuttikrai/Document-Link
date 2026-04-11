A simple and efficient Google Apps Script solution to list all document URLs from a specific Google Drive folder and record them into a Google Sheet. This tool helps you manage large volumes of files by centralizing their names and links in one spreadsheet.

🚀 Getting Started
Follow these steps to set up the script:

1. Prepare Your Google Sheet
Create a new Google Sheet in your Google Drive.

(Optional) Rename the sheet tab (e.g., "File List") to match your configuration.

2. Copy the Script
Open the file named getFilesInFolder provided in this repository.

Copy the entire code block.

3. Open Google Apps Script
In your Google Sheet, go to the top menu.

Select Extensions (ส่วนขยาย) > Apps Script.

4. Configure the Script
Paste the copied code into the script editor.

Update the Configuration: Locate the variables in the code to specify:

The Folder ID you want to scan.

The Sheet Name where the data should be saved.

Save the project (e.g., "Drive Indexer").

5. Run the Script
Click the Run button in the Apps Script editor.

Authorization: The first time you run it, Google will ask for permission to access your Drive and Sheets. Grant the necessary permissions.

Your Google Sheet will now be populated with the file names and URLs.
