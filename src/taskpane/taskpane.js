/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    

    // Enrollees button
    const enrollBtn = document.getElementById("enrollees-button");
    if (enrollBtn) {
        enrollBtn.onclick = refreshEnrollmentData;
    }

    // NEW: Template button integration
    const templateBtn = document.getElementById("template-button");
    if (templateBtn) {
        templateBtn.onclick = () => {
            // Call your function: URL, Sheet Name, Mode (1 = Replace)
            const fileUrl = "https://www.tesdadvts.org/ClassRecordTemplate.xlsx";
            downloadTemplate(fileUrl, "Base60, Attendance, Midterm, FinalTerm, Gradesheet,TraineeList-F2,Advisory,InstructorSchedule", 1);
        };
    }
  }
});

async function refreshEnrollmentData() {
  const status = document.getElementById("status-message");
  status.innerText = "Connecting to Registrar Server...";

  try {
    const response = await fetch("https://node.tesdadvts.org/fl_get_enrollment");
    
    if (!response.ok) throw new Error("Server responded with error: " + response.status);
    
    const data = await response.json();

    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const sheetName = "enrollment";
      
      // 1. Check if the sheet exists
      let sheet = sheets.getItemOrNullObject(sheetName);
      await context.sync();

      // 2. If it doesn't exist (isNullObject), create it
      if (sheet.isNullObject) {
        sheet = sheets.add(sheetName);
      }

      // 3. Always clear the entire sheet to ensure a fresh overwrite
      sheet.getUsedRange().clear();

      // Headers
      const headers = [["recordid", "lastname", "firstname", "middlename", "suffix", "birthday", "gender", "address", "contact", "email", "batch", "idnumber", "uli"]];
      sheet.getRange("A1:M1").values = headers;
      sheet.getRange("A1:M1").format.font.bold = true;

      const excelRows = data.map(item => [
        item.recordid, item.lastname, item.firstname, item.middlename, 
        item.suffix, item.birthday, item.gender, item.address, 
        item.contact, item.email, item.batchid, item.idnumber, item.ulino
      ]);

if (excelRows.length > 0) {
    // 1. Define the full range of the data (Headers + Rows)
    const fullRange = sheet.getRange("A1").getResizedRange(excelRows.length, headers[0].length - 1);
    fullRange.values = [headers[0], ...excelRows]; // Write headers and rows in one go for speed

    // 2. CONVERT TO TABLE
    const enrollmentTable = sheet.tables.add(fullRange, true);
    enrollmentTable.name = "Enrollment";
   
        sheet.getUsedRange().getEntireColumn().format.autofitColumns();

          
        
        status.innerText = `✅ Successfully loaded ${excelRows.length} records into '${sheetName}'.`;
      } else {
        status.innerText = "⚠️ No records found.";
      }

      await context.sync();
    });
  } catch (error) {
    status.innerText = "❌ Error: " + error.message;
    console.error(error);
  }
}

/**
 * @param {string} fileUrl - URL of the .xlsx file
 * @param {string} sheetNamesCommaSeparated - e.g., "Attendance, Gradesheet"
 * @param {number} mode - 1 to Replace existing, 0 to Ignore
 */
async function downloadTemplate(fileUrl, sheetNamesCommaSeparated, mode) {
    const statusElement = document.getElementById("status-message");
    
    // 1. Convert string "Sheet1, Sheet2" into array ["Sheet1", "Sheet2"]
    const targetSheets = sheetNamesCommaSeparated.split(',').map(s => s.trim());

    try {
        statusElement.innerText = "Checking workbook...";
        
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            // 2. Logic for multiple sheets
            for (const name of targetSheets) {
                const existingSheet = sheets.items.find(s => s.name === name);
                if (existingSheet) {
                    if (mode === 1) {
                        existingSheet.delete(); 
                    } else {
                        // If mode 0 and even ONE sheet exists, we stop to prevent duplicates
                        statusElement.innerText = `Sheet ${name} already exists. Skipping.`;
                        return; 
                    }
                }
            }

            // 3. Download the file
            statusElement.innerText = "Downloading templates...";
            const response = await fetch(fileUrl);
            if (!response.ok) throw new Error("Network response was not ok");
            const buffer = await response.arrayBuffer();
            const base64 = arrayBufferToBase64(buffer);

            // 4. Insert the array of sheets
            statusElement.innerText = "Inserting sheets...";
            context.workbook.insertWorksheetsFromBase64(base64, {
                sheetNamesToInsert: targetSheets, // Passes the whole array
                positionType: Excel.WorksheetPositionType.after,
                relativeTo: context.workbook.worksheets.getActiveWorksheet()
            });

            await context.sync();
            statusElement.innerText = "All sheets synchronized!";
            statusElement.style.color = "green";
        });

    } catch (error) {
        console.error(error);
        statusElement.innerText = "Error: " + error.message;
        statusElement.style.color = "red";
    }
}

/**
 * Helper function to convert ArrayBuffer to Base64
 */
function arrayBufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
}