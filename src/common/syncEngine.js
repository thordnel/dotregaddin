// src/common/syncEngine.js

async function performFullSync(setProgress, status, baseUrl) {
    await Excel.run(async (context) => {
        // 1. SET CALCULATION TO MANUAL
        context.workbook.application.calculationMode = Excel.CalculationMode.manual;
        
        // 2. SUSPEND SCREEN UPDATING
        context.workbook.application.suspendScreenUpdatingUntilNextSync();
        
        await context.sync();
    });
    setProgress(2);
    status.innerText = "Cleaning existing data...";
    await cleanupApiData();
    setProgress(5);
    status.innerText = "Syncing Class Record...";
    setProgress(15);
    status.innerText = "Importing batches data...";
    await refreshBatchlistData();
    setProgress(20);
    status.innerText = "Importing instructors data...";
    await refreshInstructorData();
    setProgress(27);
    status.innerText = "Importing enrollment data...";
    await refreshEnrollmentData();
    setProgress(32);
    status.innerText = "Importing transcript data...";
    await refreshTranscriptData();
    setProgress(35);
    status.innerText = "Importing attendance data...";
    await refreshAttendanceData();
    setProgress(40);
    status.innerText = "Importing schedule data...";
    await refreshScheduleData();
    setProgress(47);
    status.innerText = "Importing class standing data...";
    await refreshClassStandingData();
    setProgress(47);
    status.innerText = "Importing transmutation data...";
    await refreshTransmutationData();
    setProgress(51);
    status.innerText = "Importing room data...";
    setProgress(57);
    await refreshRoomsData();
    status.innerText = "Importing subject data...";
    await refreshSubjectData();

    // 2. Template Downloads
    status.innerText = "Downloading templates...";
    const templateUrl = `${baseUrl}/download/CLSRCDTemplate.xlsx`;
    const myBatches = await getAssignedBatchIds();
    const sheetsToCopy = "TraineeList,FinalTerm,Midterm,Gradesheet,Attendance";

    await downloadCRperBatch(templateUrl, sheetsToCopy, myBatches);
    setProgress(97);
    await downloadTemplate(templateUrl, "Advisory,InstructorSchedule,Base60", 1);
    setProgress(98);
    status.innerText = "Reconstructing formulas...";
    await reapplyAllFormulas();
    status.innerText = "Recalculating workbook...";
    
    await Excel.run(async (context) => {
        // 1. Restore Calculation to Automatic
        context.workbook.application.calculationMode = Excel.CalculationMode.automatic;
        
        // 2. Force a full recalculation to fix any remaining #REF! or #BUSY! displays
        context.workbook.application.calculate(Excel.CalculationType.full);
        
        // 3. (Optional) Re-enable screen updating if you manually turned it off
        // context.workbook.application.screenUpdating = true; 

        await context.sync();
    });
    setProgress(100);
    status.innerText = "Sync Complete!";
}

async function cleanupApiData() {
    await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        const tables = context.workbook.tables;
        sheets.load("items/name");
        tables.load("items/name");
        await context.sync();

        // 1. Delete Sheets ending in 'tab'
        const sheetsToDelete = sheets.items.filter(sheet => 
            sheet.name.toLowerCase().endsWith("tab")
        );
        sheetsToDelete.forEach(sheet => sheet.delete());

        // 2. Safety: Delete Tables specifically if the sheet survived
        const tablesToDelete = tables.items.filter(table => 
            table.name.toLowerCase().endsWith("tab")
        );
        tablesToDelete.forEach(table => table.delete());

        await context.sync();
    }).catch(err => console.log("Cleanup handled."));
}

async function getAssignedBatchIds() {
    
    const instructorId = localStorage.getItem("user_id");
    
    if (!instructorId) {
        console.error("No instructor ID found in session.");
        return [];
    }

    return await Excel.run(async (context) => {
        // 1. Get references to both tables
        const transcriptTable = context.workbook.tables.getItem("transcripttab");
        const batchListTable = context.workbook.tables.getItem("batchlisttab");
        
        const transBody = transcriptTable.getDataBodyRange();
        const transHeader = transcriptTable.getHeaderRowRange();
        const batchBody = batchListTable.getDataBodyRange();
        const batchHeader = batchListTable.getHeaderRowRange();
        
        transBody.load("values");
        transHeader.load("values");
        batchBody.load("values");
        batchHeader.load("values");
        
        await context.sync();

        // 2. Map Transcript Headers
        const tHeaders = transHeader.values[0].map(h => String(h).toLowerCase().trim());
        const idxTransInst = tHeaders.indexOf("instructorid");
        const idxTransBatch = tHeaders.indexOf("batchid");

        // 3. Map Batch List Headers
        const bHeaders = batchHeader.values[0].map(h => String(h).toLowerCase().trim());
        const idxBatchId = bHeaders.indexOf("batchid");
        const idxBatchName = bHeaders.indexOf("batchname");

        if (idxTransInst === -1 || idxTransBatch === -1 || idxBatchId === -1 || idxBatchName === -1) {
            console.error("Column mapping failed. Check header names in both tables.");
            return [];
        }

        // 4. Get unique Batch IDs for this instructor from Transcript Table
        const targetInstructorId = String(instructorId).trim();
        const uniqueIds = new Set();

        transBody.values.forEach(row => {
            const instId = String(row[idxTransInst]).trim();
            const bId = row[idxTransBatch] ? String(row[idxTransBatch]).trim() : "";
            
            if (instId === targetInstructorId && bId !== "" && bId !== "undefined") {
                uniqueIds.add(bId);
            }
        });

        // 5. Cross-reference with Batch List Table to get Names
        const assignedBatches = [];
        
        // Convert Set to Array to loop through our matches
        const uniqueIdArray = Array.from(uniqueIds);

        uniqueIdArray.forEach(id => {
            // Find the row in batchnametab that matches this ID
            const matchingRow = batchBody.values.find(row => String(row[idxBatchId]).trim() === id);
            
            assignedBatches.push({
                id: id,
                name: matchingRow ? String(matchingRow[idxBatchName]).trim() : `Batch ${id}` // Fallback if name not found
            });
        });

        // 6. Sort alphabetically by name
        return assignedBatches.sort((a, b) => a.name.localeCompare(b.name));
    });
}

function setProgress(percent) {
    const container = document.getElementById("progress-container");
    const bar = document.getElementById("progress-bar-width");
    const label = document.getElementById("progress-percentage");
    
    container.style.display = "block";
    bar.style.width = percent + "%";
    label.innerText = percent + "%";
}
/**
 * Modified Generic Sync Function
 * @param {string} endpoint - The API route
 * @param {string} sheetName - Target worksheet name
 * @param {string} tableName - Target Excel table name
 * @param {Object} params - Optional query parameters (e.g., { InstructorId: 123 })
 */
async function syncTableFromApi(endpoint, sheetName, tableName, params = {}) {
    
    const rawAddress = localStorage.getItem("registrar_url");
    const token = localStorage.getItem("access_token");
    
    // Construct URL with parameters
    const url = new URL(`https://${rawAddress}${endpoint}`);
    Object.keys(params).forEach(key => url.searchParams.append(key, params[key]));

    try {
        const response = await fetch(url, {
            method: 'GET',
            headers: { "Authorization": `Bearer ${token}` }
        });
        
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const data = await response.json();
        
        // If data is empty, we just clear the sheet and stop
        if (!data || data.length === 0) {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
                await context.sync();
                if (!sheet.isNullObject) sheet.getUsedRange().clear();
            });
            return;
        }

        await Excel.run(async (context) => {
            
            //context.workbook.application.suspendScreenUpdatingUntilNextSync()
            const sheets = context.workbook.worksheets;
            let sheet = sheets.getItemOrNullObject(sheetName);
            await context.sync();

            if (sheet.isNullObject) {
                sheet = sheets.add(sheetName);
                sheet.visibility = Excel.SheetVisibility.hidden;
            }

            // Cleanup existing table
            let oldTable = context.workbook.tables.getItemOrNullObject(tableName);
            await context.sync();
            if (!oldTable.isNullObject) {
                oldTable.delete();
            }

            sheet.getUsedRange().clear();

            // Dynamic Header/Row Logic
            const headers = Object.keys(data[0]);
            const excelRows = data.map(item => headers.map(key => item[key] ?? ""));

            const tableData = [headers, ...excelRows];
            const targetRange = sheet.getRange("A1").getResizedRange(
                tableData.length - 1, 
                headers.length - 1
            );
            targetRange.values = tableData;
            
            const newTable = sheet.tables.add(targetRange, true);
            newTable.name = tableName;
            
            //sheet.getUsedRange().getEntireColumn().format.autofitColumns();
            await context.sync();
        });
    } catch (error) {
        throw new Error(`${tableName} Sync Error: ${error.message}`);
    }
}
// Sync Enrollment
async function refreshEnrollmentData() {
    await syncTableFromApi('/fl_get_enrollment', 'enrollmenttab', 'EnrollmentTab');
}

// Sync Batchlist
async function refreshBatchlistData() {
    await syncTableFromApi('/fl_get_batchlist', 'batchlisttab', 'BatchlistTab');
}

// Sync Instructors
async function refreshInstructorData() {
    await syncTableFromApi('/fl_get_instructors', 'instructorstab', 'InstructorsTab');
}

async function refreshScheduleData() {
    await syncTableFromApi('/fl_get_schedule', 'scheduletab', 'ScheduleTab');
}

async function refreshAttendanceData() {
    // Retrieve the ID saved during login
    const instructorId = localStorage.getItem("user_id"); 
    
    if (!instructorId) {
        console.error("No Instructor ID found. Please log in again.");
        return;
    }

    // Pass the parameter as an object
    await syncTableFromApi(
        '/fl_get_attendance', 
        'attendancetab', 
        'AttendanceTab', 
        { InstructorId: instructorId }
    );
}

async function refreshClassStandingData() {
    // 1. Get the instructor/user ID saved during the login process
    const instructorId = localStorage.getItem("user_id");

    // 2. Safety check: If not logged in, stop and alert the user
    if (!instructorId) {
        console.error("Sync Failed: No Instructor ID found. Please log in first.");
        // Optional: show a message in your UI status element
        document.getElementById("status-message").innerText = "Error: Please log in again.";
        return;
    }

    try {
        //console.log(`Syncing Class Standing for Instructor ID: ${instructorId}...`);

        // 3. Use the helper function
        // Arguments: (API Endpoint, Sheet Name, Table Name, Query Parameters)
        await syncTableFromApi(
            '/fl_get_classstanding', 
            'classstandingtab', 
            'ClassstandingTab', 
            { InstructorId: instructorId }
        );

        //console.log("Class Standing sync complete!");
        
    } catch (error) {
        console.error("Class Standing Sync Error:", error.message);
        throw error; // Re-throw so your UI can catch it and display the error
    }
}

//async function refreshTranscriptData() {
//    await syncTableFromApi('/fl_get_transcript', 'transcripttab', 'TranscriptTab');
//}

async function refreshTranscriptData() {
    // 1. Get the instructor/user ID saved during the login process
    const instructorId = localStorage.getItem("user_id");

    // 2. Safety check: If not logged in, stop and alert the user
    if (!instructorId) {
        console.error("Sync Failed: No Instructor ID found. Please log in first.");
        // Optional: show a message in your UI status element
        document.getElementById("status-message").innerText = "Error: Please log in again.";
        return;
    }

    try {
        //console.log(`Syncing Class Standing for Instructor ID: ${instructorId}...`);

        // 3. Use the helper function
        // Arguments: (API Endpoint, Sheet Name, Table Name, Query Parameters)
        await syncTableFromApi(
            '/fl_get_transcript', 
            'transcripttab', 
            'TranscriptTab', 
            { InstructorId: instructorId }
        );

        //console.log("Class Standing sync complete!");
        
    } catch (error) {
        console.error("Transcript table Sync Error:", error.message);
        throw error; // Re-throw so your UI can catch it and display the error
    }
}


async function refreshRoomsData() {
    await syncTableFromApi('/fl_get_rooms', 'roomstab', 'RoomsTab');
}

async function refreshTransmutationData() {
    await syncTableFromApi('/fl_get_transmutation', 'transmutationtab', 'TransmutationTab');
}

async function refreshSubjectData() {
    await syncTableFromApi('/fl_get_subject', 'subjecttab', 'SubjectTab');
}

async function downloadTemplate(fileUrl, sheetNamesCommaSeparated, mode) {
    const statusElement = document.getElementById("status-message");
    const targetSheets = sheetNamesCommaSeparated.split(',').map(s => s.trim());
    
    // Get the token from local storage
    const token = localStorage.getItem("access_token");

    try {
        await Excel.run(async (context) => {
            context.workbook.application.suspendScreenUpdatingUntilNextSync()
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();

            // 1. Handle existing sheets based on mode
            for (const name of targetSheets) {
                const existingSheet = sheets.items.find(s => s.name === name);
                if (existingSheet) {
                    if (mode === 1) {
                        existingSheet.delete(); 
                    } else {
                        return; // Skip if mode 0 and sheet exists
                    }
                }
            }

            // 2. Fetch the template file using the Authorization header
            const response = await fetch(fileUrl, {
                method: 'GET',
                headers: {
                    "Authorization": `Bearer ${token}` // Mandatory for tokenRequired route
                }
            });

            if (!response.ok) {
                if (response.status === 403 || response.status === 401) {
                    throw new Error("Unauthorized: Please log in again.");
                }
                throw new Error(`Template download failed with status: ${response.status}`);
            }

            const buffer = await response.arrayBuffer();
            const base64 = arrayBufferToBase64(buffer);

            // 3. Insert the templates into the workbook
            
            context.workbook.insertWorksheetsFromBase64(base64, {
                sheetNamesToInsert: targetSheets,
                positionType: Excel.WorksheetPositionType.after,
                relativeTo: context.workbook.worksheets.getActiveWorksheet()
            });

            await context.sync();
        });
    } catch (error) {
        // Log details for debugging in the console
        console.error("Template Error:", error);
        throw new Error("Template Setup: " + error.message);
    }
}

async function downloadCRperBatch(fileUrl, sheetNamesCommaSeparated, batches) {
    const baseSheetNames = sheetNamesCommaSeparated.split(',').map(s => s.trim());
    const token = localStorage.getItem("access_token");
    const status = document.getElementById("status-message");
    const startProgress = 63;
    const endProgress = 97;
    const progressRange = endProgress - startProgress;
    const totalTasks = batches.length * baseSheetNames.length;
    let completedTasks = 0;

    try {
        const response = await fetch(fileUrl, {
            method: 'GET',
            headers: { "Authorization": `Bearer ${token}` }
        });

        if (!response.ok) throw new Error(`Download failed: ${response.status}`);
        
        const buffer = await response.arrayBuffer();
        const base64 = arrayBufferToBase64(buffer);

        await Excel.run(async (context) => {
            context.workbook.application.suspendScreenUpdatingUntilNextSync();
            const sheets = context.workbook.worksheets;

            for (const batch of batches) {
                const batchId = batch.id;
                const batchName = batch.name;

                for (const baseName of baseSheetNames) {
                    completedTasks++;
                    const currentProgress = startProgress + Math.floor((completedTasks / totalTasks) * progressRange);
                    if (typeof setProgress === "function") {
                        setProgress(currentProgress);
                    }

                    const newName = `${baseName}_${batchId}`;
                    let existingSheet = sheets.getItemOrNullObject(newName);
                    await context.sync();

                    if (!existingSheet.isNullObject) {
                        continue; 
                    }

                    // Update status with the Batch Name!
                    status.innerText = `Preparing ${baseName} table for ${batchName}. Please wait...`;
                    
                    context.workbook.insertWorksheetsFromBase64(base64, {
                        sheetNamesToInsert: [baseName], 
                        positionType: Excel.WorksheetPositionType.end
                    });

                    await context.sync();

                    const newlyAddedSheet = sheets.getItem(baseName);
                    newlyAddedSheet.name = newName;
                    newlyAddedSheet.visibility = Excel.SheetVisibility.hidden;
                    newlyAddedSheet.protection.unprotect("tesda");
                
                    newlyAddedSheet.customProperties.add("batchid", String(batchId));
                    newlyAddedSheet.customProperties.add("sheetType", `${baseName.toLowerCase()}_record`);
                    
                    await injectSheetFormulas(newlyAddedSheet, baseName, batchId)
                    await context.sync();
                }
            }
            setProgress(97);
            status.innerText = "All batch tables are ready!";
        });

    } catch (error) {
        console.error("Batch Template Error:", error);
        status.innerText = "Error during setup.";
        throw new Error("Batch Setup Failed: " + error.message);
    }
}

function arrayBufferToBase64(buffer) {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    for (let i = 0; i < bytes.byteLength; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
}


/**
 * Centralized place for all workbook formulas.
 * Prevents #REF! errors by re-applying logic after tables are rebuilt.
 */
async function injectSheetFormulas(sheet, baseName, batchId) {
    switch (baseName) {
                     
                    case "Attendance":
                        sheet.getRange("B6").values = [[batchId]]; 
                        sheet.getRange("A5").formulas = [[`=XLOOKUP(B6, 'BatchlistTab'!A:A, 'BatchlistTab'!C:C)`]];
                        // Fixed: Ensure the IF logic for Middle Initial handles empty strings correctly
                        sheet.getRange("A15").formulas = [[`=FILTER(HSTACK('EnrollmentTab'!L:L, LEFT('EnrollmentTab'!G:G, 1), 'EnrollmentTab'!B:B & ", " & 'EnrollmentTab'!C:C & " " & IF(AND('EnrollmentTab'!D:D<>".", 'EnrollmentTab'!D:D<>""), LEFT('EnrollmentTab'!D:D, 1) & ". ", "")), 'EnrollmentTab'!K:K=B6)`]];
                        sheet.getRange("B7").formulas = [[`=XLOOKUP('Settings'!B3, 'InstructorsTab'!A:A, 'InstructorsTab'!D:D & " " & LEFT('InstructorsTab'!E:E, 1) & ". " & 'InstructorsTab'!C:C & IF('InstructorsTab'!F:F<>"", ", " & 'InstructorsTab'!F:F, ""))`]];
                        sheet.getRange("B8").formulas = [[`=XLOOKUP(XLOOKUP(B6, 'BatchlistTab'!A:A, 'BatchlistTab'!J:J), 'InstructorsTab'!A:A, 'InstructorsTab'!D:D & " " & LEFT('InstructorsTab'!E:E, 1) & ". " & 'InstructorsTab'!C:C & IF('InstructorsTab'!F:F<>"", ", " & 'InstructorsTab'!F:F, ""))`]];
                        sheet.getRange("B9").formulas = [[`=XLOOKUP(XLOOKUP(1, ('TranscriptTab'!B:B=B6) * ('TranscriptTab'!D:D='Settings'!B3), 'TranscriptTab'!E:E), 'SubjectTab'!A:A, 'SubjectTab'!D:D)`]];
                        
                        // FIXED: Removed the extra '=' before XLOOKUP
                        sheet.getRange("B10").formulas = [[`=UPPER(XLOOKUP(XLOOKUP(1, ('TranscriptTab'!B:B=B6) * ('TranscriptTab'!D:D='Settings'!B3), 'TranscriptTab'!E:E), 'SubjectTab'!A:A, 'SubjectTab'!F:F))`]];
                        
                        sheet.getRange("E8").formulas = [[`=XLOOKUP(B6, 'batchlisttab'!A:A, 'batchlisttab'!H:H)`]];
                        sheet.getRange("E12").formulas = [[`=XLOOKUP(B6, 'batchlisttab'!A:A, 'batchlisttab'!K:K)`]];
                        sheet.getRange("E9").formulas = [[`=XLOOKUP(B6, 'batchlisttab'!A:A, 'batchlisttab'!I:I)`]];
                        sheet.getRange("F12").formulas = [[`=XLOOKUP(B6, 'batchlisttab'!A:A, 'batchlisttab'!L:L)`]];
                        break;

                    case "Gradesheet":
                        sheet.getRange("K15").values = [[batchId]]; 
                        sheet.getRange("B20").formulas = [[`=FILTER(HSTACK('EnrollmentTab'!B:B & ", " & 'EnrollmentTab'!C:C & " " & IF(AND('EnrollmentTab'!D:D<>".", 'EnrollmentTab'!D:D<>""), LEFT('EnrollmentTab'!D:D, 1) & ". ", "")), 'EnrollmentTab'!K:K=K15)`]];
                        sheet.getRange("A8").formulas = [[`=XLOOKUP(XLOOKUP(1, ('TranscriptTab'!B:B=K15) * ('TranscriptTab'!D:D='Settings'!B3), 'TranscriptTab'!E:E), 'SubjectTab'!A:A, 'SubjectTab'!D:D)`]];
                        sheet.getRange("A11").formulas = [[`=UPPER(XLOOKUP(XLOOKUP(1, ('TranscriptTab'!B:B=K15) * ('TranscriptTab'!D:D='Settings'!B3), 'TranscriptTab'!E:E), 'SubjectTab'!A:A, 'SubjectTab'!F:F))`]];
                        sheet.getRange("C8").formulas = [[`=XLOOKUP(K15, 'batchlisttab'!A:A, 'batchlisttab'!F:F)`]];
                        sheet.getRange("C11").formulas = [[`=XLOOKUP(K15, 'batchlisttab'!A:A, 'batchlisttab'!E:E)`]];
                        sheet.getRange("C14").formulas = [[`=XLOOKUP(XLOOKUP(K15, 'BatchlistTab'!A:A, 'BatchlistTab'!J:J), 'InstructorsTab'!A:A, 'InstructorsTab'!D:D & " " & LEFT('InstructorsTab'!E:E, 1) & ". " & 'InstructorsTab'!C:C & IF('InstructorsTab'!F:F<>"", ", " & 'InstructorsTab'!F:F, ""))`]];
                        sheet.getRange("I8").formulas = [[`=XLOOKUP(K15, 'BatchlistTab'!A:A, 'BatchlistTab'!C:C)`]];
                        sheet.getRange("I11").formulas = [[`=XLOOKUP('Settings'!B3, 'InstructorsTab'!A:A, 'InstructorsTab'!D:D & " " & LEFT('InstructorsTab'!E:E, 1) & ". " & 'InstructorsTab'!C:C & IF('InstructorsTab'!F:F<>"", ", " & 'InstructorsTab'!F:F, ""))`]];
                        break;

                    case "Midterm":
                        //sheet.getRange("B6").values = [[batchId]];
                    case "FinalTerm":
                        //sheet.getRange("B6").values = [[batchId]];
                        sheet.getRange("B21").formulas = [[`=FILTER(HSTACK(LEFT('EnrollmentTab'!G:G, 1), 'EnrollmentTab'!L:L, 'EnrollmentTab'!B:B & ", " & 'EnrollmentTab'!C:C), 'EnrollmentTab'!K:K="${batchId}")`]];
                        break;

                    case "TraineeList":
                        //sheet.getRange("B6").values = [[batchId]];
                        sheet.getRange("B16").formulas = [[`=FILTER('EnrollmentTab'!B:B & ", " & 'EnrollmentTab'!C:C, ('EnrollmentTab'!K:K="${batchId}")*('EnrollmentTab'!G:G="Male"), "")`]];
                        sheet.getRange("E16").formulas = [[`=FILTER('EnrollmentTab'!B:B & ", " & 'EnrollmentTab'!C:C, ('EnrollmentTab'!K:K="${batchId}")*('EnrollmentTab'!G:G="Female"), "")`]];
                        break;
                }
          
                
}
    
async function reapplyAllFormulas() {
    await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name, items/customProperties");
        await context.sync();

        for (const sheet of sheets.items) {
            // Find batch sheets (e.g., Attendance_211)
            if (sheet.name.includes("_")) {
                const batchProp = sheet.customProperties.getItemOrNullObject("batchid");
                const typeProp = sheet.customProperties.getItemOrNullObject("sheetType");
                batchProp.load("value");
                typeProp.load("value");
                await context.sync();

                if (!batchProp.isNullObject && !typeProp.isNullObject) {
                    const batchId = batchProp.value;
                    // Extract baseName (Attendance, Gradesheet, etc.)
                    const baseName = sheet.name.split("_")[0]; 
                    await injectSheetFormulas(sheet, baseName, batchId);
                }
            }
        }
        await context.sync();
    });
}