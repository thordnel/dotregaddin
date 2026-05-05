// src/common/syncEngine.js

async function performFullSync(setProgress, status, baseUrl) {
    status.style.color = "#0078d4";
    await ensureServerAwake(status, baseUrl);
    await Excel.run(async (context) => {
        // 1. SET CALCULATION TO MANUAL
        context.workbook.application.calculationMode = Excel.CalculationMode.manual;
        
        // 2. SUSPEND SCREEN UPDATING
        context.workbook.application.suspendScreenUpdatingUntilNextSync();
        
        await context.sync();
    });
    setProgress(2);
    status.innerText = "Cleaning existing data...";
    status.style.color = "#0078d4";

    setProgress(5);
    status.innerText = "Importing batches data...";
    await refreshBatchlistData();
    await sleep(300);
    setProgress(20);
    status.innerText = "Importing instructors data...";
    await refreshInstructorData();
    await sleep(300);
    setProgress(27);
    status.innerText = "Importing enrollment data...";
    await refreshEnrollmentData();
    await sleep(300);
    setProgress(32);
    status.innerText = "Importing transcript data...";
    await refreshTranscriptData();
    await sleep(300);
    setProgress(35);
    status.innerText = "Importing attendance data...";
    await refreshAttendanceData()
    await sleep(300);
    setProgress(40);
    status.innerText = "Importing schedule data...";
    await refreshScheduleData();
    await sleep(300);
    setProgress(47);
    status.innerText = "Importing class standing data...";
    await refreshClassStandingData();
    await sleep(300);
    setProgress(47);
    status.innerText = "Importing transmutation data...";
    await refreshTransmutationData();
    await sleep(300);
    setProgress(51);
    status.innerText = "Importing room data...";
    setProgress(57);
    await refreshRoomsData();
    await sleep(300);
    status.innerText = "Importing subject data...";
    await refreshSubjectData();
    await sleep(300);

    // 2. Template Downloads
    status.innerText = "Downloading templates...";
    const templateUrl = `${baseUrl}/download/CLSRCDTemplate.xlsx`;
    const myBatches = await getAssignedBatchIds();
    const sheetsToCopy = "TraineeList,FinalTerm,Midterm,Gradesheet,Attendance";

    await downloadCRperBatch(templateUrl, sheetsToCopy, myBatches);
    await sleep(300);
    setProgress(97);
    await downloadTemplate(templateUrl, "Advisory,InstructorSchedule", 1);
    await sleep(300);
    setProgress(98);
    status.innerText = "Rebuilding formulas...";
    await reapplyAllFormulas();
    status.innerText = "Recalculating workbook...";
        
    await Excel.run(async (context) => {
        // 1. Enable Automatic calculation
        context.workbook.application.calculationMode = Excel.CalculationMode.automatic;
        
        // 2. Use regular calculation instead of 'Full' to save time
        context.workbook.application.calculate(Excel.CalculationType.recalculate);
        
        await context.sync();
    });
    setProgress(100);
    status.innerText = "Sync Complete!";
    status.style.color = "green";
    }

/**
 * Simple delay helper
 * @param {number} ms - Milliseconds to wait
 */
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

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

/**
 * Generic helper to pull a setting from the "Settings" sheet.
 * Row 2 = Server, 3 = UserID, 4 = Username, 5 = InstName, 6 = Token
 */
async function getSettingValue(rowNumber) {
    return await Excel.run(async (context) => {
        const settingsSheet = context.workbook.worksheets.getItem("Settings");
        settingsSheet.calculate();
        // We target column B (index 2) and the specific row
        const range = settingsSheet.getRange(`B${rowNumber}`);
        range.load("values");
        await context.sync();
        
        return range.values[0][0]; 
    });
}

async function getWorkbookSession() {
    return await Excel.run(async (context) => {
        const settingsSheet = context.workbook.worksheets.getItem("Settings");
        settingsSheet.calculate();
        // Load the entire B2:B6 range (Server, UserID, User, Inst, Token)
        const range = settingsSheet.getRange("B2:B6");
        range.load("values");
        await context.sync();

        return {
            url: range.values[0][0],      // B2
            userId: range.values[1][0],   // B3
            username: range.values[2][0], // B4
            instName: range.values[3][0], // B5
            token: range.values[4][0]      // B6
        };
    });
}

/**
 * Saves a setting to the Excel "Settings" sheet.
 * Row 2=URL, 3=UserID, 4=Username, 5=InstName, 6=Token
 */
async function setWorkbookSetting(rowNumber, value) {
    await Excel.run(async (context) => {
        const settingsSheet = context.workbook.worksheets.getItem("Settings");
        settingsSheet.calculate();
        const range = settingsSheet.getRange(`B${rowNumber}`);
        range.values = [[String(value)]];
        context.workbook.application.calculate(Excel.CalculationType.recalculate);
        await context.sync();
    });
}

/**
 * Saves the entire session to the workbook in one trip.
 */
async function setWorkbookSession(url, userId, username, instName, token) {
    await Excel.run(async (context) => {
        const settingsSheet = context.workbook.worksheets.getItem("Settings");
        
        // Update B2 through B6
        settingsSheet.getRange("B2:B6").values = [
            [url],      // B2
            [userId],   // B3
            [username], // B4
            [instName], // B5
            [token]     // B6
        ];
        context.workbook.application.calculate(Excel.CalculationType.recalculate);
        await context.sync();
    });
}

/**
 * Clears a specific setting from the Excel "Settings" sheet.
 * Row 3 = UserID, 4 = Username, 5 = InstName, 6 = Token
 */
async function clearWorkbookSetting(rowNumber) {
    await Excel.run(async (context) => {
        const settingsSheet = context.workbook.worksheets.getItem("Settings");
        //settingsSheet.calculate();
        const range = settingsSheet.getRange(`B${rowNumber}`);
        range.clear(); // This removes the value and formatting
        context.workbook.application.calculate(Excel.CalculationType.recalculate);
        await context.sync();
    });
}

async function getAssignedBatchIds() {
    
    const instructorId = await getSettingValue(3);
    
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
// 1. Get the session for THIS specific workbook
    const session = await getWorkbookSession();
    
    if (!session.token || !session.url) {
        throw new Error("Missing auth token. Please log in.");
    }

    const url = new URL(`https://${session.url}${endpoint}`);
    //console.log(url);
    // 2. Inject the ID from the sheet into the API parameters
    // This prevents Workbook B from using Workbook A's ID
    const instructorId = session.userId;
    if (params) {
        // Many of your routes use 'InstructorId' (capitalized)
        params.InstructorId = instructorId;
    }

    Object.keys(params).forEach(key => url.searchParams.append(key, params[key]));


    try {
        const response = await fetch(url, {
            method: 'GET',
            headers: { "Authorization": `Bearer ${session.token}` } // Use session.token
        });
        
        if (response.status === 401 || response.status === 403) {
            throw new Error("401"); // Trigger the re-auth overlay
        }
        if (response.status === 502 && retries > 0) {
            console.warn(`Re-importing ${tableName}...`);
            await new Promise(res => setTimeout(res, 2000)); // Wait 2s
            return makeRequest(retries - 1);
        }
        
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
            if (tableName === "batchlisttab") {
                await cleanupApiData();
                setProgress(5);                
                }

  
        const data = await response.json();
        
        await Excel.run(async (context) => {
            // 1. FREEZE UI IMMEDIATELY
            context.workbook.application.suspendScreenUpdatingUntilNextSync();

            const sheets = context.workbook.worksheets;
            let sheet = sheets.getItemOrNullObject(sheetName);
            let oldTable = context.workbook.tables.getItemOrNullObject(tableName);
            
            // 2. Load existence check in one go
            sheet.load("isNullObject");
            oldTable.load("isNullObject");
            
            await context.sync();

            // Handle Empty Data Case
            if (!data || data.length === 0) {
                if (!sheet.isNullObject) sheet.getUsedRange().clear();
                return;
            }

            // 3. Setup Sheet
            if (sheet.isNullObject) {
                sheet = sheets.add(sheetName);
                sheet.visibility = Excel.SheetVisibility.hidden;
            }

            // 4. Cleanup Table and Range
            if (!oldTable.isNullObject) {
                oldTable.delete();
            }
            sheet.getUsedRange().clear();

            // 5. Prepare Data
            const headers = Object.keys(data[0]);
            const excelRows = data.map(item => headers.map(key => item[key] ?? ""));
            const tableData = [headers, ...excelRows];

            const targetRange = sheet.getRange("A1").getResizedRange(
                tableData.length - 1, 
                headers.length - 1
            );
            
            // 6. Write Data and Create Table
            targetRange.values = tableData;
            const newTable = sheet.tables.add(targetRange, true);
            newTable.name = tableName;
            
            // Final Sync - this is the only time the UI will refresh
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
    const instructorId = await getSettingValue(3); 
    
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
    const instructorId = await getSettingValue(3);

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
    const instructorId = await getSettingValue(3);

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
    const token = await getSettingValue(6);

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
                positionType: Excel.WorksheetPositionType.end,
                //relativeTo: context.workbook.worksheets.getActiveWorksheet()
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
    const token = await getSettingValue(6);
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
                    status.innerText = `Preparing ${baseName} table for ${batchName}`;
                    status.style.color = "#0078d4";
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
                    
                    await injectSheetFormulas(context,newlyAddedSheet, baseName, batchId)
                    await context.sync();
                }
            }
            setProgress(97);
            status.innerText = "All registrar tables are updated.";
        });

    } catch (error) {
        console.error("Batch Template Error:", error);
        status.innerText = "Error during setup.";
        status.style.color = "red";
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
async function injectSheetFormulas(context,sheet, baseName, batchId) {
    let startRow = 21;
    let studentRowsCount = 30; 
    let formulaPayload, uniqueName, cell, subjectIdRange, currentRow, subjectId, allNames;
    switch (baseName) {        
        case "Attendance":
            sheet.getRange("B6").values = [[batchId]]; 
            sheet.getRange("F10").values = [[`=XLOOKUP(1,(TranscriptTab[BatchID]=B6)*(TranscriptTab[instructorid]=Settings!B3),TranscriptTab[subjectno])`]]; 
            sheet.getRange("A5").formulas = [[`=XLOOKUP(B6, BatchlistTab[batchid], BatchlistTab[batchname])`]];
            sheet.getRange("A15").formulas = [[`=FILTER(HSTACK(EnrollmentTab[idnumber], LEFT(EnrollmentTab[gender], 1), EnrollmentTab[lastname] & ", " & EnrollmentTab[firstname] & " " & IF(AND(EnrollmentTab[middlename]<>".", EnrollmentTab[middlename]<>""), LEFT(EnrollmentTab[middlename], 1) & ". ", "")), EnrollmentTab[batchid]=B6)`]];
            sheet.getRange("B7").formulas = [[`=XLOOKUP(Settings!B3, InstructorsTab[idnumber], InstructorsTab[Firstname] & " " & LEFT(InstructorsTab[Middlename], 1) & ". " & InstructorsTab[Lastname] & IF(InstructorsTab[Suffix]<>"", ", " & InstructorsTab[Suffix], ""))`]];
            sheet.getRange("B8").formulas = [[`=XLOOKUP(XLOOKUP(B6, BatchlistTab[batchid], BatchlistTab[adviser]), InstructorsTab[idnumber], InstructorsTab[Firstname] & " " & LEFT(InstructorsTab[Middlename], 1) & ". " & InstructorsTab[Lastname] & IF(InstructorsTab[Suffix]<>"", ", " & InstructorsTab[Suffix], ""))`]];
            sheet.getRange("B9").formulas = [[`=XLOOKUP(XLOOKUP(1, (TranscriptTab[BatchID]=B6) * (TranscriptTab[instructorid]=Settings!B3), TranscriptTab[subjectno]), SubjectTab[subjectno], SubjectTab[subjectcode])`]];
            sheet.getRange("B10").formulas = [[`=UPPER(XLOOKUP(XLOOKUP(1, (TranscriptTab[BatchID]=B6) * (TranscriptTab[instructorid]=Settings!B3), TranscriptTab[subjectno]), SubjectTab[subjectno], SubjectTab[subjecttitle]))`]];
            sheet.getRange("E8").formulas = [[`=XLOOKUP(B6, batchlisttab[batchid], batchlisttab[trainingstart])`]];
            sheet.getRange("H12").formulas = [[`=E8`]] 
            sheet.getRange("E12").formulas = [[`=XLOOKUP(B6, batchlisttab[batchid], batchlisttab[midtermexamdate])`]];
            sheet.getRange("E9").formulas = [[`=XLOOKUP(B6, batchlisttab[batchid], batchlisttab[trainingend])`]];
            sheet.getRange("F12").formulas = [[`=XLOOKUP(B6, batchlisttab[batchid], batchlisttab[finaltermexamdate])`]];

            //attendance calculator
            sheet.getRange("H1").formulas = [[`=IFERROR(LEFT(ADDRESS(1,MATCH(E8,H12:IW12,0)+COLUMN(H12)-1,4),LEN(ADDRESS(1,MATCH(E8,H12:IW12,0)+COLUMN(H12)-1,4))-1)&"13",H12)`]]
            sheet.getRange("I1").formulas = [[`=IFERROR(LEFT(ADDRESS(1,MATCH(E12,H12:IW12,0)+COLUMN(H12)-1,4),LEN(ADDRESS(1,MATCH(E12,H12:IW12,0)+COLUMN(H12)-1,4))-1)&"13",H12)`]]
            sheet.getRange("J1").formulas = [[`=SUBSTITUTE(ADDRESS(1,COLUMN(INDIRECT(H1)),4),1,"")`]]
            sheet.getRange("K1").formulas = [[`=SUBSTITUTE(ADDRESS(1,COLUMN(INDIRECT(I1)),4),1,"")`]]
            sheet.getRange("H2").formulas = [[`=IFERROR(LEFT(ADDRESS(1,MATCH(E12+1,H12:IW12,0)+COLUMN(H12)-1,4),LEN(ADDRESS(1,MATCH(E12+1,H12:IW12,0)+COLUMN(H12)-1,4))-1)&"13",H12)`]]
            sheet.getRange("I2").formulas = [[`=IFERROR(LEFT(ADDRESS(1,MATCH(E9,H12:IW12,0)+COLUMN(H12)-1,4),LEN(ADDRESS(1,MATCH(E9,H12:IW12,0)+COLUMN(H12)-1,4))-1)&"13",H12)`]]
            sheet.getRange("J2").formulas = [[`=SUBSTITUTE(ADDRESS(1,COLUMN(INDIRECT(H2)),4),"1","")`]]
            sheet.getRange("K2").formulas = [[`=SUBSTITUTE(ADDRESS(1,COLUMN(INDIRECT(I2)),4),1,"")`]]          
            sheet.getRange("E13").formulas = [[`=SUMPRODUCT((INDIRECT(H1):INDIRECT(I1)=1)*(INDIRECT(H1):INDIRECT(I1)<>""))`]]
            sheet.getRange("F13").formulas = [[`=SUMPRODUCT((INDIRECT(H2):INDIRECT(I2)=1)*(INDIRECT(H2):INDIRECT(I2)<>""""))`]]  
            sheet.getRange("H51").formulas = [[`=UNIQUE(LEFT(FILTER(ScheduleTab[schedulecode], (ScheduleTab[instructorid]=Settings!B3) * (ScheduleTab[subjectno]=F10)), 3))`]]
            //sheet.getRange("H51:H61").setFontColor("white");

            const startColIndex = 7; // Column H is index 7 (0-indexed)
            const endColIndex = 256; // Column IW is index 256
            const numColumns = endColIndex - startColIndex + 1;
            const targetRow = 8;

            // Helper to convert index to Column Letter (e.g., 7 -> "H")
            const getColumnLetter = (index) => {
                let letter = "";
                while (index >= 0) {
                    letter = String.fromCharCode((index % 26) + 65) + letter;
                    index = Math.floor(index / 26) - 1;
                }
                return letter;
            };

            // Generate horizontal payload: [[formula1, formula2, ...]]
            formulaPayload = [
                Array.from({ length: numColumns }, (_, i) => {
                    const colLetter = getColumnLetter(startColIndex + i);
                    return `=--ISNUMBER(MATCH(UPPER(TEXT(${colLetter}$12,"DDD")),$H$51:$H$59,0))`;
                })
            ];

            // Target Range: H12 to IW12
            const targetRange = sheet.getRangeByIndexes(targetRow - 1, startColIndex, 1, numColumns);
            targetRange.formulas = formulaPayload;

            await context.sync();

            startRow = 15;
            formulaPayload = Array.from({ length: studentRowsCount }, (_, i) => {
                currentRow = startRow + i;
                return [
                `=IF(A${currentRow}<>"",SUMPRODUCT((INDIRECT(H$1):INDIRECT(I$1)=1)*(INDIRECT(J$1&ROW()):INDIRECT(K$1&ROW())<>"")),0)`,
                `=IF(A${currentRow}<>"",SUMPRODUCT((INDIRECT(H$2):INDIRECT(I$2)=1)*(INDIRECT(J$2&ROW()):INDIRECT(K$2&ROW())<>"")),0)`
                ];
            });
            sheet.getRange(`E${startRow}:F${startRow + studentRowsCount - 1}`).formulas = formulaPayload;

            startRow = 15;
            subjectIdRange = sheet.getRange("F10");
            context.workbook.application.calculate("Full");
            subjectIdRange.load("values");
            allNames = context.workbook.names;
            allNames.load("items/name");
            await context.sync();
            subjectId = subjectIdRange.values[0][0];
            if (typeof subjectId === "string" && subjectId.startsWith("#")) {
                throw new Error("Subject ID not found yet. Please ensure TranscriptTab is populated.");
                }
            const AMpref = `AM_${batchId}_${subjectId}_`;
            const AFpref = `AF_${batchId}_${subjectId}_`;
            allNames.items
                .filter(n => n.name.startsWith(AMpref))
                .forEach(n => n.delete());
            allNames.items
                .filter(n => n.name.startsWith(AFpref))
                .forEach(n => n.delete());            
            for (let i = 0; i < studentRowsCount; i++) {
                currentRow = startRow  + i;
                uniqueNameAM = `${AMpref}${currentRow+6}`;
                uniqueNameAF = `${AFpref}${currentRow+6}`;
                cellAM = sheet.getRange(`E${currentRow}`);
                cellAF = sheet.getRange(`F${currentRow}`);
                allNames.add(uniqueNameAM, cellAM);
                allNames.add(uniqueNameAF, cellAF);
            }
            allNames.add(`AM_${batchId}_${subjectId}_18`, sheet.getRange("E13"));
            allNames.add(`AF_${batchId}_${subjectId}_18`, sheet.getRange("F13"));
            await context.sync();

            console.log("Restored formulas in attendance tab");
            break;
        case "Gradesheet":
            sheet.getRange("K15").values = [[batchId]]; 
            sheet.getRange("M12").values = [[`=XLOOKUP(1,(TranscriptTab[BatchID]=K15)*(TranscriptTab[instructorid]=Settings!B3),TranscriptTab[subjectno])`]]; 
            sheet.getRange("K14").values = [[`=COUNTIF(B20:B51, "?*")`]]; 
            sheet.getRange("B21").formulas = [[`=FILTER(HSTACK(EnrollmentTab[lastname] & ", " & EnrollmentTab[firstname] & " " & IF(AND(EnrollmentTab[middlename]<>".", EnrollmentTab[middlename]<>""), LEFT(EnrollmentTab[middlename], 1) & ". ", "")), EnrollmentTab[batchid]=K15)`]];
            sheet.getRange("M21").formulas = [[`=FILTER(EnrollmentTab[recordid], EnrollmentTab[batchid]=K15)`]];
            sheet.getRange("A8").formulas = [[`=XLOOKUP(XLOOKUP(1, (TranscriptTab[BatchID]=K15) * (TranscriptTab[instructorid]=Settings!B3), TranscriptTab[subjectno]), SubjectTab[subjectno], SubjectTab[subjectcode])`]];
            sheet.getRange("A11").formulas = [[`=UPPER(XLOOKUP(XLOOKUP(1, (TranscriptTab[BatchID]=K15) * (TranscriptTab[instructorid]=Settings!B3), TranscriptTab[subjectno]), SubjectTab[subjectno], SubjectTab[subjecttitle]))`]];
            sheet.getRange("C8").formulas = [[`=XLOOKUP(K15, batchlisttab[batchid], batchlisttab[year])`]];
            sheet.getRange("C11").formulas = [[`=XLOOKUP(K15, batchlisttab[batchid], batchlisttab[period])`]];
            sheet.getRange("C14").formulas = [[`=XLOOKUP(XLOOKUP(K15, BatchlistTab[batchid], BatchlistTab[adviser]), InstructorsTab[idnumber], InstructorsTab[Firstname] & " " & LEFT(InstructorsTab[Middlename], 1) & ". " & InstructorsTab[Lastname] & IF(InstructorsTab[Suffix]<>"", ", " & InstructorsTab[Suffix], ""))`]];
            sheet.getRange("H8").formulas = [[`=UPPER(XLOOKUP(K15, BatchlistTab[batchid], BatchlistTab[course])) & " (" & MID(K13, SEARCH("Batch", K13), SEARCH(" (", K13) - SEARCH("Batch", K13)) & ")"`]];
            sheet.getRange("K13").formulas = [[`=XLOOKUP(K15, BatchlistTab[batchid], BatchlistTab[batchname])`]]
            sheet.getRange("H11").formulas = [[`=XLOOKUP(Settings!B3, InstructorsTab[idnumber], InstructorsTab[Firstname] & " " & LEFT(InstructorsTab[Middlename], 1) & ". " & InstructorsTab[Lastname] & IF(InstructorsTab[Suffix]<>"", ", " & InstructorsTab[Suffix], ""))`]];
            sheet.getRange("C20:H50").clear(Excel.ClearApplyTo.contents);
            sheet.getRange("A20:A50").clear(Excel.ClearApplyTo.contents);
            sheet.getRange("A20:S20").clear(Excel.ClearApplyTo.contents);
            await context.sync();

            //gradesheet GP formulas
            startRow = 21;
            formulaPayload = Array.from({ length: studentRowsCount }, (_, i) => {
                currentRow = startRow + i;
                return [
                `=IF(B${currentRow}<>"", SUMIFS(TransmutationTab[gradepoint], TransmutationTab[rawscore_min], "<="&N${currentRow}, TransmutationTab[rawscore_max], ">="&N${currentRow}), "")`,
                `=IF(B${currentRow}<>"", SUMIFS(TransmutationTab[gradepoint], TransmutationTab[rawscore_min], "<="&O${currentRow}, TransmutationTab[rawscore_max], ">="&O${currentRow}), "")`,
                `=IF(B${currentRow}<>"", SUMIFS(TransmutationTab[gradepoint], TransmutationTab[rawscore_min], "<="&P${currentRow}, TransmutationTab[rawscore_max], ">="&P${currentRow}), "")`,
                `=IF(B${currentRow}<>"", SUMIFS(TransmutationTab[gradepoint], TransmutationTab[rawscore_min], "<="&Q${currentRow}, TransmutationTab[rawscore_max], ">="&Q${currentRow}), "")`,
                `=IF(B${currentRow}<>"", SUMIFS(TransmutationTab[gradepoint], TransmutationTab[rawscore_min], "<="&R${currentRow}, TransmutationTab[rawscore_max], ">="&R${currentRow}), "")`,
                `=IF(B${currentRow}<>"", SUMIFS(TransmutationTab[gradepoint], TransmutationTab[rawscore_min], "<="&S${currentRow}, TransmutationTab[rawscore_max], ">="&S${currentRow}), "")`
                ];
            });
            sheet.getRange(`C${startRow}:H${startRow + studentRowsCount - 1}`).formulas = formulaPayload;

            //gradesheet raw grades formulas
            subjectIdRange = sheet.getRange("M12");
            context.workbook.application.calculate("Full");
            subjectIdRange.load("values");
            allNames = context.workbook.names;
            allNames.load("items/name");
            await context.sync();
            subjectId = subjectIdRange.values[0][0];
            if (typeof subjectId === "string" && subjectId.startsWith("#")) {
                throw new Error("Subject ID not found yet. Please ensure TranscriptTab is populated.");
                }
            startRow = 21;
            formulaPayload = Array.from({ length: studentRowsCount }, (_, i) => {
                currentRow = startRow + i;
                return [
                `=IFERROR(MT_${batchId}_${subjectId}_${currentRow},0)`,
                `=IFERROR(ML_${batchId}_${subjectId}_${currentRow},0)`,
                `=IFERROR(FT_${batchId}_${subjectId}_${currentRow},0)`,
                `=IFERROR(FL_${batchId}_${subjectId}_${currentRow},0)`
                ];
            });
            sheet.getRange(`N${startRow}:Q${startRow + studentRowsCount - 1}`).formulas = formulaPayload;            
            
            startRow = 21;
            formulaPayload = Array.from({ length: studentRowsCount }, (_, i) => {
                currentRow = startRow + i;
                return [
                `=IFERROR(MT_${batchId}_${subjectId}_${currentRow},0)`,
                `=IFERROR(ML_${batchId}_${subjectId}_${currentRow},0)`,
                `=IFERROR(FT_${batchId}_${subjectId}_${currentRow},0)`,
                `=IFERROR(FL_${batchId}_${subjectId}_${currentRow},0)`
                ];
            });
            sheet.getRange(`N${startRow}:Q${startRow + studentRowsCount - 1}`).formulas = formulaPayload; 

            //raw total (40+60%)
            startRow = 21;
            formulaPayload = Array.from({ length: studentRowsCount }, (_, i) => {
                currentRow = startRow + i;
                return [
                    `=iferror(if(B${currentRow}<>"",(N${currentRow}*N19)+(P${currentRow}*P19),""),0)`,
                    `=iferror(if(B${currentRow}<>"",(O${currentRow}*O19)+(Q${currentRow}*Q19),""),0)`
                ];
            });
            sheet.getRange(`R${startRow}:S${startRow + studentRowsCount - 1}`).formulas = formulaPayload;

            //numbering
            startRow = 21;
            formulaPayload = Array.from({ length: studentRowsCount }, (_, i) => {
                currentRow = startRow + i;
                return [`=IF(B${currentRow}<>"",ROW()-20 & ".","")`];
            });
            sheet.getRange(`A${startRow}:A${startRow + studentRowsCount - 1}`).formulas = formulaPayload;
            
            console.log("Restored formulas in gradesheet tab");
            break;
        
        case "Midterm":
            sheet.getRange("I12").values = [[batchId]]; 
            sheet.getRange("I13").values = [[`=XLOOKUP(1,(TranscriptTab[BatchID]=I12)*(TranscriptTab[instructorid]=Settings!B3),TranscriptTab[subjectno])`]]; 
            sheet.getRange("B21").formulas = [[`=FILTER(HSTACK(LEFT(EnrollmentTab[gender], 1), EnrollmentTab[idnumber], EnrollmentTab[lastname] & ", " & EnrollmentTab[firstname] & " " & IF((EnrollmentTab[middlename]<>".")*(EnrollmentTab[middlename]<>""), LEFT(EnrollmentTab[middlename], 1) & ". ", "")), EnrollmentTab[batchid]=I12)`]]; 
            sheet.getRange("BR21").formulas = [[`=FILTER(EnrollmentTab[recordid], EnrollmentTab[batchid]=I12)`]]; 
            sheet.getRange("C6").formulas = [[`=XLOOKUP(I12, batchlisttab[batchid], batchlisttab[year])`]];
            sheet.getRange("C7").formulas = [[`=XLOOKUP(I12, batchlisttab[batchid], batchlisttab[period])`]];    
            sheet.getRange("C8").values = [[`MIDTERM`]];    
            sheet.getRange("C9").formulas = [[`=UPPER(XLOOKUP(I12, batchlisttab[batchid], batchlisttab[course]))`]];    
            sheet.getRange("C11").formulas = [[`=XLOOKUP(XLOOKUP(I12, BatchlistTab[batchid], BatchlistTab[adviser]), InstructorsTab[idnumber], InstructorsTab[Firstname] & " " & LEFT(InstructorsTab[Middlename], 1) & ". " & InstructorsTab[Lastname] & IF(InstructorsTab[Suffix]<>"", ", " & InstructorsTab[Suffix], ""))`]];    
            sheet.getRange("C10").formulas = [[`=XLOOKUP(Settings!B3, InstructorsTab[idnumber], InstructorsTab[Firstname] & " " & LEFT(InstructorsTab[Middlename], 1) & ". " & InstructorsTab[Lastname] & IF(InstructorsTab[Suffix]<>"", ", " & InstructorsTab[Suffix], ""))`]];   
            sheet.getRange("C12").formulas = [[`=XLOOKUP(I12, BatchlistTab[batchid], BatchlistTab[batchname])`]];    
            sheet.getRange("C13").formulas = [[`=XLOOKUP(XLOOKUP(1, (TranscriptTab[BatchID]=I12) * (TranscriptTab[instructorid]=Settings!B3), TranscriptTab[subjectno]), SubjectTab[subjectno], SubjectTab[subjectcode]) & " - " & UPPER(XLOOKUP(XLOOKUP(1, (TranscriptTab[BatchID]=I12) * (TranscriptTab[instructorid]=Settings!B3), TranscriptTab[subjectno]), SubjectTab[subjectno], SubjectTab[subjecttitle]))`]];   
            sheet.getRange("C14").formulas = [[`=IF(XLOOKUP(XLOOKUP(Settings!B3& I12, ScheduleTab[instructorid] & ScheduleTab[batchid], ScheduleTab[subjectno]),SubjectTab[subjectno],SubjectTab[labhours])<1,"TOOL SUBJECT","CORE SUBJECT")`]]
            sheet.getRange("K18").values = 100;
            sheet.getRange("Z18").values = 100;
            sheet.getRange("AI18").values = 100;
            sheet.getRange("AL18").values = 100;
            sheet.getRange("AX18").values = 100;
            sheet.getRange("BH18").values = 100;
            sheet.getRange("BQ18").values = 100;
            

            
            startRow = 21;
            formulaPayload = Array.from({ length: studentRowsCount }, (_, i) => {
                currentRow = startRow + i;
                return [
                    `=IF(E${currentRow}="Failed",1,(K${currentRow}*$N$8+Z${currentRow}*$N$9+AI${currentRow}*$N$10+AL${currentRow}*$N$11))`,
                    `=SUMIFS(TransmutationTab[gradepoint], TransmutationTab[rawscore_min], "<="&F${currentRow}, TransmutationTab[rawscore_max], ">="&F${currentRow})`,
                    `=IF(E${currentRow}="Failed",1,AX${currentRow}*$AR$8+BH${currentRow}*$AR$9+BQ${currentRow}*$AR$10)`,
                    `=SUMIFS(TransmutationTab[gradepoint], TransmutationTab[rawscore_min], "<="&H${currentRow}, TransmutationTab[rawscore_max], ">="&H${currentRow})`
                ];
            });
            sheet.getRange(`F${startRow}:I${startRow + studentRowsCount - 1}`).formulas = formulaPayload;
          
            startRow = 21;
            formulaPayload = Array.from({ length: studentRowsCount }, (_, i) => {
                currentRow = startRow + i;
                return [`=IF(B${currentRow}<>"",ROW()-20 & ".","")`];
            });
            sheet.getRange(`A${startRow}:A${startRow + studentRowsCount - 1}`).formulas = formulaPayload;

            //setting named ranges for Midterm Lec and Lab
            startRow = 21;
            subjectIdRange = sheet.getRange("I13");
            context.workbook.application.calculate("Full");
            subjectIdRange.load("values");
            allNames = context.workbook.names;
            allNames.load("items/name");
            await context.sync();
            subjectId = subjectIdRange.values[0][0];
            if (typeof subjectId === "string" && subjectId.startsWith("#")) {
                throw new Error("Subject ID not found yet. Please ensure TranscriptTab is populated.");
                }
            const MTpref = `MT_${batchId}_${subjectId}_`;
            const MLpref = `ML_${batchId}_${subjectId}_`;
            allNames.items
                .filter(n => n.name.startsWith(MTpref))
                .forEach(n => n.delete());
            allNames.items
                .filter(n => n.name.startsWith(MLpref))
                .forEach(n => n.delete());            
            for (let i = 0; i < studentRowsCount; i++) {
                currentRow = startRow + i;
                uniqueNameMT = `${MTpref}${currentRow}`;
                uniqueNameML = `${MLpref}${currentRow}`;
                cellMT = sheet.getRange(`F${currentRow}`);
                cellML = sheet.getRange(`H${currentRow}`);
                allNames.add(uniqueNameMT, cellMT);
                allNames.add(uniqueNameML, cellML);
                }
            await context.sync();

            formulaPayload = Array.from({ length: studentRowsCount }, (_, i) => {
                currentRow = startRow + i;
                return [`=IFERROR(AM_${batchId}_${subjectId}_${currentRow},0)`];
            });
            sheet.getRange(`J${startRow}:J${startRow + studentRowsCount - 1}`).formulas = formulaPayload;
            sheet.getRange("J18").formulas = [[`=iferror(AM_${batchId}_${subjectId}_18,0)`]]


            console.log("Restored formulas in midterm tab");
            break;
                
        case "FinalTerm":
            sheet.getRange("I12").values = [[batchId]]; 
            sheet.getRange("I13").values = [[`=XLOOKUP(1,(TranscriptTab[BatchID]=I12)*(TranscriptTab[instructorid]=Settings!B3),TranscriptTab[subjectno])`]]; 
            sheet.getRange("B21").formulas = [[`=FILTER(HSTACK(LEFT(EnrollmentTab[gender], 1), EnrollmentTab[idnumber], EnrollmentTab[lastname] & ", " & EnrollmentTab[firstname] & " " & IF((EnrollmentTab[middlename]<>".")*(EnrollmentTab[middlename]<>""), LEFT(EnrollmentTab[middlename], 1) & ". ", "")), EnrollmentTab[batchid]=I12)`]];
            sheet.getRange("BR21").formulas = [[`=FILTER(EnrollmentTab[recordid], EnrollmentTab[batchid]=I12)`]];
            sheet.getRange("C6").formulas = [[`=XLOOKUP(I12, batchlisttab[batchid], batchlisttab[year])`]];
            sheet.getRange("C7").formulas = [[`=XLOOKUP(I12, batchlisttab[batchid], batchlisttab[period])`]];    
            sheet.getRange("C8").values = [[`FINAL TERM`]];    
            sheet.getRange("C9").formulas = [[`=UPPER(XLOOKUP(I12, batchlisttab[batchid], batchlisttab[course]))`]];    
            sheet.getRange("C11").formulas = [[`=XLOOKUP(XLOOKUP(I12, BatchlistTab[batchid], BatchlistTab[adviser]), InstructorsTab[idnumber], InstructorsTab[Firstname] & " " & LEFT(InstructorsTab[Middlename], 1) & ". " & InstructorsTab[Lastname] & IF(InstructorsTab[Suffix]<>"", ", " & InstructorsTab[Suffix], ""))`]];    
            sheet.getRange("C10").formulas = [[`=XLOOKUP(Settings!B3, InstructorsTab[idnumber], InstructorsTab[Firstname] & " " & LEFT(InstructorsTab[Middlename], 1) & ". " & InstructorsTab[Lastname] & IF(InstructorsTab[Suffix]<>"", ", " & InstructorsTab[Suffix], ""))`]];   
            sheet.getRange("C12").formulas = [[`=XLOOKUP(I12, BatchlistTab[batchid], BatchlistTab[batchname])`]];    
            sheet.getRange("C13").formulas = [[`=XLOOKUP(XLOOKUP(1, (TranscriptTab[BatchID]=I12) * (TranscriptTab[instructorid]=Settings!B3), TranscriptTab[subjectno]), SubjectTab[subjectno], SubjectTab[subjectcode]) & " - " & UPPER(XLOOKUP(XLOOKUP(1, (TranscriptTab[BatchID]=I12) * (TranscriptTab[instructorid]=Settings!B3), TranscriptTab[subjectno]), SubjectTab[subjectno], SubjectTab[subjecttitle]))`]];   
            sheet.getRange("C14").formulas = [[`=IF(XLOOKUP(XLOOKUP(Settings!B3& I12, ScheduleTab[instructorid] & ScheduleTab[batchid], ScheduleTab[subjectno]),SubjectTab[subjectno],SubjectTab[labhours])<1,"TOOL SUBJECT","CORE SUBJECT")`]];
            startRow = 21;
            
            formulaPayload = Array.from({ length: studentRowsCount }, (_, i) => {
                currentRow = startRow + i;
                return [
                    `=IF(E${currentRow}="Failed",1,(K${currentRow}*$N$8+Z${currentRow}*$N$9+AI${currentRow}*$N$10+AL${currentRow}*$N$11))`,
                    `=SUMIFS(TransmutationTab[gradepoint], TransmutationTab[rawscore_min], "<="&F${currentRow}, TransmutationTab[rawscore_max], ">="&F${currentRow})`,
                    `=IF(E${currentRow}="Failed",1,AX${currentRow}*$AR$8+BH${currentRow}*$AR$9+BQ${currentRow}*$AR$10)`,
                    `=SUMIFS(TransmutationTab[gradepoint], TransmutationTab[rawscore_min], "<="&H${currentRow}, TransmutationTab[rawscore_max], ">="&H${currentRow})`
                ];
            });
            sheet.getRange(`F${startRow}:I${startRow + studentRowsCount - 1}`).formulas = formulaPayload;



            formulaPayload = Array.from({ length: studentRowsCount }, (_, i) => {
                currentRow = startRow + i;
                return [
                    `=IF(B${currentRow}<>"",ROW()-20 & ".","")`
                ];
            });
            sheet.getRange(`A${startRow}:A${startRow + studentRowsCount - 1}`).formulas = formulaPayload;

            //setting named ranges for Midterm Lec and Lab
            startRow = 21;
            subjectIdRange = sheet.getRange("I13");
            context.workbook.application.calculate("Full");
            subjectIdRange.load("values");
            allNames = context.workbook.names;
            allNames.load("items/name");
            await context.sync();
            subjectId = subjectIdRange.values[0][0];
            if (!subjectId || typeof subjectId === "string" && (subjectId.startsWith("#") || subjectId === "")) {
                throw new Error(`Invalid Subject ID (${subjectId}) in FinalTerm. Sync failed.`);
            }
            const FTpref = `FT_${batchId}_${subjectId}_`;
            const FLpref = `FL_${batchId}_${subjectId}_`;
            allNames.items
                .filter(n => n.name.startsWith(FTpref))
                .forEach(n => n.delete());
            allNames.items
                .filter(n => n.name.startsWith(FLpref))
                .forEach(n => n.delete());            
            for (let i = 0; i < studentRowsCount; i++) {
                currentRow = startRow + i;
                uniqueNameFT = `${FTpref}${currentRow}`;
                uniqueNameFL = `${FLpref}${currentRow}`;
                cellFT = sheet.getRange(`F${currentRow}`);
                cellFL = sheet.getRange(`H${currentRow}`);
                allNames.add(uniqueNameFT, cellFT);
                allNames.add(uniqueNameFL, cellFL);
                //console.log(uniqueNameFT);
                }
            await context.sync();

            formulaPayload = Array.from({ length: studentRowsCount }, (_, i) => {
                currentRow = startRow + i;
                return [`=IFERROR(AF_${batchId}_${subjectId}_${currentRow},0)`];
            });
            sheet.getRange(`J${startRow}:J${startRow + studentRowsCount - 1}`).formulas = formulaPayload;
            sheet.getRange("J18").formulas = [[`=iferror(AF_${batchId}_${subjectId}_18,0)`]]

            console.log("Restored formulas in final term tab");
            break;
        case "TraineeList":
            sheet.getRange("F8").values = [[batchId]]; 
            sheet.getRange("B16").formulas = [[
                `=FILTER(EnrollmentTab[lastname] & ", " & EnrollmentTab[firstname] & " " & IF((EnrollmentTab[middlename]<>".")*(EnrollmentTab[middlename]<>""), LEFT(EnrollmentTab[middlename], 1) & ". ", ""), (EnrollmentTab[batchid]=F8)*(EnrollmentTab[gender]="Male"), "")`
            ]];
            sheet.getRange("E16").formulas = [[
                `=FILTER(EnrollmentTab[lastname] & ", " & EnrollmentTab[firstname] & " " & IF((EnrollmentTab[middlename]<>".")*(EnrollmentTab[middlename]<>""), LEFT(EnrollmentTab[middlename], 1) & ". ", ""), (EnrollmentTab[batchid]=F8)*(EnrollmentTab[gender]="Female"), "")`
            ]];
            sheet.getRange("C8").formulas = [[`=XLOOKUP(F8, BatchlistTab[batchid], BatchlistTab[batchname])`]];
            sheet.getRange("C9").formulas = [[`=UPPER(XLOOKUP(F8, batchlisttab[batchid], batchlisttab[course]))`]];  
            sheet.getRange("C10").formulas = [[`=TEXT(XLOOKUP(F8, BatchlistTab[batchid], BatchlistTab[trainingstart]),"dd MMM YYYY") & " - " & TEXT(XLOOKUP(F8, BatchlistTab[batchid], BatchlistTab[trainingend]),"dd MMM YYYY")`]];
            sheet.getRange("C11").formulas = [[`=XLOOKUP(F8, batchlisttab[batchid], batchlisttab[period]) & " - " & XLOOKUP(F8, batchlisttab[batchid], batchlisttab[year])`]];
            sheet.getRange("C12").formulas = [[`=XLOOKUP(XLOOKUP(1, (TranscriptTab[BatchID]=F8) * (TranscriptTab[instructorid]=Settings!B3), TranscriptTab[subjectno]), SubjectTab[subjectno], SubjectTab[subjectcode])`]];   
            sheet.getRange("F10").formulas = [[`=XLOOKUP(Settings!B3, InstructorsTab[idnumber], InstructorsTab[Firstname] & " " & LEFT(InstructorsTab[Middlename], 1) & ". " & InstructorsTab[Lastname] & IF(InstructorsTab[Suffix]<>"", ", " & InstructorsTab[Suffix], ""))`]];
            sheet.getRange("F11").formulas = [[`=XLOOKUP(XLOOKUP(F8, BatchlistTab[batchid], BatchlistTab[adviser]), InstructorsTab[idnumber], InstructorsTab[Firstname] & " " & LEFT(InstructorsTab[Middlename], 1) & ". " & InstructorsTab[Lastname] & IF(InstructorsTab[Suffix]<>"", ", " & InstructorsTab[Suffix], ""))`]];
            console.log("Restored formulas in traineeslist tab");
            
            break;
    }            
}
    
async function reapplyAllFormulas() {
    // Calling Excel.run without an external context parameter 
    // ensures 'context' is scoped ONLY to this task pane's parent workbook.
    await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        
        // 1. Load names and the properties collection for ALL sheets
        sheets.load("items/name, items/id, items/customProperties");
        await context.sync();

        let sheetsToProcess = [];

        // 2. Queue up the 'value' loads for properties
        for (let sheet of sheets.items) {
            // Safety: only target sheets following your naming pattern
            if (sheet.name.includes("_")) {
                const bProp = sheet.customProperties.getItemOrNullObject("batchid");
                const tProp = sheet.customProperties.getItemOrNullObject("sheetType");
                
                bProp.load("value");
                tProp.load("value");

                sheetsToProcess.push({
                    sheet: sheet,
                    batchProp: bProp,
                    typeProp: tProp,
                    name: sheet.name
                });
            }
        }

        // 3. Middle Sync: Get metadata for all relevant sheets at once
        await context.sync();

        // 4. Process the formulas in memory
        for (const item of sheetsToProcess) {
            // Only proceed if the sheet has the required metadata
            if (!item.batchProp.isNullObject && !item.typeProp.isNullObject) {
                const batchId = item.batchProp.value;
                const type = item.typeProp.value;
                const baseName = item.name.split("_")[0];
                
                // Ensure injectSheetFormulas uses the specific 'item.sheet' 
                // passed into it, which is tied to this specific 'context'.
                await injectSheetFormulas(context,item.sheet, baseName, batchId);
            }
        }

        // 5. Final Sync: Commit all formulas to the parent workbook
        await context.sync();
        console.log("Formulas reapplied to parent workbook successfully.");
    });
}

async function ensureServerAwake(status, baseUrl) {
    //console.log(baseUrl);
    status.innerText = "Connecting to server... this may take a while";
    status.style.color = "#ffa500"; // Orange for "Working on it"

    let isAwake = false;
    let attempts = 0;
    const maxAttempts = 12; // Wait up to 60 seconds (12 * 5s)

    while (!isAwake && attempts < maxAttempts) {
        try {
            const response = await fetch(`${baseUrl}/check`, { method: 'GET' });
            if (response.ok) {
                isAwake = true;
                return true;
            }
        } catch (e) {
            // Server is still sleeping, ignore the network error
        }
        
        attempts++;
        // Wait 5 seconds before trying again
        await new Promise(resolve => setTimeout(resolve, 5000));
    }

    if (!isAwake) {
        throw new Error("The server is taking too long to wake up. Please try again in a moment.");
    }
}