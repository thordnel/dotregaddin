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
    await sleep(100);
    setProgress(97);
    await downloadTemplate(templateUrl, "Advisory,InstructorSchedule", 1);
    await sleep(100);
    setProgress(98);
    status.innerText = "Rebuilding formulas. Please wait...";

    const userid = await getSettingValue(3);
    await createSchedule(userid);

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
                sheet.visibility = Excel.SheetVisibility.veryHidden;
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
    const instructorId = await getSettingValue(3); 
    
    if (!instructorId) {
        console.error("No Instructor ID found. Please log in again.");
        return;
    }
    await syncTableFromApi(
        '/fl_get_schedule',
        'scheduletab',
        'ScheduleTab',
        { instructorid: instructorId }
    );
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

            context.workbook.insertWorksheetsFromBase64(base64, {
                sheetNamesToInsert: targetSheets,
                positionType: Excel.WorksheetPositionType.end
            });


            await context.sync();

            // 3. Loop through the target sheet names and access them from the workbook
            targetSheets.forEach((sheetName) => {
                const sheet = context.workbook.worksheets.getItem(sheetName);
                
                if (sheetName === "InstructorSchedule") {
                    sheet.customProperties.add("sheetType", "schedule_record");
                } else if (sheetName === "Advisory") {
                    sheet.customProperties.add("sheetType", "advisory_record");
                }
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
async function injectSheetFormulas(context, sheet, baseName, batchId) {
    //return;
    let startRow = 21;
    let studentRowsCount = 30; 
    let formulaPayload, uniqueName, cell, subjectIdRange, currentRow, subjectId, allNames;
    switch (baseName) {        
        case "Attendance":
            sheet.getRange("B6").values = [[batchId]]; 
            sheet.getRange("F10").values = [[`=XLOOKUP(1,(TranscriptTab[BatchID]=B6)*(TranscriptTab[instructorid]=Settings!B3),TranscriptTab[subjectno])`]]; 
            sheet.getRange("A5").formulas = [[`=XLOOKUP(B6, BatchlistTab[batchid], BatchlistTab[batchname])`]];
            sheet.getRange("A15").formulas = [[`=FILTER(HSTACK(TranscriptTab[TraineesID], LEFT(TranscriptTab[gender], 1),TranscriptTab[lastname] & ", " & TranscriptTab[firstname] & IF(TranscriptTab[suffix]<>"", ", " & TranscriptTab[suffix], "") & IF(OR(TRIM(TranscriptTab[middlename])<>"", TranscriptTab[middlename]<>"."), " " & LEFT(TRIM(TranscriptTab[middlename]), 1) & ".", "")), TranscriptTab[BatchID]=B6, "")`]];
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
            sheet.getRange("B21").formulas = [[`=FILTER(HSTACK(TranscriptTab[lastname] & ", " & TranscriptTab[firstname] & IF(TranscriptTab[suffix]<>"", ", " & TranscriptTab[suffix], "") & IF(TRIM(TranscriptTab[middlename])<>"", IF(TranscriptTab[middlename]<>".", " " & LEFT(TRIM(TranscriptTab[middlename]), 1) & ".", ""), "")), TranscriptTab[BatchID]=K15)`]];
            sheet.getRange("M21").formulas = [[`=FILTER(TranscriptTab[recordid], TranscriptTab[batchid]=K15)`]];
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
            sheet.getRange("B21").formulas = [[`=FILTER(HSTACK(LEFT(TranscriptTab[gender], 1), TranscriptTab[TraineesID],TranscriptTab[lastname] & ", " & TranscriptTab[firstname] & IF(TranscriptTab[suffix]<>"", ", " & TranscriptTab[suffix], "") & IF(TRIM(TranscriptTab[middlename])<>"", IF(TranscriptTab[middlename]<>".", " " & LEFT(TRIM(TranscriptTab[middlename]), 1) & ".", ""), "")), TranscriptTab[BatchID]=I12)`]]; 
            sheet.getRange("BR21").formulas = [[`=FILTER(TranscriptTab[recordid], TranscriptTab[batchid]=I12)`]]; 
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
            sheet.getRange("B21").formulas = [[`=FILTER(HSTACK(LEFT(TranscriptTab[gender], 1), TranscriptTab[TraineesID],TranscriptTab[lastname] & ", " & TranscriptTab[firstname] & IF(TranscriptTab[suffix]<>"", ", " & TranscriptTab[suffix], "") & IF(TRIM(TranscriptTab[middlename])<>"", IF(TranscriptTab[middlename]<>".", " " & LEFT(TRIM(TranscriptTab[middlename]), 1) & ".", ""), "")), TranscriptTab[BatchID]=I12)`]];
            sheet.getRange("BR21").formulas = [[`=FILTER(TranscriptTab[recordid], TranscriptTab[batchid]=I12)`]];
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
                `=FILTER(TranscriptTab[lastname] & ", " & TranscriptTab[firstname] & IF(TranscriptTab[suffix]<>"", ", " & TranscriptTab[suffix], "") & IF(TRIM(TranscriptTab[middlename])<>"", IF(TranscriptTab[middlename]<>".", " " & LEFT(TRIM(TranscriptTab[middlename]), 1) & ".", ""), ""), (TranscriptTab[BatchID]=F8)*(TranscriptTab[gender]="Male"), "")`
            ]];
            sheet.getRange("E16").formulas = [[
                `=FILTER(TranscriptTab[lastname] & ", " & TranscriptTab[firstname] & IF(TranscriptTab[suffix]<>"", ", " & TranscriptTab[suffix], "") & IF(TRIM(TranscriptTab[middlename])<>"", IF(TranscriptTab[middlename]<>".", " " & LEFT(TRIM(TranscriptTab[middlename]), 1) & ".", ""), ""), (TranscriptTab[BatchID]=F8)*(TranscriptTab[gender]="Female"), "")`
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
        
        case "InstructorSchedule":
            sheet.getRange("A10").values = [[batchId]]; 
            sheet.getRange("B6").formulas = [[`=XLOOKUP(A10, batchlisttab[batchid], batchlisttab[year])`]];
            sheet.getRange("B7").formulas = [[`=XLOOKUP(A10, batchlisttab[batchid], batchlisttab[period])`]];
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


async function setSchedule(context, sheet, cellsRange, subjectcode, shortname, room, subjecttitle, backcolor) {
    // Safety Check: Ensure cellsRange is a valid array with exactly 4 addresses
    if (!cellsRange || !Array.isArray(cellsRange) || cellsRange.length !== 4) {
        console.warn(`[setSchedule] Invalid range for ${subjectcode}. Expected 4 cells, got:`, cellsRange);
        return;
    }

    try {
        // 1. Set values individually
        const cell0 = sheet.getRange(cellsRange[0]);
        cell0.values = [[(subjectcode || "").toUpperCase()]];
        cell0.format.font.set({ name: "Calibri", size: 18, bold: true });
        
        const cell1 = sheet.getRange(cellsRange[1]);
        const truncatedTitle = (subjecttitle || "").length > 25 
            ? (subjecttitle.slice(0, 25) + "…") 
            : (subjecttitle || "");

        cell1.values = [[truncatedTitle.toUpperCase()]];
        cell1.format.font.set({ name: "Calibri", size: 15, bold: false });

        const cell2 = sheet.getRange(cellsRange[2]);
        const cleanName = (shortname || "").replace("Batch ", "").trim();
        cell2.values = [[cleanName]];
        cell2.format.font.set({ name: "Calibri", size: 16, bold: true });

        const cell3 = sheet.getRange(cellsRange[3]);
        cell3.values = [[room || ""]];
        cell3.format.font.set({ name: "Calibri", size: 15, italic: false });

        // 2. Apply formatting to the block
        const fullRange = sheet.getRange(`${cellsRange[0]}:${cellsRange[3]}`);
        
        if (backcolor) {
            let color = backcolor.toString();
            if (!color.startsWith("#")) color = "#" + color;
            if (color.length === 7) fullRange.format.fill.color = color;
        }

        fullRange.format.horizontalAlignment = "Center";
        fullRange.format.verticalAlignment = "Center";

    } catch (error) {
        console.error("Error in setSchedule:", error);
    }
}

/**
 * Handles 30-minute slots (Expects 2 cells)
 */
async function setScheduleSplit(context, sheet, cellsRange, subjectcode, shortname, room, subjecttitle, backcolor) {
    // Safety Check: Ensure cellsRange is a valid array with exactly 2 addresses
    if (!cellsRange || !Array.isArray(cellsRange) || cellsRange.length !== 2) {
        console.warn(`[setScheduleSplit] Invalid range for ${subjectcode}. Expected 2 cells, got:`, cellsRange);
        return;
    }

    try {
        // For split slots, we often combine info due to limited space
        const cell0 = sheet.getRange(cellsRange[0]);
        cell0.values = [[(subjectcode || "").toUpperCase()]];
        cell0.format.font.set({ name: "Calibri", size: 18, bold: true });
        
        const cell1 = sheet.getRange(cellsRange[1]);
        // Combine Room and Batch/Shortname to fit in the second cell
        const cleanName = (shortname || "").replace("Batch ", "").trim();
        cell1.values = [[`${cleanName} | ${room || ""}`]];
        cell1.format.font.set({ name: "Calibri", size: 15, bold: false });

        const fullRange = sheet.getRange(`${cellsRange[0]}:${cellsRange[1]}`);
        
        if (backcolor) {
            let color = backcolor.toString();
            if (!color.startsWith("#")) color = "#" + color;
            if (color.length === 7) fullRange.format.fill.color = color;
        }

        fullRange.format.horizontalAlignment = "Center";
        fullRange.format.verticalAlignment = "Center";

    } catch (error) {
        console.error("Error in setScheduleSplit:", error);
    }
}

/**
 * Updated calling function
 */
async function createSchedule(instructorid) {
    const rawAddress = await getSettingValue(2);
    const token = await getSettingValue(6);
    const status = document.getElementById("status-message");

    if (!token) {
        showAuthOverlay(async () => await createSchedule(instructorid));
        return;
    }

    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name, items/customProperties");
            await context.sync();

            let targetSheet = null;
            for (let sheet of worksheets.items) {
                const typeProp = sheet.customProperties.getItemOrNullObject("sheetType");
                typeProp.load("value");
                await context.sync();

                if (!typeProp.isNullObject && typeProp.value === "schedule_record") {
                    targetSheet = sheet;
                    targetSheet.activate();
                    await context.sync();
                    break;
                }
            }

            if (!targetSheet) {

                console.error("Schedule sheet not found.");
                return;
            }
    status.innerText = "Retrieving latest schedule...";
    status.style.color = "#0078d4";
                        // --- INITIALIZE SHEET (CLEANING) ---
            // Define the morning and afternoon ranges
            const morningRange = targetSheet.getRange("B11:G30");
            const afternoonRange = targetSheet.getRange("B32:G51");

            // Clear values and formatting
            morningRange.clear(Excel.ClearApplyTo.contents);
            morningRange.format.fill.color = "#FFFFFF"; // Set to White

            afternoonRange.clear(Excel.ClearApplyTo.contents);
            afternoonRange.format.fill.color = "#FFFFFF"; // Set to White

            // Sync after clearing to ensure UI updates before drawing new data
            await context.sync();
            //console.log(instructorid)
            const response = await fetch(`https://${rawAddress}/schedule/fetch?instructorid=${encodeURIComponent(instructorid)}`, {
                method: 'GET',
                headers: { 'Authorization': `Bearer ${token}` }
            });

            if (response.status === 401 || response.status === 403) {
                showAuthOverlay(async () => await createSchedule(instructorid));
                return;
            }

            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

            status.innerText = "Importing latest schedule data...";
            await refreshScheduleData();
            const scheddata = await response.json();
            let loadhr = 0;

            for (let row of scheddata) {
                let {
                    schedulecode, 
                    shortname, 
                    room, 
                    splitstate, 
                    backcolor,
                    subjectcode,
                    subjecttitle
                } = row;

                if (splitstate === 1) {
                    // Use schedule_map_split (2 cells)
                    const cellsRange = typeof schedule_map_split !== 'undefined' ? schedule_map_split[schedulecode] : null;
                    if (cellsRange) {
                        await setScheduleSplit(context, targetSheet, cellsRange, subjectcode, shortname, room, subjecttitle, backcolor);
                        loadhr += 0.5;
                    } else {
                        console.warn(`No split mapping found for: ${schedulecode}`);
                    }
                } else {
                    // Use schedule_map (4 cells)
                    const cellsRange = typeof schedule_map !== 'undefined' ? schedule_map[schedulecode] : null;
                    if (cellsRange) {
                        await setSchedule(context, targetSheet, cellsRange, subjectcode, shortname, room, subjecttitle, backcolor);
                        loadhr += 1;
                    } else {
                        console.warn(`No standard mapping found for: ${schedulecode}`);
                    }
                }
            }
            const range = targetSheet.getRange("A9");
            range.select();

            await context.sync();
            
            //console.log(`Schedule successfully populated. Total load: ${loadhr} hrs.`);
            status.innerText = "Instructor's schedule updated.";
            status.style.color = "green";
            setTimeout(() => {
               
                status.innerText = "Ready";
                status.style.color = "#605e5c"; // Neutral gray
            }, 5000);
        });
        
    } catch (error) {
        console.error("Error creating schedule:", error);
    }
}

schedule_map = {
    "MON0": ["B11", "B12", "B13", "B14"],
    "MON1": ["B15", "B16", "B17", "B18"],
    "MON2": ["B19", "B20", "B21", "B22"],
    "MON3": ["B23", "B24", "B25", "B26"],
    "MON4": ["B27", "B28", "B29", "B30"],
    "MON5": ["B32", "B33", "B34", "B35"],
    "MON6": ["B36", "B37", "B38", "B39"],
    "MON7": ["B40", "B41", "B42", "B43"],
    "MON8": ["B44", "B45", "B46", "B47"],
    "MON9": ["B48", "B49", "B50", "B51"],
    "MON0": ["B11", "B12", "B13", "B14"],
    "MON1": ["B15", "B16", "B17", "B18"],
    "MON2": ["B19", "B20", "B21", "B22"],
    "MON3": ["B23", "B24", "B25", "B26"],
    "MON4": ["B27", "B28", "B29", "B30"],
    "MON5": ["B32", "B33", "B34", "B35"],
    "MON6": ["B36", "B37", "B38", "B39"],
    "MON7": ["B40", "B41", "B42", "B43"],
    "MON8": ["B44", "B45", "B46", "B47"],
    "MON9": ["B48", "B49", "B50", "B51"],
    "TUE0": ["C11", "C12", "C13", "C14"],
    "TUE1": ["C15", "C16", "C17", "C18"],
    "TUE2": ["C19", "C20", "C21", "C22"],
    "TUE3": ["C23", "C24", "C25", "C26"],
    "TUE4": ["C27", "C28", "C29", "C30"],
    "TUE5": ["C32", "C33", "C34", "C35"],
    "TUE6": ["C36", "C37", "C38", "C39"],
    "TUE7": ["C40", "C41", "C42", "C43"],
    "TUE8": ["C44", "C45", "C46", "C47"],
    "TUE9": ["C48", "C49", "C50", "C51"],
    "TUE0": ["C11", "C12", "C13", "C14"],
    "TUE1": ["C15", "C16", "C17", "C18"],
    "TUE2": ["C19", "C20", "C21", "C22"],
    "TUE3": ["C23", "C24", "C25", "C26"],
    "TUE4": ["C27", "C28", "C29", "C30"],
    "TUE5": ["C32", "C33", "C34", "C35"],
    "TUE6": ["C36", "C37", "C38", "C39"],
    "TUE7": ["C40", "C41", "C42", "C43"],
    "TUE8": ["C44", "C45", "C46", "C47"],
    "TUE9": ["C48", "C49", "C50", "C51"],
    "WED0": ["D11", "D12", "D13", "D14"],
    "WED1": ["D15", "D16", "D17", "D18"],
    "WED2": ["D19", "D20", "D21", "D22"],
    "WED3": ["D23", "D24", "D25", "D26"],
    "WED4": ["D27", "D28", "D29", "D30"],
    "WED5": ["D32", "D33", "D34", "D35"],
    "WED6": ["D36", "D37", "D38", "D39"],
    "WED7": ["D40", "D41", "D42", "D43"],
    "WED8": ["D44", "D45", "D46", "D47"],
    "WED9": ["D48", "D49", "D50", "D51"],
    "WED0": ["D11", "D12", "D13", "D14"],
    "WED1": ["D15", "D16", "D17", "D18"],
    "WED2": ["D19", "D20", "D21", "D22"],
    "WED3": ["D23", "D24", "D25", "D26"],
    "WED4": ["D27", "D28", "D29", "D30"],
    "WED5": ["D32", "D33", "D34", "D35"],
    "WED6": ["D36", "D37", "D38", "D39"],
    "WED7": ["D40", "D41", "D42", "D43"],
    "WED8": ["D44", "D45", "D46", "D47"],
    "WED9": ["D48", "D49", "D50", "D51"],
    "THU0": ["E11", "E12", "E13", "E14"],
    "THU1": ["E15", "E16", "E17", "E18"],
    "THU2": ["E19", "E20", "E21", "E22"],
    "THU3": ["E23", "E24", "E25", "E26"],
    "THU4": ["E27", "E28", "E29", "E30"],
    "THU5": ["E32", "E33", "E34", "E35"],
    "THU6": ["E36", "E37", "E38", "E39"],
    "THU7": ["E40", "E41", "E42", "E43"],
    "THU8": ["E44", "E45", "E46", "E47"],
    "THU9": ["E48", "E49", "E50", "E51"],
    "THU0": ["E11", "E12", "E13", "E14"],
    "THU1": ["E15", "E16", "E17", "E18"],
    "THU2": ["E19", "E20", "E21", "E22"],
    "THU3": ["E23", "E24", "E25", "E26"],
    "THU4": ["E27", "E28", "E29", "E30"],
    "THU5": ["E32", "E33", "E34", "E35"],
    "THU6": ["E36", "E37", "E38", "E39"],
    "THU7": ["E40", "E41", "E42", "E43"],
    "THU8": ["E44", "E45", "E46", "E47"],
    "THU9": ["E48", "E49", "E50", "E51"],
	"FRI0": ["F11", "F12", "F13", "F14"],
    "FRI1": ["F15", "F16", "F17", "F18"],
    "FRI2": ["F19", "F20", "F21", "F22"],
    "FRI3": ["F23", "F24", "F25", "F26"],
    "FRI4": ["F27", "F28", "F29", "F30"],
    "FRI5": ["F32", "F33", "F34", "F35"],
    "FRI6": ["F36", "F37", "F38", "F39"],
    "FRI7": ["F40", "F41", "F42", "F43"],
    "FRI8": ["F44", "F45", "F46", "F47"],
    "FRI9": ["F48", "F49", "F50", "F51"],
    "FRI0": ["F11", "F12", "F13", "F14"],
    "FRI1": ["F15", "F16", "F17", "F18"],
    "FRI2": ["F19", "F20", "F21", "F22"],
    "FRI3": ["F23", "F24", "F25", "F26"],
    "FRI4": ["F27", "F28", "F29", "F30"],
    "FRI5": ["F32", "F33", "F34", "F35"],
    "FRI6": ["F36", "F37", "F38", "F39"],
    "FRI7": ["F40", "F41", "F42", "F43"],
    "FRI8": ["F44", "F45", "F46", "F47"],
    "FRI9": ["F48", "F49", "F50", "F51"],
    "SAT0": ["G11", "G12", "G13", "G14"],
    "SAT1": ["G15", "G16", "G17", "G18"],
    "SAT2": ["G19", "G20", "G21", "G22"],
    "SAT3": ["G23", "G24", "G25", "G26"],
    "SAT4": ["G27", "G28", "G29", "G30"],
    "SAT5": ["G32", "G33", "G34", "G35"],
    "SAT6": ["G36", "G37", "G38", "G39"],
    "SAT7": ["G40", "G41", "G42", "G43"],
    "SAT8": ["G44", "G45", "G46", "G47"],
    "SAT9": ["G48", "G49", "G50", "G51"],
    "SAT0": ["G11", "G12", "G13", "G14"],
    "SAT1": ["G15", "G16", "G17", "G18"],
    "SAT2": ["G19", "G20", "G21", "G22"],
    "SAT3": ["G23", "G24", "G25", "G26"],
    "SAT4": ["G27", "G28", "G29", "G30"],
    "SAT5": ["G32", "G33", "G34", "G35"],
    "SAT6": ["G36", "G37", "G38", "G39"],
    "SAT7": ["G40", "G41", "G42", "G43"],
    "SAT8": ["G44", "G45", "G46", "G47"],
    "SAT9": ["G48", "G49", "G50", "G51"],    
}

schedule_map_split = {
    "MON0": ["B11", "B12"],
    "MON0B": ["B13", "B14"],
    "MON1": ["B15", "B16"],
    "MON1B": ["B17", "B18"],    
    "MON2": ["B19", "B20"],
    "MON2B": ["B21", "B22"],
    "MON3": ["B23", "B24"],
    "MON3B": ["B25", "B26"],    
    "MON4": ["B27", "B28"],
    "MON4B": ["B29", "B30"],    
    "MON5": ["B32", "B33"],
    "MON5B": ["B34", "B35"],
    "MON6": ["B36", "B37"],
    "MON6B": ["B38", "B39"],
    "MON7": ["B40", "B41"],
    "MON7B": ["B42", "B43"],  
    "MON8": ["B44", "B45"],
    "MON8B": ["B46", "B47"],
    "MON9": ["B48", "B49"],
    "MON9B": ["B50", "B51"],
    "TUE0": ["C11", "C12"],
    "TUE0B": ["C13", "C14"],
    "TUE1": ["C15", "C16"],
    "TUE1B": ["C17", "C18"],    
    "TUE2": ["C19", "C20"],
    "TUE2B": ["C21", "C22"],
    "TUE3": ["C23", "C24"],
    "TUE3B": ["C25", "C26"],    
    "TUE4": ["C27", "C28"],
    "TUE4B": ["C29", "C30"],    
    "TUE5": ["C32", "C33"],
    "TUE5B": ["C34", "C35"],
    "TUE6": ["C36", "C37"],
    "TUE6B": ["C38", "C39"],
    "TUE7": ["C40", "C41"],
    "TUE7B": ["C42", "C43"],  
    "TUE8": ["C44", "C45"],
    "TUE8B": ["C46", "C47"],
    "TUE9": ["C48", "C49"],
    "TUE9B": ["C50", "C51"],  
    "WED0": ["D11", "D12"],
    "WED0B": ["D13", "D14"],
    "WED1": ["D15", "D16"],
    "WED1B": ["D17", "D18"],    
    "WED2": ["D19", "D20"],
    "WED2B": ["D21", "D22"],
    "WED3": ["D23", "D24"],
    "WED3B": ["D25", "D26"],    
    "WED4": ["D27", "D28"],
    "WED4B": ["D29", "D30"],    
    "WED5": ["D32", "D33"],
    "WED5B": ["D34", "D35"],
    "WED6": ["D36", "D37"],
    "WED6B": ["D38", "D39"],
    "WED7": ["D40", "D41"],
    "WED7B": ["D42", "D43"],  
    "WED8": ["D44", "D45"],
    "WED8B": ["D46", "D47"],
    "WED9": ["D48", "D49"],
    "WED9B": ["D50", "D51"], 
    "THU0": ["E11", "E12"],
    "THU0B": ["E13", "E14"],
    "THU1": ["E15", "E16"],
    "THU1B": ["E17", "E18"],    
    "THU2": ["E19", "E20"],
    "THU2B": ["E21", "E22"],
    "THU3": ["E23", "E24"],
    "THU3B": ["E25", "E26"],    
    "THU4": ["E27", "E28"],
    "THU4B": ["E29", "E30"],    
    "THU5": ["E32", "E33"],
    "THU5B": ["E34", "E35"],
    "THU6": ["E36", "E37"],
    "THU6B": ["E38", "E39"],
    "THU7": ["E40", "E41"],
    "THU7B": ["E42", "E43"],  
    "THU8": ["E44", "E45"],
    "THU8B": ["E46", "E47"],
    "THU9": ["E48", "E49"],
    "THU9B": ["E50", "E51"],  
    "FRI0": ["F11", "F12"],
    "FRI0B": ["F13", "F14"],
    "FRI1": ["F15", "F16"],
    "FRI1B": ["F17", "F18"],    
    "FRI2": ["F19", "F20"],
    "FRI2B": ["F21", "F22"],
    "FRI3": ["F23", "F24"],
    "FRI3B": ["F25", "F26"],    
    "FRI4": ["F27", "F28"],
    "FRI4B": ["F29", "F30"],    
    "FRI5": ["F32", "F33"],
    "FRI5B": ["F34", "F35"],
    "FRI6": ["F36", "F37"],
    "FRI6B": ["F38", "F39"],
    "FRI7": ["F40", "F41"],
    "FRI7B": ["F42", "F43"],  
    "FRI8": ["F44", "F45"],
    "FRI8B": ["F46", "F47"],
    "FRI9": ["F48", "F49"],
    "FRI9B": ["F50", "F51"],  
    "SAT0": ["G11", "G12"],
    "SAT0B": ["G13", "G14"],
    "SAT1": ["G15", "G16"],
    "SAT1B": ["G17", "G18"],    
    "SAT2": ["G19", "G20"],
    "SAT2B": ["G21", "G22"],
    "SAT3": ["G23", "G24"],
    "SAT3B": ["G25", "G26"],    
    "SAT4": ["G27", "G28"],
    "SAT4B": ["G29", "G30"],    
    "SAT5": ["G32", "G33"],
    "SAT5B": ["G34", "G35"],
    "SAT6": ["G36", "G37"],
    "SAT6B": ["G38", "G39"],
    "SAT7": ["G40", "G41"],
    "SAT7B": ["G42", "G43"],  
    "SAT8": ["G44", "G45"],
    "SAT8B": ["G46", "G47"],
    "SAT9": ["G48", "G49"],
    "SAT9B": ["G50", "G51"],  
}
