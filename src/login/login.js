// src/login/login.js
/* global document, Excel, Office */

Office.onReady(async (info) => {
    if (info.host === Office.HostType.Excel) {
        const connectBtn = document.getElementById("connect-button");
        const saveBtn = document.getElementById("save-settings-button");
        const loginBtn = document.getElementById("login-button"); // Select the button

        if (connectBtn) connectBtn.onclick = toggleConnection;
        if (saveBtn) saveBtn.onclick = saveServerToExcel;
        
        // Attach the handleLogin function to the click event
        if (loginBtn) loginBtn.onclick = handleLogin; 

        await loadSettingsFromExcel();
    }
});

let isConnected = false;


async function toggleConnection() {
    const urlInput = document.getElementById("server-url");
    const connectLabel = document.getElementById("connect-label");
    const saveBtn = document.getElementById("save-settings-button");
    const credentialsSection = document.getElementById("credentials-section");
    const status = document.getElementById("status-message");

    if (!connectLabel || !urlInput) return;

    if (!isConnected) {
        const rawValue = urlInput.value.trim();

        // 1. Strict Validation: Fail if ANY protocol is detected
        if (rawValue.includes("://")) {
            status.innerText = "❌ Error: Do not include 'https://' or 'http://'. Use address only.";
            status.style.color = "red";
            return;
        }

        if (!rawValue) {
            status.innerText = "Please enter an address";
            return;
        }

        // 2. Prepend https:// for the connection test only
        const baseUrl = `https://${rawValue}`;
        status.innerText = "Testing connection...";

        try {
            // Testing your /check endpoint
            const response = await fetch(`${baseUrl}/check`, { method: 'GET' }); 
            
            if (response.ok) {
                isConnected = true;
                connectLabel.innerText = "Disconnect";
                urlInput.disabled = true;
                if (saveBtn) saveBtn.disabled = false;
                
                if (credentialsSection) {
                    credentialsSection.classList.remove("disabled");
                    const inputs = credentialsSection.querySelectorAll("input");
                    inputs.forEach(i => i.disabled = false);
                }
                
                status.innerText = "✅ Connected to Secure Server";
                status.style.color = "green";
            } else {
                throw new Error();
            }
        } catch (err) {
            status.innerText = "❌ Connection Failed. Ensure address is correct and supports HTTPS. Please contact your system administrator for support.";
            status.style.color = "red";
        }
    } else {
        // DISCONNECT logic
        isConnected = false;
        connectLabel.innerText = "Connect";
        urlInput.disabled = false;
        if (saveBtn) saveBtn.disabled = true;
        if (credentialsSection) {
            credentialsSection.classList.add("disabled");
            const inputs = credentialsSection.querySelectorAll("input");
            inputs.forEach(i => i.disabled = true);
        }
        status.innerText = "Disconnected";
        status.style.color = "black";
    }
}

async function saveServerToExcel() {
    const serverUrl = document.getElementById("server-url").value.trim() || "";
    
    await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        let settingsSheet = sheets.getItemOrNullObject("Settings");
        await context.sync();

        if (settingsSheet.isNullObject) {
            settingsSheet = sheets.add("Settings");
            await context.sync();
        }

        //settingsSheet.protection.unprotect("dogi");
        settingsSheet.visibility = Excel.SheetVisibility.veryHidden;

        const range = settingsSheet.getRange("A1:B2");
        range.values = [
            ["Setting", "Value"],
            ["server", serverUrl]
        ];

        //settingsSheet.protection.protect({
        //    allowEditObjects: false,
        //    password: "dogi" 
        //});

        context.workbook.save();
        await context.sync();
    }).catch(error => console.error(error));
}
// src/login/login.js

async function loadSettingsFromExcel() {
    await Excel.run(async (context) => {
        const settingsSheet = context.workbook.worksheets.getItemOrNullObject("Settings");
        // Get Server (A2/B2), UserID (A3/B3), Username (A4/B4), InstName (A5/B5)
        const range = settingsSheet.getRange("A2:B5"); 
        range.load("values");
        await context.sync();

        if (!settingsSheet.isNullObject) {
            const values = range.values;
            const savedServer = values[0][1];   // B2
            const savedUserID = values[1][1];   // B3
            const savedUsername = values[2][1]; // B4
            const savedInstName = values[3][1]; // B5

            // 1. If B2 (Server) is missing: Stay on Login
            if (!savedServer) return;

            document.getElementById("server-url").value = savedServer;

            // 2. If B2 exists, automatically connect
            await toggleConnection();

            // 3. If everything (B2-B5) exists, skip to dashboard
            if (savedServer && savedUserID && savedUsername && savedInstName) {
                const status = document.getElementById("status-message");
                status.innerText = `Welcome back, ${savedInstName}. Resuming session...`;
                status.style.color = "green";

                // Ensure localStorage is ready for the dashboard
                localStorage.setItem("registrar_url", savedServer);
                localStorage.setItem("user_id", savedUserID);
                localStorage.setItem("username", savedUsername);
                localStorage.setItem("instructor_name", savedInstName);

                setTimeout(() => {
                    window.location.href = "dashboard.html";
                }, 1000);
            } else {
                // SERVER FOUND BUT SESSION INCOMPLETE: Pre-fill username if available
                if (savedUsername) {
                    const usernameField = document.getElementById("username");
                    usernameField.value = savedUsername;
                    usernameField.disabled = true; // This prevents user interaction
                }
            }
        }
    }).catch(() => {
        // Settings sheet doesn't exist, user stays on login for initial setup
    });
}

async function handleLogin() {
    const user = document.getElementById("username").value;
    const pass = document.getElementById("password").value;
    const status = document.getElementById("status-message");
    const rawAddress = document.getElementById("server-url").value.trim();
    const baseUrl = `https://${rawAddress}`;

    if (!user || !pass) {
        status.innerText = "Please enter username and password.";
        return;
    }

    // --- NEW VALIDATION BLOCK ---
    var registeredUser = ""
    const isUserValid = await Excel.run(async (context) => {
        const settingsSheet = context.workbook.worksheets.getItemOrNullObject("Settings");
        settingsSheet.load("isNullObject");
        await context.sync();

        if (!settingsSheet.isNullObject) {
            const userRange = settingsSheet.getRange("B4");
            userRange.load("values");
            await context.sync();
            
            registeredUser = String(userRange.values[0][0]).trim();
  
            // If C3 has a value and it doesn't match the input username
            if (registeredUser && registeredUser !== "" && user !== registeredUser) {
                return false; 
            }
        }
        return true; // Sheet doesn't exist or user matches
    });

    if (!isUserValid) {
        status.innerText = "Error: Username does not match the registered user for this workbook.";
        status.style.color = "red";
        const usernameField = document.getElementById("username");
        usernameField.value = registeredUser;
        usernameField.disabled = true;
        return; // STOP the login process here
    }
    console.log(isUserValid);
    saveServerToExcel() 
    status.innerText = "Authenticating...";

    try {
        const response = await fetch(`${baseUrl}/apilogin`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ username: user, password: pass })
        });

        if (response.ok) {
            const data = await response.json();
            
            // Store session info
            localStorage.setItem("access_token", data.access_token);
            localStorage.setItem("registrar_url", rawAddress);
            localStorage.setItem("instructor_name", data.instructor);
            localStorage.setItem("user_id", data.user_id);
            localStorage.setItem("username", user);

            setProgress(5);
            status.innerText = "Login successful! Preparing your Class Record...";

            // 1. Update Settings Sheet but keep it visible/hidden (not VeryHidden) for now
            await Excel.run(async (context) => {
                const sheets = context.workbook.worksheets;
                let settingsSheet = sheets.getItemOrNullObject("Settings");
                await context.sync();

                if (settingsSheet.isNullObject) {
                    settingsSheet = sheets.add("Settings");
                    await context.sync();
                }

                //settingsSheet.protection.unprotect("dogi");
                
                const uId = data.user_id ? String(data.user_id) : "";
                const uName = user ? String(user) : "";
                const iName = data.instructor ? String(data.instructor) : "";

                settingsSheet.getRange("A3:B5").values = [
                    ["userid", uId],
                    ["username", uName],
                    ["instname", iName]
                ];

                await context.sync();
            });

 
            setProgress(15);
            status.innerText = "Retrieving batches data...";
            await refreshBatchlistData();
            setProgress(20);
            status.innerText = "Retrieving instructors data...";
            await refreshInstructorData();
            setProgress(27);
            status.innerText = "Retrieving enrollment data...";
            await refreshEnrollmentData();
            setProgress(32);
            status.innerText = "Retrieving transcript data...";
            await refreshTranscriptData();
            setProgress(35);
            status.innerText = "Retrieving attendance data...";
            await refreshAttendanceData();
            setProgress(40);
            status.innerText = "Retrieving schedule data...";
            await refreshScheduleData();
            setProgress(47);
            status.innerText = "Retrieving class standing data...";
            await refreshClassStandingData();
            setProgress(47);
            status.innerText = "Retrieving transmutation data...";
            await refreshTransmutationData();
            setProgress(51);
            status.innerText = "Retrieving room data...";
            setProgress(57);
            await refreshRoomsData();
            status.innerText = "Retrieving subject data...";
            await refreshSubjectData();

            setProgress(61);
           // Download Templates (This creates Attendance and Gradesheet)
            status.innerText = "Downloading class record sheets templates...";
            const templateUrl = `${baseUrl}/download/ClassrecordTemplate.xlsx`; 
            const myBatches = await getAssignedBatchIds(); // The function we made earlier ["211", "214"]
            const sheetsToCopy = "Attendance,Gradesheet,Midterm,FinalTerm,TraineeList";

            setProgress(63);
            await downloadCRperBatch(templateUrl, sheetsToCopy, myBatches);
            setProgress(97);
            await downloadTemplate(templateUrl, "Advisory,InstructorSchedule,Base60", 1);
            // 3. Final Step: Verify Attendance exists, activate it, hide Settings, and remove Sheet1
            await Excel.run(async (context) => {
                const sheets = context.workbook.worksheets;
                const attendance = sheets.getItemOrNullObject("Attendance");
                const sheet1 = sheets.getItemOrNullObject("Sheet1");
                const notepad = sheets.getItemOrNullObject("Notepad");
                const settings = sheets.getItem("Settings");
                await context.sync();

                settings.visibility = Excel.SheetVisibility.veryHidden;
                context.workbook.save();
                await context.sync();
            });
            setProgress(100);
            status.innerText = "Setup complete. Opening dashboard...";
            setTimeout(() => { window.location.href = "dashboard.html"; }, 1500);

        } else {
            const errorData = await response.json();
            status.innerText = "Login failed: " + (errorData.message || "Invalid credentials");
        }
    } catch (error) {
        status.innerText = "❌ Error: " + error.message;
        console.error(error);
    }
}


/**
 * Generic function to sync API data to a named Excel Table
 * @param {string} endpoint - The API route (e.g., '/fl_get_batchlist')
 * @param {string} sheetName - Target worksheet name
 * @param {string} tableName - Target Excel table name
 */
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
            
            context.workbook.application.suspendScreenUpdatingUntilNextSync()
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
            'classstanding', 
            'ClassStanding', 
            { InstructorId: instructorId }
        );

        //console.log("Class Standing sync complete!");
        
    } catch (error) {
        console.error("Class Standing Sync Error:", error.message);
        throw error; // Re-throw so your UI can catch it and display the error
    }
}

async function refreshTranscriptData() {
    await syncTableFromApi('/fl_get_transcript', 'transcripttab', 'TranscriptTab');
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
                    
                    // Logic for formula injection remains the same...
                    switch (baseName) {
                        case "Attendance":
                            newlyAddedSheet.getRange("B6").values = [[batchId]]; 
                            newlyAddedSheet.getRange("A5").formulas = [[`=XLOOKUP(B6, BatchlistTab[batchid], BatchlistTab[batchname])`]];
                            newlyAddedSheet.getRange("A15").formulas = [[`=FILTER(HSTACK(EnrollmentTab[idnumber], LEFT(EnrollmentTab[gender], 1), EnrollmentTab[lastname] & ", " & EnrollmentTab[firstname] & " " & IF(AND(EnrollmentTab[middlename]<>".", EnrollmentTab[middlename]<>""), LEFT(EnrollmentTab[middlename], 1) & ". ", "")), EnrollmentTab[batchid]=B6)`]];
                            newlyAddedSheet.getRange("B7").formulas = [[`=XLOOKUP(Settings!B3, InstructorsTab[idnumber], InstructorsTab[Firstname] & " " & LEFT(InstructorsTab[Middlename], 1) & ". " & InstructorsTab[Lastname] & IF(InstructorsTab[Suffix]<>"", ", " & InstructorsTab[Suffix], ""))`]];
                            newlyAddedSheet.getRange("B8").formulas = [[`=XLOOKUP(XLOOKUP(B6, BatchlistTab[batchid], BatchlistTab[adviser]), InstructorsTab[idnumber], InstructorsTab[Firstname] & " " & LEFT(InstructorsTab[Middlename], 1) & ". " & InstructorsTab[Lastname] & IF(InstructorsTab[Suffix]<>"", ", " & InstructorsTab[Suffix], ""))`]];
                            newlyAddedSheet.getRange("B9").formulas = [[`=XLOOKUP(XLOOKUP(1, (TranscriptTab[BatchID]=B6) * (TranscriptTab[instructorid]=Settings!B3), TranscriptTab[subjectno]), SubjectTab[subjectno], SubjectTab[subjectcode])`]];
                            newlyAddedSheet.getRange("B10").formulas = [[`=UPPER(XLOOKUP(XLOOKUP(1, (TranscriptTab[BatchID]=B6) * (TranscriptTab[instructorid]=Settings!B3), TranscriptTab[subjectno]), SubjectTab[subjectno], SubjectTab[subjecttitle]))`]];
                            newlyAddedSheet.getRange("E8").formulas = [[`=XLOOKUP(B6, batchlisttab[batchid], batchlisttab[trainingstart])`]];
                            newlyAddedSheet.getRange("E12").formulas = [[`=XLOOKUP(B6, batchlisttab[batchid], batchlisttab[midtermexamdate])`]];
                            newlyAddedSheet.getRange("E9").formulas = [[`=XLOOKUP(B6, batchlisttab[batchid], batchlisttab[trainingend])`]];
                            newlyAddedSheet.getRange("F12").formulas = [[`=XLOOKUP(B6, batchlisttab[batchid], batchlisttab[finaltermexamdate])`]];

                            break;
                        case "Gradesheet":
                            newlyAddedSheet.getRange("K15").values = [[batchId]]; 
                            newlyAddedSheet.getRange("B21").formulas = [[`=FILTER(HSTACK(EnrollmentTab[lastname] & ", " & EnrollmentTab[firstname] & " " & IF(AND(EnrollmentTab[middlename]<>".", EnrollmentTab[middlename]<>""), LEFT(EnrollmentTab[middlename], 1) & ". ", "")), EnrollmentTab[batchid]=K15)`]];
                            newlyAddedSheet.getRange("A8").formulas = [[`=XLOOKUP(XLOOKUP(1, (TranscriptTab[BatchID]=K15) * (TranscriptTab[instructorid]=Settings!B3), TranscriptTab[subjectno]), SubjectTab[subjectno], SubjectTab[subjectcode])`]];
                            newlyAddedSheet.getRange("A11").formulas = [[`=UPPER(XLOOKUP(XLOOKUP(1, (TranscriptTab[BatchID]=K15) * (TranscriptTab[instructorid]=Settings!B3), TranscriptTab[subjectno]), SubjectTab[subjectno], SubjectTab[subjecttitle]))`]];
                            newlyAddedSheet.getRange("C8").formulas = [[`=XLOOKUP(K15, batchlisttab[batchid], batchlisttab[year])`]];
                            newlyAddedSheet.getRange("C11").formulas = [[`=XLOOKUP(K15, batchlisttab[batchid], batchlisttab[period])`]];
                            newlyAddedSheet.getRange("C14").formulas = [[`=XLOOKUP(XLOOKUP(K15, BatchlistTab[batchid], BatchlistTab[adviser]), InstructorsTab[idnumber], InstructorsTab[Firstname] & " " & LEFT(InstructorsTab[Middlename], 1) & ". " & InstructorsTab[Lastname] & IF(InstructorsTab[Suffix]<>"", ", " & InstructorsTab[Suffix], ""))`]];
                            newlyAddedSheet.getRange("I8").formulas = [[`=XLOOKUP(K15, BatchlistTab[batchid], BatchlistTab[batchname])`]];
                            newlyAddedSheet.getRange("I11").formulas = [[`=XLOOKUP(Settings!B3, InstructorsTab[idnumber], InstructorsTab[Firstname] & " " & LEFT(InstructorsTab[Middlename], 1) & ". " & InstructorsTab[Lastname] & IF(InstructorsTab[Suffix]<>"", ", " & InstructorsTab[Suffix], ""))`]];
                            break;
                        case "Midterm":
                        case "FinalTerm":
                            newlyAddedSheet.getRange("B21").formulas = [[`=FILTER(HSTACK(LEFT(EnrollmentTab[gender], 1), EnrollmentTab[idnumber], EnrollmentTab[lastname] & ", " & EnrollmentTab[firstname]), EnrollmentTab[batchid]=${batchId})`]];
                            break;
                        case "TraineeList":
                            newlyAddedSheet.getRange("B16").formulas = [[`=FILTER(EnrollmentTab[lastname] & ", " & EnrollmentTab[firstname], (EnrollmentTab[batchid]=${batchId})*(EnrollmentTab[gender]="Male"),"")`]];
                            newlyAddedSheet.getRange("E16").formulas = [[`=FILTER(EnrollmentTab[lastname] & ", " & EnrollmentTab[firstname], (EnrollmentTab[batchid]=${batchId})*(EnrollmentTab[gender]="Female"),"")`]];
                            break;
                        default:
                           //console.log("Created a general sheet with no specific template logic.");
                    }
                    // Sync again to finalize the name change before the next loop iteration
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
 * Returns an array of unique Batch IDs assigned to the current instructor
 * by querying the local 'transcripttable'.
 */
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