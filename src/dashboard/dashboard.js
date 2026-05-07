// @ts-check
/* global document, Office, Excel, localStorage, window */
let authCallback = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.onActivated.add(onSheetActivated);
            await context.sync();
            //console.log("Event listener registered successfully."); // This only logs ONCE at startup
        });

        initializeDashboard();
        initScheduleMonitor();


        // --- BUTTON SELECTIONS ---
        const submitBtn = document.getElementById('submit-gradesheet');
        const refreshBtn = document.getElementById("refresh-batch-view");
        const dropdown = document.getElementById("class-dropdown");
        const disconnectBtn = document.getElementById("disconnect-button");
        const syncBtn = document.getElementById('sync-attendance');
        const syncUpdatesBtn = document.getElementById("sync-classupdates");
        const confirmBtn = document.getElementById("confirm-submit");
        const closeBtn = document.getElementById('close-panel');
        const authOverlay = document.getElementById('auth-overlay');
        const syncSched = document.getElementById('sync-instsched');

        // --- REFRESH BATCH VIEW ---
        if (refreshBtn) {
            refreshBtn.onclick = async () => {
                const selectedValue = dropdown.value;
                if (selectedValue && selectedValue !== "loading") {
                    const status = document.getElementById("status-message");
                    status.innerText = "Refreshing view...";
                    await handleBatchSelection(selectedValue);
                    status.innerText = "View Refreshed";
                    setTimeout(() => { status.innerText = ""; }, 2000);
                }
            };
        }
if (syncSched) {
            syncSched.onclick = async () => {
                // 1. Trigger the confirmation dialog
                const confirmed = await showConfirmDialog(
                    "Refresh Schedule?",
                    "Run this only when needed, such as when there are changes or updates to your schedule.\n\nNote: You will also need to click the 'Sync Attendance' button afterward to reformat your attendance sheets."
                );

                // 2. Proceed only if the user clicked "OK"
                if (confirmed) {
                    const status = document.getElementById("status-message");
                    status.innerText = "Refreshing schedule...";
                    
                    const userid = await getSettingValue(3);
                    await createSchedule(userid);
                    
                    status.innerText = "✅ Schedule refreshed.";
                    setTimeout(() => { status.innerText = ""; }, 3000);
                }
            };
        }
        // --- AUTH PANEL CONTROLS ---
        if (submitBtn) {
            submitBtn.onclick = () => showAuthOverlay(() => handleSubmitGrades());
        }

        if (closeBtn) closeBtn.onclick = hidePanel;
        
        if (authOverlay) {
            authOverlay.onclick = (e) => {
                if (e.target === authOverlay) hidePanel();
            };
        }

        if (confirmBtn) {
            confirmBtn.onclick = async () => {
                const passField = document.getElementById("panel-pass");
                const pass = passField.value;
                const user = await getSettingValue(4);
                const rawAddress = await getSettingValue(2);
                const mainStatus = document.getElementById("status-message");
                const authStatus = document.getElementById("auth-status"); // Target the internal status

                if (!pass) {
                    authStatus.innerText = "❌ Password is required.";
                    authStatus.style.color = "red";
                    return;
                }

                confirmBtn.disabled = true;
                //confirmBtn.innerText = "Verifying...";
                authStatus.innerText = "Verifying...";
                authStatus.style.color = "#0078d4";

                try {
                    const response = await fetch(`https://${rawAddress}/apilogin`, {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ username: user, password: pass })
                    });

                    if (response.ok) {
                        const data = await response.json();
                        await setWorkbookSetting(6, data.access_token);
                    
                        passField.value = "";
                        authStatus.innerText = ""; // Clear internal status
                        hidePanel();
                    
                        if (authCallback) {
                            // Switch back to main status since the panel is now hidden
                            mainStatus.innerText = "✅ Verified. Resuming task...";
                            mainStatus.style.color = "green";
                            await authCallback();
                            authCallback = null;
                        }
                    } else {
                        authStatus.innerText = "❌ Incorrect password. Try again.";
                        authStatus.style.color = "red";
                        passField.value = "";
                        passField.focus();
                    }
                } catch (err) {
                    console.error("Re-auth Error:", err);
                    authStatus.innerText = "❌ Connection error.";
                    authStatus.style.color = "red";
                } finally {
                    confirmBtn.disabled = false;
                    confirmBtn.innerText = "Submit";
                }
            };
        }
        
        if (closeBtn) {
            closeBtn.onclick = () => hidePanel(true); // Pass true for cancellation
        }

        if (authOverlay) {
            authOverlay.onclick = (e) => {
                if (e.target === authOverlay) {
                    hidePanel(true); // Pass true for cancellation
                }
            };
        }
        
        // --- SYNC CLASS UPDATES ---
// --- SYNC CLASS UPDATES ---
        if (syncUpdatesBtn) {
            syncUpdatesBtn.onclick = async () => {
                // 1. Trigger the critical warning dialog
                const confirmed = await showConfirmDialog(
                    "Download Class Updates?",
                    "⚠️ WARNING: Use this only on an incomplete sync during initialization, or when the class record has not been used yet.\n\nThis command downloads new updates from the server that might overwrite your current manual input data. Proceed?"
                );

                // 2. Only execute the sync if the user explicitly confirms
                if (confirmed) {
                    const status = document.getElementById("status-message");
                    const rawAddress = await getSettingValue(2);
                    const baseUrl = `https://${rawAddress}`;
                    
                    status.innerText = "Starting class updates...";

                    const startSync = async () => {
                        try {
                            await performFullSync(setProgress, status, baseUrl);
                            await postSyncCleanup();
                        } catch (error) {
                            if (error.message.includes("401") || error.message.includes("Missing")) {
                                status.innerText = "Authentication required...";
                                showAuthOverlay(async () => {
                                    status.innerText = "Resuming sync...";
                                    await startSync(); 
                                });
                            } else {
                                status.innerText = "❌ Sync Failed: " + error.message;
                                status.style.color = "red";
                            }
                        }
                    };

                    await startSync();
                }
            };
        }

        if (disconnectBtn) disconnectBtn.onclick = handleDisconnect;
        if (syncBtn) syncBtn.onclick = syncAttendance;
    

        const manualRawToggle = document.getElementById("manual-raw-toggle");
        const manualAutoToggle = document.getElementById("manual-auto-toggle");

if (manualRawToggle && manualAutoToggle) {
const handleGradeToggle = async (mode, activeToggle, otherToggle) => {
    const status = document.getElementById("status-message");
    
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const props = sheet.customProperties;
        
        // Retrieve batchid and sheetType from the sheet itself
        const bProp = props.getItemOrNullObject("batchid");
        const tProp = props.getItemOrNullObject("sheetType");
        
        bProp.load("value");
        tProp.load("value");
        await context.sync();

        // REPLACEMENT VALIDATION: Use Custom Properties
        if (bProp.isNullObject || tProp.value !== "gradesheet_record") {
            status.innerText = "⚠️ This action is only available on a valid Gradesheet.";
            status.style.color = "orange";
            activeToggle.checked = !activeToggle.checked; // Revert UI
            return;
        }

        const batchId = bProp.value;
        const targetRange = mode === "RAW" ? "N20:Q51" : "C20:K51";

        // DIALOG LOGIC: Triggered when UNCHECKING
        if (!activeToggle.checked) {
            const confirmed = await showConfirmDialog(
                `⚠️ Restore Automatic Calculations?`,
                `This will clear your existing manual grades in the ${mode} range to be replaced by automatic formulas. This cannot be undone. Continue?`
            );
            
            if (!confirmed) {
                activeToggle.checked = true;
                return;
            }
        }

        if (activeToggle.checked) otherToggle.checked = false;

        status.innerText = "Processing...";
        status.style.color = "#43484c";

        try {
            await injectSheetFormulas(context, sheet, "Gradesheet", batchId);
            sheet.getRange("N20:Q51").format.fill.color = "#ffffff"
            sheet.getRange("C20:K51").format.fill.color = "#ffffff"
            if (activeToggle.checked) {
                // User enabled Manual Mode: Clear for fresh input[cite: 3]
                sheet.getRange(targetRange).clear(Excel.ClearApplyTo.contents);
                sheet.getRange(targetRange).format.fill.color = "#eff6fd";
                status.innerText = `✅ Manual ${mode} mode enabled.`;
            } else {
                
                status.innerText = `✅ Automatic calculations restored.`;
            }

            await context.sync();
            status.style.color = "green";
        } catch (err) {
            status.innerText = "❌ Error: " + err.message;
            status.style.color = "red";
            activeToggle.checked = !activeToggle.checked;
        }
    });
};

    manualRawToggle.onchange = () => handleGradeToggle("RAW", manualRawToggle, manualAutoToggle);
    manualAutoToggle.onchange = () => handleGradeToggle("GP", manualAutoToggle, manualRawToggle);
}
        
    }

});

// --- HELPER FUNCTIONS (Outside Office.onReady for access) ---

async function onSheetActivated(event) {
    const manualRawToggle = document.getElementById("manual-raw-toggle");
    const manualAutoToggle = document.getElementById("manual-auto-toggle");
    const submitGradesBtn = document.getElementById("submit-gradesheet");

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(event.worksheetId);
        const props = sheet.customProperties;
        
        const tProp = props.getItemOrNullObject("sheetType");
        tProp.load("value");
        await context.sync();

        // Check if we are on a Gradesheet
        const isGradesheet = !tProp.isNullObject && tProp.value === "gradesheet_record";

        // Enable/Disable the Submit button dynamically
if (submitGradesBtn) {
            if (!isGradesheet) {
                // Not a gradesheet -> Disable button
                submitGradesBtn.disabled = true;
                submitGradesBtn.style.opacity = "0.5"; // Make it look faded
                submitGradesBtn.style.cursor = "not-allowed"; // Change mouse pointer
                submitGradesBtn.style.pointerEvents = "none"; // Prevent all clicks entirely
            } else {
                // Is a gradesheet -> Enable button
                submitGradesBtn.disabled = false;
                submitGradesBtn.style.opacity = "1"; // Restore full color
                submitGradesBtn.style.cursor = "pointer"; // Restore normal pointer
                submitGradesBtn.style.pointerEvents = "auto"; // Allow clicks
            }
        }

        if (manualRawToggle && manualAutoToggle) {
            // Enable toggles only if on a gradesheet, otherwise disable
            manualRawToggle.disabled = !isGradesheet;
            manualAutoToggle.disabled = !isGradesheet;

            if (isGradesheet) {
                const rangeRaw = sheet.getRange("N21");
                const rangeGP = sheet.getRange("C21");

                rangeRaw.load("formulas");
                rangeGP.load("formulas");
                await context.sync();

                // Sync the UI checkmarks with the existing cell state
                const rawFormulaValue = rangeRaw.formulas[0][0];
                const gpFormulaValue = rangeGP.formulas[0][0];

                manualRawToggle.checked = !String(rawFormulaValue).startsWith("=");
                manualAutoToggle.checked = !String(gpFormulaValue).startsWith("=");
            } else {
                // Clear checks if on a non-gradesheet to avoid confusion
                manualRawToggle.checked = false;
                manualAutoToggle.checked = false;
            }
        }
    });
}
async function findSheetByMetadata(context, batchId, sheetType) {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name, items/customProperties");
    await context.sync();

    for (let sheet of worksheets.items) {
        const bProp = sheet.customProperties.getItemOrNullObject("batchid");
        const tProp = sheet.customProperties.getItemOrNullObject("sheetType");
        
        bProp.load("value");
        tProp.load("value");
        await context.sync();

        if (!bProp.isNullObject && !tProp.isNullObject) {
            if (String(bProp.value) === String(batchId) && tProp.value === sheetType) {
                return sheet;
            }
        }
    }
    return null;
}

function showAuthOverlay(onSuccessCallback) {
    const authOverlay = document.getElementById('auth-overlay');
    const authPanel = document.getElementById('auth-panel');
    authCallback = onSuccessCallback; 
    
    authOverlay.style.setProperty('display', 'flex', 'important');
    setTimeout(() => {
        authPanel.classList.add('show');
        const passInput = document.getElementById("panel-pass");
        if (passInput) passInput.focus();
    }, 50);
}

function hidePanel(isCancellation = false) {
    const authOverlay = document.getElementById('auth-overlay');
    const authPanel = document.getElementById('auth-panel');
    const status = document.getElementById("status-message");

    authPanel.classList.remove('show');
    
    if (isCancellation) {
        status.innerText = "⚠️ Sync cancelled by user.";
        status.style.color = "#ffa500"; // Orange
         setProgress(0);
        authCallback = null; // Clear the callback so it doesn't run later
    }

    setTimeout(() => {
        authOverlay.style.display = 'none';
    }, 300);
}

async function postSyncCleanup() {
    const status = document.getElementById("status-message");
    await initializeDashboard();
    await hideAllBatchSheets();
    status.innerText = "Sync complete.";
    status.style.color = "green";

    setTimeout(() => {
        setProgress(0);
        status.innerText = "Ready";
        status.style.color = "#605e5c"; // Neutral gray
    }, 5000);
}

async function handleSubmitGrades() {
    const status = document.getElementById("status-message");
    
    // Improved Prompt Message
    const confirmed = await showConfirmDialog(
        "Submit Grades to Registrar?",
        "You are about to submit this gradesheet. Any existing unfinalized grades in the system for these trainees will be overwritten by the current sheet data.\n\nDo you wish to continue?"
    );

    if (!confirmed) return;

    // Check token right away
    let token = await getSettingValue(6);
    if (!token) {
        showAuthOverlay(() => handleSubmitGrades());
        return;
    }

    status.innerText = "Processing gradesheet submission...";
    status.style.color = "#43484c";

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            // SECURITY CHECK: Double-verify they are on a gradesheet
            const tProp = sheet.customProperties.getItemOrNullObject("sheetType");
            tProp.load("value");
            await context.sync();

            if (tProp.isNullObject || tProp.value !== "gradesheet_record") {
                status.innerText = "⚠️ This action requires an active Gradesheet.";
                status.style.color = "orange";
                return;
            }

            // PERFORMANCE TRICK: Read rows 21 down to 80 all at once to minimize Excel.run calls.
            // B to M contains Name (B), GP's (C,D,E,F), Temps (G,H), and RecordID (M).
            const dataRange = sheet.getRange("B21:M80"); 
            dataRange.load("values");
            await context.sync();

            const userId = await getSettingValue(4); // ModBy user
            const payload = [];
            const rows = dataRange.values;

            // Map columns to their respective indexes:
            // B=0, C=1, D=2, E=3, F=4, G=5, H=6 ... M=11
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                const traineeName = row[0];

                // "Iterate until there is no data in Bx"
                if (!traineeName || String(traineeName).trim() === "") {
                    break; 
                }

                const recordId = row[11];
                if (!recordId) continue; // Skip if no RecordID exists

                payload.push({
                    ModBy: userId,
                    MidLecGP: row[1],
                    MidLabGP: row[2],
                    FinLecGP: row[3],
                    FinLabGP: row[4],
                    LecGradeTemp: row[5],
                    LabGradeTemp: row[6],
                    RecordId: recordId
                });
            }

            if (payload.length === 0) {
                status.innerText = "⚠️ No valid grades found to submit.";
                status.style.color = "orange";
                return;
            }

            status.innerText = `Uploading ${payload.length} records...`;

            // API Upload
            const rawAddress = await getSettingValue(2);
            const response = await fetch(`https://${rawAddress}/submit-grades`, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${token}`
                },
                body: JSON.stringify(payload)
            });

            if (response.ok) {
                status.innerText = "✅ Gradesheet submitted successfully.";
                status.style.color = "green";
                setTimeout(() => { status.innerText = ""; }, 5000);
            } else if (response.status === 401 || response.status === 403) {
                status.innerText = "Authentication expired...";
                showAuthOverlay(() => handleSubmitGrades());
            } else {
                throw new Error(`Server returned ${response.status}`);
            }
        });
    } catch (error) {
        console.error("Submission Error:", error);
        status.innerText = "❌ Submission failed: " + error.message;
        status.style.color = "red";
    }
}

async function initializeDashboard() {
   
    const nameDisplay = document.getElementById("instructor-name");
    const dropdown = document.getElementById("class-dropdown");

    const currentUrl = await getSettingValue(2);
    if (currentUrl === "render-demoaddin-api.onrender.com") {
        document.getElementById("demo-strip").style.display = "block";
    }
    dropdown.onchange = () => handleBatchSelection(dropdown.value);
    
    const instructorName = await getSettingValue(5);
    const instructorId = await getSettingValue(3);

    if (instructorName) nameDisplay.innerText = instructorName;
    if (!instructorId) return;

    try {
        await Excel.run(async (context) => {
            const transcriptTable = context.workbook.tables.getItem("transcripttab");
            const batchTable = context.workbook.tables.getItem("batchlisttab");

            const transcriptBody = transcriptTable.getDataBodyRange();
            const transcriptHeader = transcriptTable.getHeaderRowRange();
            const batchBody = batchTable.getDataBodyRange();
            const batchHeader = batchTable.getHeaderRowRange();

            transcriptBody.load("values");
            transcriptHeader.load("values");
            batchBody.load("values");
            batchHeader.load("values");

            await context.sync();

            // 1. Normalize Headers to Lowercase
            const transHeaders = transcriptHeader.values[0].map(h => String(h).toLowerCase().trim());
            const batchHeaders = batchHeader.values[0].map(h => String(h).toLowerCase().trim());

            const idxTransInstructor = transHeaders.indexOf("instructorid");
            const idxTransBatch = transHeaders.indexOf("batchid");
            const idxBatchId = batchHeaders.indexOf("batchid");
            const idxBatchName = batchHeaders.indexOf("batchname");

            // Check if columns exist
            if (idxTransInstructor === -1 || idxTransBatch === -1) {
                console.error("Missing columns in transcripttable. Found headers:", transHeaders);
                return;
            }

            // 2. Build the Set of assigned Batch IDs
            const assignedBatchIds = new Set();
            const targetInstructor = String(instructorId).trim();

            transcriptBody.values.forEach(row => {
                const rowInstructor = String(row[idxTransInstructor]).trim();
                const rowBatch = row[idxTransBatch] ? String(row[idxTransBatch]).trim() : "";
                
                // Only add if it's a real value and matches instructor
                if (rowInstructor === targetInstructor && rowBatch !== "" && rowBatch !== "undefined") {
                    assignedBatchIds.add(rowBatch);
                }
            });

            //console.log("Assigned Batch IDs found:", Array.from(assignedBatchIds));

            // 3. Map to Names from the Batch List
            const assignedClasses = batchBody.values
                .filter(row => {
                    const bId = String(row[idxBatchId]).trim();
                    return assignedBatchIds.has(bId);
                })
                .map(row => ({
                    id: String(row[idxBatchId]).trim(),
                    name: String(row[idxBatchName]).trim()
                }))
                .sort((a, b) => a.name.localeCompare(b.name));

            updateDropdown(dropdown, assignedClasses);
        });
    } catch (error) {
        console.error("Dashboard Error:", error);
        dropdown.innerHTML = '<option>Error loading classes</option>';
    }
}

function updateDropdown(dropdown, classes) {
    // This line prevents the list from just growing longer and longer
    dropdown.innerHTML = '<option value="" disabled selected>Select a class...</option>';
    
    if (classes.length === 0) {
        dropdown.innerHTML = '<option value="">No classes assigned</option>';
        return;
    }

    classes.forEach(item => {
        const option = document.createElement("option");
        option.value = item.id;
        option.text = item.name;
        dropdown.add(option);
    });
}

async function handleDisconnect() {
    try {
        // 1. Clear the specific metadata settings
        await clearWorkbookSetting(6); // Token
        await clearWorkbookSetting(5); // Instructor Name
        await clearWorkbookSetting(3); // Instructor ID

        // 2. Excel Workbook Cleanup
        await Excel.run(async (context) => {
            const settingsSheet = context.workbook.worksheets.getItem("Settings");
            const worksheets = context.workbook.worksheets;

            // Load worksheet names so we can iterate through them
            worksheets.load("items/name");

            // Clear the range B3:B5 in the Settings sheet
            // This clears UserID, Username, and InstName but leaves B2 (Server URL)
            settingsSheet.getRange("B5:B5").clear();

            await context.sync();

            // 3. Make all hidden class sheets visible again before disconnecting
            // This ensures the next user doesn't start with a broken/empty UI
            worksheets.items.forEach((sheet) => {
                if (sheet.name.includes("_")) {
                    sheet.visibility = Excel.SheetVisibility.visible;
                }
            });

            await context.sync();
        });

        // 4. Redirect back to login page
        window.location.href = "login.html";

    } catch (error) {
        console.error("Disconnect Error: ", error);
        // Fallback: Ensure the user is redirected even if Excel logic fails
        window.location.href = "login.html";
    }
}

async function hideAllBatchSheets() {
  await Excel.run(async (context) => {
    // Retrieve the collection of all worksheets in the workbook
    const worksheets = context.workbook.worksheets;
    
    // Load the 'name' property for all sheets so we can inspect them
    worksheets.load("items/name");

    await context.sync();

    // Iterate through the sheets
    worksheets.items.forEach((sheet) => {
      if (sheet.name.includes("_")) {
        // Set visibility to Hidden (user can unhide) 
        // or 'VeryHidden' (can only be unhidden via code)
        sheet.visibility = Excel.SheetVisibility.hidden;
      }
    });

    await context.sync();
    //console.log("Process complete.");
  });
}

async function handleBatchSelection(selectedBatchId) {
    if (!selectedBatchId) return;

    try {
        await Excel.run(async (context) => {
            // 1. Freeze UI
            context.workbook.application.suspendScreenUpdatingUntilNextSync();
            
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name");
            await context.sync();

            // 2. Map sheets to their property objects and load them
            const sheetProcessingList = worksheets.items
                .filter(sheet => sheet.name.includes("_"))
                .map(sheet => {
                    const batchProp = sheet.customProperties.getItemOrNullObject("batchid");
                    batchProp.load("value");
                    return { sheet, batchProp }; // Keep them paired up
                });
            
            // 3. Sync once to get all property values
            await context.sync();

            // 4. Perform toggles using the paired list
            for (const item of sheetProcessingList) {
                // Now item.batchProp.value and item.batchProp.isNullObject are ready
                if (!item.batchProp.isNullObject && String(item.batchProp.value) === String(selectedBatchId)) {
                    item.sheet.visibility = Excel.SheetVisibility.visible;
                } else {
                    item.sheet.visibility = Excel.SheetVisibility.hidden;
                }
            }

            // 5. Push all changes at once
            await context.sync();
        });

        // 6. Move to the sheet
        await activateAttendanceByBatch(selectedBatchId);

    } catch (error) {
        console.error("Error toggling sheets via metadata:", error);
    }
}


async function activateAttendanceByBatch(targetBatchId) {
    await Excel.run(async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name"); 
        await context.sync();

        let targetSheet = null;

        // Use a standard loop for better compatibility and scope control
        for (let i = 0; i < worksheets.items.length; i++) {
            let currentSheet = worksheets.items[i];
            
            // Load the custom properties for this specific sheet
            const props = currentSheet.customProperties;
            const batchProp = props.getItemOrNullObject("batchid");
            const typeProp = props.getItemOrNullObject("sheetType");

            batchProp.load("value");
            typeProp.load("value");
            
            // We must sync inside the loop to check the values of these properties
            await context.sync();

            if (!batchProp.isNullObject && !typeProp.isNullObject) {
                const isCorrectBatch = String(batchProp.value) === String(targetBatchId);
                const isAttendance = typeProp.value === "attendance_record";

                if (isCorrectBatch && isAttendance) {
                    targetSheet = currentSheet;
                    break; // Exit the loop as soon as we find it
                }
            }
        }

        if (targetSheet) {
            targetSheet.activate();
            // Optional: select A1 to reset the view
            targetSheet.getRange("A1").select();
            await context.sync();
        } else {
            console.warn(`Could not find Attendance sheet for Batch ID: ${targetBatchId}`);
        }
    });
}

function excelSerialToJSDate(serial) {
    // Excel base date is Dec 30, 1899. 
    // The difference between Excel and JS epochs is 25569 days.
    const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
    return date;
}

async function refreshCalendar() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            // 1. Get Values for validation (E12 and F12)
            const dateRange = sheet.getRange("E12:F12");
            dateRange.load("values");
            
            // 2. Load the header range to find the date columns (H12:OH12)
            const headerRange = sheet.getRange("H12:HO12");
            headerRange.load("values");

            await context.sync();

        // 2. Extract the numbers from the 2D arrays
            const startSerial = dateRange.values[0][0]; // The number from E12
            const endSerial = dateRange.values[0][1];   // The number from F12
const headerRow = headerRange.values[0];

    // Debugging logs
    //console.log("StartSerial:", startSerial); 

    // 2. Use .findIndex with explicit conversion to prevent Type Mismatch
    const startOffset = headerRow.findIndex(cell => {
        return cell !== "" && Math.floor(Number(cell)) === Math.floor(Number(startSerial));
    });

    const endOffset = headerRow.findIndex(cell => {
        return cell !== "" && Math.floor(Number(cell)) === Math.floor(Number(endSerial));
    });

    if (startOffset === -1) {
        // If it's still -1, let's see what is actually in the first 5 header cells
        console.warn("Match not found. First 5 headers are:", headerRow.slice(0, 5));
        return;
    }

    // 3. Map back to Excel Column Index (H is 7)
    const startColIdx = startOffset + 7;
    const endColIdx = endOffset + 7;

    //console.log(`Success! Start Column Index: ${startColIdx}`);
            
            // 5. If you NEED the JS Date for other logic (like month names), convert it AFTER finding it
            //const realStartDate = excelSerialToJSDate(startSerial);
            
            // 4. Batch set formulas (Much faster than looping rows)
            // Example: Setting the row 13 formulas
            const formulaRange = sheet.getRange("E13:F13");
            formulaRange.formulas = [[
                "=SUMPRODUCT((INDIRECT(H$1):INDIRECT(I$1)=1)*(INDIRECT(H$1):INDIRECT(I$1)<>\"\"))",
                "=SUMPRODUCT((INDIRECT(H$2):INDIRECT(I$2)=1)*(INDIRECT(H$2):INDIRECT(I$2)<>\"\"))"
            ]];

            await context.sync();
            
            // 5. Call your formatting function
            //await mergeMonthsLogic(context, sheet);
        });
    } catch (error) {
        console.error(error);
    }
}

async function mergeMonthsLogic(context, sheet) {
    // This assumes you've already found your start/end columns
    // This is a snippet of how to do the "Medium" border from your VB code
    const mergeRange = sheet.getRange("H10:Z10"); // Example range
    
    mergeRange.merge();
    mergeRange.format.horizontalAlignment = "Center";
    mergeRange.format.verticalAlignment = "Center";
    mergeRange.format.font.bold = true;

    // Apply the thick outer border
    const borders = mergeRange.format.borders;
    const sides = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"];
    
    sides.forEach(side => {
        borders.getItem(side).style = "Continuous";
        borders.getItem(side).weight = "Medium";
    });

    await context.sync();
}

function showConfirmDialog(title, message) {
    const overlay = document.getElementById('confirm-overlay');
    const titleEl = document.getElementById('confirm-title');
    const bodyEl = document.getElementById('confirm-body');
    const okBtn = document.getElementById('dialog-ok');
    const cancelBtn = document.getElementById('dialog-cancel');

    titleEl.innerText = title;
    bodyEl.innerText = message;
    overlay.style.display = 'flex';

    return new Promise((resolve) => {
        okBtn.onclick = () => {
            overlay.style.display = 'none';
            resolve(true);
        };
        cancelBtn.onclick = () => {
            overlay.style.display = 'none';
            resolve(false);
        };
    });
}
async function applyScheduleExclusions(sheet) {
    // 1. Load the Schedule Row (8) and the Action Flag Row (13)
    const scheduleRange = sheet.getRange("H8:IW8");
    const flagRange = sheet.getRange("H13:IW13");
    
    scheduleRange.load("values");
    flagRange.load("values");
    
    await sheet.context.sync();
    
    const scheduleVals = scheduleRange.values[0];
    const flagVals = flagRange.values[0];
    let hasChanges = false;
    
    // 2. Loop through and check for zeroes
    for (let i = 0; i < scheduleVals.length; i++) {
        // We explicitly check for 0 (so we don't accidentally trigger on empty blank cells "")
        if (scheduleVals[i] === 0 || scheduleVals[i] === "0") {
            // If it's a 0, but row 13 isn't already -1, update it
            if (flagVals[i] !== -1) {
                flagVals[i] = -1;
                hasChanges = true;
            }
        }
    }
    
    // 3. Write back ONLY if we made changes (Saves time and prevents Excel UI flicker)
    if (hasChanges) {
        flagRange.values = [flagVals];
        await sheet.context.sync();
        console.log("✅ Applied schedule exclusions (-1 applied to unscheduled dates).");
    }
}
async function syncAttendance() {
    const status = document.getElementById("status-message");
    const excludeOutsideSchedule = document.getElementById("exclude-outside-schedule").checked;
    const rawAddress = await getSettingValue(2);
    const token = await getSettingValue(6); // Retrieve stored token

    // Check if token exists; if not, show overlay immediately
    if (!token) {
        showAuthOverlay(async () => await syncAttendance());
        return;
    }

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const props = sheet.customProperties;
        
        // 1. Validate Sheet via Custom Properties
        const bProp = props.getItemOrNullObject("batchid");
        const tProp = props.getItemOrNullObject("sheetType");
        bProp.load("value");
        tProp.load("value");
        await context.sync();

        if (tProp.isNullObject || tProp.value !== "attendance_record") {
            status.innerText = "⚠️ Action only available on Attendance Sheets.";
            status.style.color = "orange";
            return; 
        }
        
        if (excludeOutsideSchedule) {
            status.innerText = "Applying schedule filters...";
            await applyScheduleExclusions(sheet);
        }

        status.innerText = "Preparing sync...";
        const batchId = bProp.value;
        const instructorId = await getSettingValue(3);
        await injectSheetFormulas(context, sheet, "Attendance", batchId); 
        // 2. Load Metadata from cells
        const subjectRange = sheet.getRange("F10");
        const totalTraineesRange = sheet.getRange("B13");
        subjectRange.load("values");
        totalTraineesRange.load("values");
        await context.sync();

        const subjectNo = subjectRange.values[0][0];
        const totalTrainees = parseInt(totalTraineesRange.values[0][0]);
        if (isNaN(totalTrainees) || totalTrainees <= 0) return;

        const lastRow = 15 + totalTrainees - 1;

        // 3. Load Main Data
        const dataRange = sheet.getRange(`A12:HO${lastRow}`);
        dataRange.load("values");
        await context.sync();

        const data = dataRange.values; 
        const ops = [];
        const dateRow = data[0]; 
        const flagRow = data[1]; 

        // 4. Process Loop
        for (let i = 3; i < data.length; i++) {
            const rowData = data[i];
            const traineeId = rowData[0];
            if (!traineeId) continue;

            for (let colIdx = 7; colIdx < 231; colIdx++) {
                if (flagRow[colIdx] !== 1) continue;
                const val = rowData[colIdx];
                if (val !== 1 && val !== -1) continue;

                const dateOADate = dateRow[colIdx];
                if (!dateOADate) continue;

                const dt = new Date((dateOADate - 25569) * 86400 * 1000);
                dt.setHours(dt.getHours() + 8);

                ops.push({
                    traineeid: traineeId,
                    instructorid: instructorId,
                    subjectno: subjectNo,
                    batchid: batchId,
                    timestamp: dt.toISOString().slice(0, 19).replace('T', ' '),
                    action: val === 1 ? "add" : "remove"
                });
            }
        }

        // 5. API Upload with Authorization Header
        //console.log(ops)
        let shouldRetrieve = false; // Add a flag to track if we should proceed

        if (ops.length > 0) {
            try {
                const response = await fetch(`https://${rawAddress}/attendance/batch`, {
                    method: 'POST',
                    headers: { 
                        'Content-Type': 'application/json',
                        'Authorization': `Bearer ${token}` 
                    },
                    body: JSON.stringify(ops)
                });

                if (response.ok) {
                    status.innerText = `✅ Attendance Synced`;
                    status.style.color = "green";
                    
                    
                    shouldRetrieve = true; // Sync succeeded, safe to proceed
                    
                } else if (response.status === 401 || response.status === 403) {
                    status.innerText = "Authentication expired...";
                    showAuthOverlay(async () => await syncAttendance());
                    // shouldRetrieve remains false because we are handing off to the auth overlay
                } else {
                    throw new Error(`Server error: ${response.status}`);
                }
            } catch (error) {
                status.innerText = "❌ Sync failed: " + error.message;
                status.style.color = "red";
                // shouldRetrieve remains false, skipping the fetch
            }
        } else {
            status.innerText = "No local changes to upload. Checking for updates...";
            shouldRetrieve = true; // Nothing to push, but safe to pull
        }

        // Only run if the flag was set to true
        if (shouldRetrieve) {
            await syncInstructorAttendance(context, batchId);
            await retrieveAttendanceScans(context, batchId);
        }
    });
}
async function clearAttendanceFlags(sheet, lastRow) {
    const range = sheet.getRange(`H15:HO${lastRow}`);
    range.load("values");
    await sheet.context.sync();

    const vals = range.values;
    const newVals = vals.map(row => row.map(cell => cell === -1 ? "" : cell));
    
    range.values = newVals;
    await sheet.context.sync();
}
async function syncInstructorAttendance(context, batchId) {
    const status = document.getElementById("status-message");
    try {
        const instructorId = await getSettingValue(3);
        const rawAddress = await getSettingValue(2);
        const token = await getSettingValue(6); // Retrieve token for 403 prevention
        
        // 1. Find sheets by Metadata instead of hardcoded names
        // This prevents "Resource does not exist" errors if tabs are renamed
        const wsAtn = await findSheetByMetadata(context, batchId, "attendance_record");
        const wsMid = await findSheetByMetadata(context, batchId, "gradesheet_record");

        if (!wsAtn || !wsMid) {
            console.error("Required sheets for Instructor Sync not found.");
            return;
        }

   
        const subjectRange = wsAtn.getRange("F10");
        const headerRowsRange = wsAtn.getRange("H12:IW13");

        subjectRange.load("values");
        headerRowsRange.load("values");

        await context.sync(); // One trip to Excel for all metadata[cite: 1, 2]

        const subjectNo = subjectRange.values[0][0];
        const attendanceData = headerRowsRange.values; 
        
        const changes = [];
        const dateRow = attendanceData[0];
        const flagRow = attendanceData[1];

        // 4. Process the columns in memory (High Performance)[cite: 1, 2]
for (let i = 0; i < dateRow.length; i++) {
    let tsRaw = dateRow[i];
    
    if (tsRaw && !isNaN(tsRaw)) {
        // Convert Excel OA Date to YYYY-MM-DD
        let dt = new Date((tsRaw - 25569) * 86400 * 1000);
        let timestamp = dt.toISOString().split('T')[0];

        let valObj = flagRow[i];
        
        // 1. SKIP blanks, nulls, and zeros completely
        if (valObj === null || valObj === "" || valObj === undefined || valObj === 0) {
            continue; // This skips to the next column without adding anything to 'changes'
        }

        // 2. NEW LOGIC: Less than zero = "delete", Greater than zero = "insert"
        let action = (Number(valObj) < 0) ? "delete" : "insert";

        changes.push({
            idnumber: batchId,
            instructorid: instructorId,
            subjectno: subjectNo,
            timestamp: timestamp,
            action: action
        });
    }
}
//console.log(changes)
        // 5. Submit Batch with Authorization
        if (changes.length > 0) {
            const response = await fetch(`https://${rawAddress}/submit-batchattn`, {
                method: 'POST',
                headers: { 
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${token}` // Use token to avoid 403[cite: 2]
                },
                body: JSON.stringify(changes)
            });

            if (response.ok) {
                console.log("✅ Instructor attendance headers synced.");
            } else if (response.status === 401 || response.status === 403) {
                // Handle expired token by showing overlay[cite: 2]
                showAuthOverlay(async () => await syncInstructorAttendance(context, batchId));
            } else {
                console.error("Instructor sync failed: " + response.status);
            }
        }

    } catch (error) {
        console.error("Error in syncInstructorAttendance: ", error);
    }
}
async function retrieveAttendanceScans(context, batchId) {
    const status = document.getElementById("status-message");
    
    try {
        const rawAddress = await getSettingValue(2);
        const token = await getSettingValue(6);
        const instructorId = await getSettingValue(3);

        // 1. Locate the correct Attendance sheet via Custom Properties
        const wsAtn = await findSheetByMetadata(context, batchId, "attendance_record");
        if (!wsAtn) {
            console.error("Attendance sheet not found for batch:", batchId);
            return;
        }

        // 2. Load the metadata required for the API call
        const startDateRange = wsAtn.getRange("E8");
        const endDateRange = wsAtn.getRange("E9");
        const subjectRange = wsAtn.getRange("F10");
        const totalTraineesRange = wsAtn.getRange("B13");

        startDateRange.load("values");
        endDateRange.load("values");
        subjectRange.load("values");
        totalTraineesRange.load("values");
        
        await context.sync(); // First sync to get parameters

        const subjectNo = subjectRange.values[0][0];
        const totalTrainees = parseInt(totalTraineesRange.values[0][0]);
        const startSerial = startDateRange.values[0][0];
        const endSerial = endDateRange.values[0][0];
        //console.log(endSerial);
        if (isNaN(totalTrainees) || totalTrainees <= 0) return;

        // Helper: Convert Excel Serial Date to YYYY-MM-DD string
        const formatExcelDate = (serial) => {
            const dt = new Date((serial - 25569) * 86400 * 1000);
            return dt.toISOString().split('T')[0];
        };

        const startDateStr = formatExcelDate(startSerial);
        const endDateStr = formatExcelDate(endSerial);

        status.innerText = "Retrieving records from scanner...";

        // 3. Call the Node.js API
        const url = `https://${rawAddress}/attendance/fetch?instructorid=${instructorId}&subjectno=${subjectNo}&startdate=${startDateStr}&enddate=${endDateStr}`;
        const response = await fetch(url, {
            headers: { 'Authorization': `Bearer ${token}` }
        });

        if (!response.ok) {
            if (response.status === 401 || response.status === 403) {
                // Token expired, show login overlay
                showAuthOverlay(async () => await retrieveAttendanceScans(context, batchId));
            } else {
                status.innerText = "❌ Failed to fetch records.";
                console.error("API Error:", response.status);
            }
            return;
        }

        const records = await response.json();
        //console.log(records);
        // 4. Map the Excel Structure in Memory
        const lastRow = 15 + totalTrainees - 1;
        const traineeRange = wsAtn.getRange(`A15:A${lastRow}`);
        const dateRange = wsAtn.getRange("H12:IW12");

        traineeRange.load("values");
        dateRange.load("values");
        await context.sync(); // Second sync to get map keys

        const traineeVals = traineeRange.values; 
        const dateVals = dateRange.values[0];    

        // Create fast lookup dictionaries (Objects in JS)
        const traineeRowMap = {};
        traineeVals.forEach((row, idx) => {
            if (row[0]) traineeRowMap[String(row[0])] = idx; 
        });

        const dateColMap = {};
        dateVals.forEach((serial, idx) => {
            if (serial && !isNaN(serial)) {
                dateColMap[formatExcelDate(serial)] = idx; 
            }
        });

        // 5. Prepare Blank 2D Arrays (Exactly matching the Excel dimensions)
        const rowCount = totalTrainees;
        const colCount = dateVals.length; // Let Excel tell us exactly how many columns H to HO is!
        
        // Fill arrays with empty strings to wipe out old data automatically
        const dataArr = Array.from({ length: rowCount }, () => Array(colCount).fill(""));
        const markRow13Arr = [Array(colCount).fill("")];

        // 6. Populate the 2D Arrays with JSON data
        records.forEach(rec => {
            const tid = String(rec.idnumber);
            const d = String(rec.date);
            const rid = rec.recordid;

            if (traineeRowMap.hasOwnProperty(tid) && dateColMap.hasOwnProperty(d)) {
                const r = traineeRowMap[tid];
                const c = dateColMap[d];
                
                // Write record ID into the main grid
                dataArr[r][c] = rid;

                // Mark Row 13 flag if valid
                if (!isNaN(rid) && parseInt(rid) > 1) {
                    markRow13Arr[0][c] = 1;
                }
            }
        });

        // 7. Write everything back to Excel in one massive chunk
        wsAtn.getRange(`H13:IW13`).values = markRow13Arr;
        wsAtn.getRange(`H15:IW${lastRow}`).values = dataArr;

        const startCleanupRow = lastRow + 1;

        if (startCleanupRow <= 48) {
            const cleanupRange = wsAtn.getRange(`A${startCleanupRow}:IW48`);
            cleanupRange.clear(Excel.ClearApplyTo.contents); 
        }
        await context.sync(); // Final Sync executes the writes

        status.innerText = "✅ External scans retrieved and synced.";
        status.style.color = "green";

    } catch (error) {
        console.error("Error retrieving scans:", error);
        status.innerText = "❌ Error processing scans.";
        status.style.color = "red";
    }
}
function initScheduleMonitor() {
    // 1. Run the check immediately so the card appears on load
    checkCurrentSchedule();

    // 2. Set interval for subsequent checks (every 30 seconds)
    const timerId = setInterval(checkCurrentSchedule, 30000);

    window.onbeforeunload = () => clearInterval(timerId);
}

async function startScheduleTimer() {
    // Run the check every 30 seconds
    const timerId = setInterval(async () => {
        try {
            await Excel.run(async (context) => {
                const table = context.workbook.tables.getItem("ScheduleTab");
                
                // Load the specific columns we need by their header names
                const columns = table.columns.load("name, values");
                await context.sync();

                // Convert table columns into an array of objects for easy filtering
                const scheduleData = transformTableToObject(columns.items);
                
                const now = new Date();
                const days = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"];
                const currentDay = days[now.getDay()];
                const currentMins = (now.getHours() * 60) + now.getMinutes();

                // Find the match based on day and time range
                const activeSched = scheduleData.find(row => {
                    const dayMatch = row.schedulecode.substring(0, 3).toUpperCase() === currentDay;
                    const startMins = timeToMins(row.timein);
                    const endMins = timeToMins(row.timeout);
                    
                    return dayMatch && currentMins >= startMins && currentMins <= endMins;
                });

                if (activeSched) {
    updateUI(activeSched);
} else {
    // Pass null to trigger the "pale gray" empty state
    updateUI(null); 
}
            });
        } catch (error) {
            console.error("Timer Error:", error);
        }
    }, 30000);

    // Stop timer if the pane is closed/reloaded
    window.onbeforeunload = () => clearInterval(timerId);
}

async function checkCurrentSchedule() {
    try {
        await Excel.run(async (context) => {
            const table = context.workbook.tables.getItem("ScheduleTab");
            const columns = table.columns.load("name, values");
            await context.sync();

            const scheduleData = transformTableToObject(columns.items);
            const now = new Date();
            const days = ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"];
            const currentDay = days[now.getDay()];
            const currentMins = (now.getHours() * 60) + now.getMinutes();

            const activeSched = scheduleData.find(row => {
                const dayMatch = row.schedulecode.substring(0, 3).toUpperCase() === currentDay;
                const startMins = timeToMins(row.timein);
                const endMins = timeToMins(row.timeout);
                return dayMatch && currentMins >= startMins && currentMins <= endMins;
            });

            // --- THE BLOCK YOU ASKED ABOUT ---
            if (activeSched) {
                updateUI(activeSched);
            } else {
                // Pass null to trigger the "pale gray" empty state
                updateUI(null);
            }
            // ---------------------------------
        });
    } catch (error) {
        console.error("Schedule Check Error:", error);
    }
}

/** 
 * Helper: Converts Excel column arrays into a list of JS Objects
 */
function transformTableToObject(columnItems) {
    const rowCount = columnItems[0].values.length;
    let data = [];

    // Start from index 1 to skip the header row itself
    for (let i = 1; i < rowCount; i++) {
        let obj = {};
        columnItems.forEach(col => {
            obj[col.name] = col.values[i][0];
        });
        data.push(obj);
    }
    return data;
}

/** 
 * Helper: Standardizes time strings to total minutes
 */
function timeToMins(timeVal) {
    if (!timeVal) return -1;
    // Excel might provide time as a decimal (0.677) or a string "16:15:00"
    if (typeof timeVal === 'number') {
        return Math.round(timeVal * 1440);
    }
    const parts = timeVal.split(':');
    return (parseInt(parts[0]) * 60) + parseInt(parts[1]);
}

function updateUI(sched) {
    const container = document.getElementById("schedule-container");
    
    if (!sched) {
        container.innerHTML = `
            <div class="schedule-card empty" style="background-color: #f3f2f1; padding: 12px; border: 1px dashed #ccc; text-align: center;">
                <span class="ms-fontSize-xs" style="color: #605e5c;">No active class at this time.</span>
            </div>`;
        return;
    }

    const accentColor = sched.color ? `#${sched.color}` : '#0078d4';

    // Store technical IDs in data attributes
    container.innerHTML = `
        <div class="schedule-card" style="border-left: 6px solid ${accentColor}; background: white; padding: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
            <div style="font-weight: 600; font-size: 15px;">${sched.subjectcode}</div>
            <div style="font-size: 13px; color: #323130;">${sched.subjectTitle}</div>
            <div style="font-size: 12px; color: #605e5c;">${sched.batchname}</div>
            <div style="font-size: 12px; color: #605e5c; margin-top: 4px;">
                <i class="ms-Icon ms-Icon--PoI"></i> ${sched.roomcode} - ${sched.roomdesc}
            </div>
            <button id="card-check-attendance" 
                class="ms-Button ms-Button--primary" 
                style="width: 100%; margin-top: 10px; background-color: #0078d4; color: white; border: none; height: 32px;"
                data-subjectno="${sched.subjectno}" 
                data-batchid="${sched.batchid}" 
                data-instructorid="${sched.instructorid}">
                <span class="ms-Button-label">Live Check Attendance</span>
            </button>
        </div>
    `;

document.getElementById("card-check-attendance").onclick = function() {
    const subjectNo = this.getAttribute("data-subjectno");
    const batchId = this.getAttribute("data-batchid");
    const instructorId = this.getAttribute("data-instructorid");
    // Extract batch name from the UI element or data attribute
    const batchName = sched.batchname; 
    
    showTraineeDrawer(instructorId, batchId, subjectNo, null, batchName);
};
}

async function showTraineeDrawer(instructorId, batchId, subjectNo, targetDate = null, batchName = "Select All") {
    const overlay = document.getElementById('trainee-overlay');
    const panel = document.getElementById('trainee-panel');
    const container = document.getElementById('trainee-list-container');
    const closeBtn = document.getElementById('close-trainee-panel');
    const refreshBtn = document.getElementById('confirm-trainee-selection');

    // 1. Hide Panel Logic
    const hideTraineePanel = () => {
        panel.classList.remove('show');
        setTimeout(() => { overlay.style.display = 'none'; }, 300);
    };

    closeBtn.onclick = hideTraineePanel;
    overlay.onclick = (e) => { if (e.target === overlay) hideTraineePanel(); };

// 2. The Internal Re-query Function (Fetch Logic)
    const loadData = async () => {
        container.innerHTML = '<div class="status neutral">Connecting to live records...</div>';
        const date = targetDate || new Date().toISOString().split('T')[0];
        const rawAddress = await getSettingValue(2);
        const token = await getSettingValue(6);

        try {
            // Define both URLs
            const traineesUrl = `https://${rawAddress}/fl_get_attendance_records?instructorid=${instructorId}&batchid=${batchId}&subjectno=${subjectNo}&date=${date}`;
            const classStatusUrl = `https://${rawAddress}/attendance/check-class?instructorid=${instructorId}&batchid=${batchId}&subjectno=${subjectNo}&date=${date}`;
            
            const fetchOptions = { headers: { 'Authorization': `Bearer ${token}` } };

            // Fetch BOTH at the same time for maximum speed
            const [response, statusResponse] = await Promise.all([
                fetch(traineesUrl, fetchOptions),
                fetch(classStatusUrl, fetchOptions)
            ]);

            if (response.status === 401 || response.status === 403) {
                showAuthOverlay(async () => await showTraineeDrawer(instructorId, batchId, subjectNo, targetDate, batchName));
                return;
            }

            if (!response.ok || !statusResponse.ok) throw new Error("Server error");
            
            const records = await response.json();
            const classStatus = await statusResponse.json();

            if (records.length === 0) {
                container.innerHTML = '<div class="status neutral">No trainees found for this class.</div>';
                return;
            }

            // Determine if the master checkbox should be checked
            const masterCheckedAttr = classStatus.isStarted ? 'checked' : '';

            // Render Header Checkbox + List
            container.innerHTML = `
                <div class="trainee-item" style="border-bottom: 2px solid #edebe9; margin-bottom: 10px; background: #f8f8f8;">
                    <input type="checkbox" id="batch-master-toggle" ${masterCheckedAttr}>
                    <label for="batch-master-toggle" style="font-weight: 700; font-size:16px;">
                        ${batchName.replace('Batch ', '')}
                    </label>
                </div>
                <div id="trainee-items-list">
                    ${records.map(rec => `
                        <div class="trainee-item">
                            <input type="checkbox" class="trainee-check" id="tr-${rec.traineesid}" value="${rec.traineesid}" ${rec.recordid ? 'checked' : ''}>
                            <label for="tr-${rec.traineesid}">
                                ${rec.trainee}
                                ${rec.recordid ? '<span style="color:green; font-size:10px; margin-left:8px;">(Present)</span>' : ''}
                            </label>
                        </div>
                    `).join('')}
                </div>
            `;

            // ==========================================
            // 👉 ATTACH EVENT LISTENERS HERE
            // ==========================================

            // 1. Master Toggle Logic
            const masterToggle = document.getElementById('batch-master-toggle');
            if (masterToggle) {
                masterToggle.onchange = async () => {
                    const action = masterToggle.checked ? 'add' : 'remove';
                    
                    if (action === 'remove') {
                        const confirmed = await showConfirmDialog("Clear Class Attendance?", "Unchecking this will delete ALL attendance records for this class session. Continue?");
                        if (!confirmed) {
                            masterToggle.checked = true;
                            return;
                        }
                    }

                    const result = await toggleAttendanceAPI({
                        instructorid: instructorId,
                        batchid: batchId,
                        subjectno: subjectNo,
                        targetid: batchId,
                        action: action,
                        isBulk: true
                    });

                    if (result && action === 'remove') loadData(); 
                };
            }

// 2. Individual Trainee Toggle Logic
            container.querySelectorAll('.trainee-check').forEach(cb => {
                cb.onchange = async () => {
                    const action = cb.checked ? 'add' : 'remove';

                    // --- NEW AUTO-START CLASS LOGIC ---
                    // If a trainee is checked, but the Master Batch isn't checked yet
                    if (action === 'add' && masterToggle && !masterToggle.checked) {
                        
                        masterToggle.checked = true; // 1. Visually check the Master UI
                        
                        // 2. Automatically create the Class/Batch record in the database
                        await toggleAttendanceAPI({
                            instructorid: instructorId,
                            batchid: batchId,
                            subjectno: subjectNo,
                            targetid: batchId, // Sending BatchID creates the class header
                            action: 'add',
                            isBulk: false
                        });
                        console.log("Auto-started class attendance based on trainee check.");
                    }
                    // ----------------------------------

                    // 3. Process the actual Trainee record
                    const success = await toggleAttendanceAPI({
                        instructorid: instructorId,
                        batchid: batchId,
                        subjectno: subjectNo,
                        targetid: cb.value,
                        action: action,
                        isBulk: false
                    });

                    if (!success) {
                        cb.checked = !cb.checked; // Revert trainee UI on failure
                    } else if (action === 'add') {
                        // Optional UX touch: Add the (Present) text immediately without reloading
                        const label = cb.nextElementSibling;
                        if (!label.innerHTML.includes('(Present)')) {
                            label.innerHTML += ' <span style="color:green; font-size:10px; margin-left:8px;">(Present)</span>';
                        }
                    } else if (action === 'remove') {
                        // Optional UX touch: Remove the (Present) text immediately without reloading
                        const label = cb.nextElementSibling;
                        const presentSpan = label.querySelector('span');
                        if (presentSpan) presentSpan.remove();
                    }
                };
            });
        } catch (error) {
            container.innerHTML = `<div class="status neutral" style="color: #a4262c; background: #fde7e9; border: 1px solid #a4262c;">
                <strong>Connection Problem</strong><br>You are currently offline.</div>`;
        }
    };


    overlay.style.display = 'flex';
    setTimeout(() => panel.classList.add('show'), 50);
    
    // Set Footer Button to Refresh
    refreshBtn.querySelector('.ms-Button-label').innerText = "Refresh List";
    refreshBtn.onclick = loadData;

    await loadData();
}

async function toggleAttendanceAPI(payload) {
    const rawAddress = await getSettingValue(2);
    const token = await getSettingValue(6);
    try {
        const response = await fetch(`https://${rawAddress}/attendance/toggle-live`, {
            method: 'POST',
            headers: { 
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}`
            },
            body: JSON.stringify(payload)
        });
        return response.ok;
    } catch (e) { return false; }
}