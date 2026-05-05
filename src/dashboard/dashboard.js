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
        if (syncUpdatesBtn) {
            syncUpdatesBtn.onclick = async () => {
                const status = document.getElementById("status-message");
                const rawAddress = await getSettingValue(2);
                const baseUrl = `https://${rawAddress}`;

                const startSync = async () => {
                    try {
                        await performFullSync(setProgress, status, baseUrl);
                        await postSyncCleanup();
                    } catch (error) {
                        // If the token is missing OR expired (401/403)
                        if (error.message.includes("401") || error.message.includes("Missing")) {
                            status.innerText = "Authentication required...";
                            showAuthOverlay(async () => {
                                status.innerText = "Resuming sync...";
                                await startSync(); // Try again after password is entered
                            });
                        } else {
                            status.innerText = "❌ Sync Failed: " + error.message;
                            status.style.color = "red";
                        }
                    }
                };

                await startSync();
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

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(event.worksheetId);
        const props = sheet.customProperties;
        
        const tProp = props.getItemOrNullObject("sheetType");
        tProp.load("value");
        await context.sync();

        // Check if we are on a Gradesheet
        const isGradesheet = !tProp.isNullObject && tProp.value === "gradesheet_record";

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
    status.innerText = "Processing gradesheet submission...";
    // ... your logic for reading Excel data and sending to API ...
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
            status.innerText = "No changes to sync. Checking for updates...";
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