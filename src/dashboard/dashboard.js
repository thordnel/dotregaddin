// @ts-check
/* global document, Office, Excel, localStorage, window */
let authCallback = null;

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
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
        if (syncBtn) syncBtn.onclick = refreshCalendar;
    }
});

// --- HELPER FUNCTIONS (Outside Office.onReady for access) ---

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
    console.log("StartSerial:", startSerial); 

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

    console.log(`Success! Start Column Index: ${startColIdx}`);
            
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


