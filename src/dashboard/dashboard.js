// @ts-check
/* global document, Office, Excel, localStorage, window */

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initializeDashboard();
        
        // Target the elements
    const submitBtn = document.getElementById('submit-gradesheet');
    const authOverlay = document.getElementById('auth-overlay');
    const authPanel = document.getElementById('auth-panel');
    const closeBtn = document.getElementById('close-panel');
    const refreshBtn = document.getElementById("refresh-batch-view");
    const dropdown = document.getElementById("class-dropdown");

        if (refreshBtn) {
            refreshBtn.onclick = async () => {
                const selectedValue = dropdown.value;
                if (selectedValue && selectedValue !== "loading") {
                    const status = document.getElementById("status-message");
                    status.innerText = "Refreshing view...";
                    // Manually trigger the selection logic
                    await handleBatchSelection(selectedValue);
                    status.innerText = "View Refreshed";
                    setTimeout(() => { status.innerText = ""; }, 2000);
                }
            };
        }
        
        
    if (submitBtn) {
        // --- 1. SHOW PANEL ---
        submitBtn.onclick = function() {
            authOverlay.style.setProperty('display', 'flex', 'important');
            setTimeout(() => {
                authPanel.classList.add('show');
            }, 50);
        };

        // --- 2. HIDE PANEL (Close Button) ---
        if (closeBtn) {
            closeBtn.onclick = function() {
                hidePanel();
            };
        }

        // --- 3. HIDE PANEL (Clicking the dark background) ---
        authOverlay.onclick = function(e) {
            // Only hide if the user clicked the dark part, not the white box
            if (e.target === authOverlay) {
                hidePanel();
            }
        };
    }

    function hidePanel() {
        authPanel.classList.remove('show');
        // Wait for the slide animation (0.3s) before hiding the background
        setTimeout(() => {
            authOverlay.style.display = 'none';
        }, 300);
    }
        const disconnectBtn = document.getElementById("disconnect-button");
        if (disconnectBtn) disconnectBtn.onclick = handleDisconnect;
    }

    const syncBtn = document.getElementById('sync-attendance');
if (syncBtn) {
    syncBtn.onclick = refreshCalendar;
}
    
    const syncUpdatesBtn = document.getElementById("sync-classupdates");

if (syncUpdatesBtn) {
    syncUpdatesBtn.onclick = async () => {
        const status = document.getElementById("status-message");
        const rawAddress = localStorage.getItem("registrar_url");
        const baseUrl = `https://${rawAddress}`; // Remember your demo masking logic!

        try {
            // Re-use the exact same logic from the login flow
            await performFullSync(setProgress, status, baseUrl);
            
            // Re-hide sheets as needed
            await hideAllBatchSheets(); 
            status.innerText = "Updates synchronized successfully.";
        } catch (error) {
            status.innerText = "❌ Sync Failed: " + error.message;
            status.style.color = "red";
        }
    };
}
});

async function handleSubmitGrades() {
    const status = document.getElementById("status-message");
    status.innerText = "Processing gradesheet submission...";
    // ... your logic for reading Excel data and sending to API ...
}

async function initializeDashboard() {
    const nameDisplay = document.getElementById("instructor-name");
    const dropdown = document.getElementById("class-dropdown");

    const currentUrl = localStorage.getItem("registrar_url");
    if (currentUrl === "render-demoaddin-api.onrender.com") {
        document.getElementById("demo-strip").style.display = "block";
    }
    dropdown.onchange = () => handleBatchSelection(dropdown.value);
    
    const instructorName = localStorage.getItem("instructor_name");
    const instructorId = localStorage.getItem("user_id");

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
        // 1. Clear session from localStorage so the guardrails don't auto-skip login
        localStorage.removeItem("access_token");
        localStorage.removeItem("instructor_name");
        localStorage.removeItem("user_id");
        //localStorage.removeItem("username");

        // 2. Clear values in the veryHidden Settings sheet
        await Excel.run(async (context) => {
            const settingsSheet = context.workbook.worksheets.getItem("Settings");
            
            // We only clear B3:B5 (UserID, Username, InstName) 
            // We keep B2 (Server URL) so the user doesn't have to re-type the server
            //settingsSheet.getRange("B3").clear();
            settingsSheet.getRange("B5").clear();

            // Save the changes to the file
            //context.workbook.save();
            await context.sync();
        });

        // 3. Redirect back to login page
        window.location.href = "login.html";

    } catch (error) {
        console.error("Disconnect Error: ", error);
        // Fallback: if Excel fails, still redirect so the user isn't stuck
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
            context.workbook.application.suspendScreenUpdatingUntilNextSync();
            // Note: Use 'suspendScreenUpdatingUntilNextSync' only if your version supports it.
            // If it errors, just remove this line.
            const worksheets = context.workbook.worksheets;
            
            // 1. Load sheets
            worksheets.load("items/name");
            await context.sync();

            for (let sheet of worksheets.items) {
                // Only look at sheets with underscores (Attendance_211, etc.)
                if (sheet.name.includes("_")) {
                    const props = sheet.customProperties;
                    const batchProp = props.getItemOrNullObject("batchid");
                    
                    batchProp.load("value");
                    await context.sync();

                    // 2. Check if the metadata matches the selected ID
                    if (!batchProp.isNullObject && String(batchProp.value) === String(selectedBatchId)) {
                        sheet.visibility = Excel.SheetVisibility.visible;
                    } else {
                        sheet.visibility = Excel.SheetVisibility.hidden;
                    }
                }
            }
            await context.sync();
        });

        // 3. After toggling visibility, teleport the user to the Attendance sheet
        // We pass the selectedBatchId here!
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