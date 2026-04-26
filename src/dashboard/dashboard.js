/* global document, Office, Excel, localStorage, window */

Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initializeDashboard();
        
        // Target the elements
        const submitBtn = document.getElementById('submit-gradesheet');
    const authOverlay = document.getElementById('auth-overlay');
    const authPanel = document.getElementById('auth-panel');
    const closeBtn = document.getElementById('close-panel');

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
    
});

async function handleSubmitGrades() {
    const status = document.getElementById("status-message");
    status.innerText = "Processing gradesheet submission...";
    // ... your logic for reading Excel data and sending to API ...
}

async function initializeDashboard() {
    const nameDisplay = document.getElementById("instructor-name");
    const dropdown = document.getElementById("class-dropdown");

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
            settingsSheet.getRange("B3").clear();
            settingsSheet.getRange("B5").clear();

            // Save the changes to the file
            context.workbook.save();
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