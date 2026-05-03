// src/login/login.js
/* global document, Excel, Office */



Office.onReady(async (info) => {
    if (info.host === Office.HostType.Excel) {
        const connectBtn = document.getElementById("connect-button");
        const saveBtn = document.getElementById("save-settings-button");
        const loginBtn = document.getElementById("login-button"); // Select the button
        const usernameInput = document.getElementById("username");
        
        if (connectBtn) connectBtn.onclick = toggleConnection;
        if (saveBtn) saveBtn.onclick = saveServerToExcel;
        
        // Attach the handleLogin function to the click event
        if (loginBtn) loginBtn.onclick = handleLogin; 
        if (usernameInput) {
            usernameInput.ondblclick = clearRegistrationSettings;
        }
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
        let rawValue = urlInput.value.trim();
        if (rawValue.toLowerCase() === "demo") {
            rawValue = "render-demoaddin-api.onrender.com";
        }
        // 1. Strict Validation: Fail if ANY protocol is detected
        if (rawValue.includes("://")) {
            status.innerText = "❌ Error: Use address only (no https://).";
            status.style.color = "red";
            return;
        }

        if (!rawValue) {
            status.innerText = "Please enter an address";
            return;
        }

        // 2. Prepend https:// for the connection test only
        const baseUrl = `https://${rawValue}`;
        if (urlInput.value.trim() === "demo") {
            status.innerText = "Connecting to the Demo server... ";
            status.style.color = "#0078d4";
        } else {
            status.innerText = "Connecting to server...";
            status.style.color = "#0078d4";
        }
        
        

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
                


                if (urlInput.value.trim() === "demo") {
                    status.innerText = "✅ Connected to Secure Server";
                    status.style.color = "green";
                } else {
                    status.innerText = "✅ Connected to Secure Server";
                    status.style.color = "green";
                }
            } else {
                throw new Error();
            }
        } catch (err) {
            status.innerText = "❌ Failed connection to this server.";
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
async function clearRegistrationSettings() {
    const status = document.getElementById("status-message");
    const usernameField = document.getElementById("username");

    try {
        await Excel.run(async (context) => {
            const settingsSheet = context.workbook.worksheets.getItemOrNullObject("Settings");
            await context.sync();

            if (!settingsSheet.isNullObject) {
                // Clear B3 (UserID), B4 (Username), and B5 (InstName)
                // We keep B2 (Server) so they don't have to re-type the URL
                const range = settingsSheet.getRange("B3:B5");
                range.clear();
                
                await context.sync();

                // Reset UI elements
                usernameField.value = "";
                usernameField.disabled = false;
                
                // Clear LocalStorage to prevent auto-resuming
                await clearWorkbookSetting(3)
                await clearWorkbookSetting(4)
                await clearWorkbookSetting(5)


                status.innerText = "Registration cleared. Username unblocked.";
                status.style.color = "#0078d4";
            }
        });
    } catch (error) {
        console.error("Error clearing settings:", error);
    }
}
async function saveServerToExcel() {
    let serverUrl = document.getElementById("server-url").value.trim() || "";
    
    // 1. Translate "demo" back to the real URL before saving
    if (serverUrl.toLowerCase() === "demo") {
        serverUrl = "render-demoaddin-api.onrender.com";
    }
    
    await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        let settingsSheet = sheets.getItemOrNullObject("Settings");
        await context.sync();

        if (settingsSheet.isNullObject) {
            settingsSheet = sheets.add("Settings");
            await context.sync();
        }

        // 2. Hide the sheet immediately if it was just created
        settingsSheet.visibility = Excel.SheetVisibility.veryHidden;

        // 3. Save the ACTUAL URL to B2
        const range = settingsSheet.getRange("A1:B2");
        range.values = [
            ["Setting", "Value"],
            ["server", serverUrl]
        ];

        await context.sync();
    }).catch(error => {
        // ALWAYS keep this to catch Excel-specific errors
        console.error("Save to Excel Error: ", error);
    });
}
async function loadSettingsFromExcel() {
    await Excel.run(async (context) => {
        const settingsSheet = context.workbook.worksheets.getItemOrNullObject("Settings");
        // Load Server (B2), UserID (B3), Username (B4), InstName (B5)
        const range = settingsSheet.getRange("A2:B5"); 
        settingsSheet.calculate();
        range.load("values");
        await context.sync();

        if (!settingsSheet.isNullObject) {
            const values = range.values;
            const savedServer   = values[0][1]; // B2
            const savedUserID   = values[1][1]; // B3
            const savedUsername = values[2][1]; // B4
            const savedInstName = values[3][1]; // B5

            if (!savedServer) return;

            // --- MASKING LOGIC ---
            if (savedServer === "render-demoaddin-api.onrender.com") {
                document.getElementById("server-url").value = "demo"; 
            } else {
                document.getElementById("server-url").value = savedServer;
            }

            // Always store the TRUE URL for background sync functions
            await setWorkbookSetting(2, savedServer); 

            
            // --- END MASKING LOGIC ---

            // --- MODIFIED LOGIC START ---
            // Only trigger the connection test if the session is INCOMPLETE.
            // If B2, B3, B4, and B5 all have data, we skip the ping and go to dashboard.
            const isSessionComplete = savedServer && savedUserID && savedUsername && savedInstName;

            if (isSessionComplete) {
                // SKIP toggleConnection() to avoid the "Testing connection..." ping
                const urlInput = document.getElementById("server-url");
                urlInput.disabled = true;
                const status = document.getElementById("status-message");
                status.innerText = `Welcome back, ${savedInstName}. Resuming session...`;
                status.style.color = "green";

                await setWorkbookSetting(2, savedServer); 
                await setWorkbookSetting(3, savedUserID); 
                await setWorkbookSetting(4, savedUsername); 
                await setWorkbookSetting(5, savedInstName); 


                setTimeout(() => {
                    window.location.href = "dashboard.html";
                }, 1000);
            } else {
                // Only test connection if the user isn't fully logged in yet
                await toggleConnection();

                // Pre-fill username if it exists but other fields are missing
                if (savedUsername) {
                    const usernameField = document.getElementById("username");
                    usernameField.value = savedUsername;
                    usernameField.disabled = true; 
                }
            }
            // --- MODIFIED LOGIC END ---
        }
    }).catch(() => {
        // Settings sheet doesn't exist, stay on login
    });
}
async function handleLogin() {
    const user = document.getElementById("username").value;
    const pass = document.getElementById("password").value;
    const status = document.getElementById("status-message");
    
    let rawAddress = document.getElementById("server-url").value.trim();
    if (rawAddress.toLowerCase() === "demo") {
        rawAddress = "render-demoaddin-api.onrender.com";
    }
    
    const baseUrl = `https://${rawAddress}`;

    if (!user || !pass) {
        status.innerText = "Please enter username and password.";
        return;
    }

    // 1. Validation: Check if this workbook is already registered to someone else
    let registeredUser = "";
    const isUserValid = await Excel.run(async (context) => {
        const settingsSheet = context.workbook.worksheets.getItemOrNullObject("Settings");
        settingsSheet.load("isNullObject");
        await context.sync();

        if (!settingsSheet.isNullObject) {
            const userRange = settingsSheet.getRange("B4");
            userRange.load("values");
            await context.sync();
            registeredUser = String(userRange.values[0][0]).trim();
  
            if (registeredUser && registeredUser !== "" && user !== registeredUser) {
                return false; 
            }
        }
        return true;
    });

    if (!isUserValid) {
        status.innerText = "Error: Username does not match the registered user.";
        status.style.color = "red";
        document.getElementById("username").value = registeredUser;
        return;
    }

    status.innerText = "Authenticating...";
    status.style.color = "#0078d4";

    try {
        const response = await fetch(`${baseUrl}/apilogin`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ username: user, password: pass })
        });

        if (response.ok) {
            const data = await response.json();
            const tablesExist = await checkCoreTablesExist();

            // 2. SUCCESS: Save everything to the Settings Sheet (B2:B6)
            // This is the "Passport" that makes multiple workbooks work safely!
            await Excel.run(async (context) => {
                let settingsSheet = context.workbook.worksheets.getItemOrNullObject("Settings");
                await context.sync();

                if (settingsSheet.isNullObject) {
                    settingsSheet = context.workbook.worksheets.add("Settings");
                }

                // Update Labels (A) and Values (B)
                // B2: Server, B3: UserID, B4: Username, B5: InstName, B6: Token
                settingsSheet.getRange("A2:B6").values = [
                    ["server", rawAddress],
                    ["userid", String(data.user_id)],
                    ["username", user],
                    ["instname", data.instructor],
                    ["token", data.access_token] 
                ];
                settingsSheet.calculate();

                settingsSheet.visibility = Excel.SheetVisibility.veryHidden;
                await context.sync();
            });

            // 3. Sync Logic
            if (!tablesExist) {
                status.innerText = "New Workbook detected. Starting full sync...";
                await performFullSync(setProgress, status, baseUrl);
            } else {
                status.innerText = "Welcome back! Opening dashboard...";
                setProgress(100);
            }

            // 4. Final Cleanup and Redirect
            setTimeout(() => { window.location.href = "dashboard.html"; }, 1500);

        } else {
            const errorData = await response.json();
            status.innerText = "Login failed: " + (errorData.message || "Invalid credentials");
            status.style.color = "red";
        }
    } catch (error) {
        status.innerText = "❌ Error: " + error.message;
        console.error(error);
    }
}

async function checkCoreTablesExist() {
    return await Excel.run(async (context) => {
        // We check for EnrollmentTab as the "anchor" table
        const enrollmentTable = context.workbook.tables.getItemOrNullObject("EnrollmentTab");
        enrollmentTable.load("isNullObject");
        await context.sync();
        
        // If the table exists, we assume the workbook is already set up
        return !enrollmentTable.isNullObject;
    });
}


