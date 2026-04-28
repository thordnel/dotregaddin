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
            status.innerText = "❌ Connection Failed. Please verify that the address is correct. If the issue persists, contact your system administrator for assistance.";
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
                localStorage.removeItem("user_id");
                localStorage.removeItem("username");
                localStorage.removeItem("instructor_name");

                status.innerText = "Registration cleared. You can now enter a new username.";
                status.style.color = "blue";
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
            localStorage.setItem("registrar_url", savedServer); 
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

                localStorage.setItem("registrar_url", savedServer);
                localStorage.setItem("user_id", savedUserID);
                localStorage.setItem("username", savedUsername);
                localStorage.setItem("instructor_name", savedInstName);

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
    
    // Change 'const' to 'let' so you can reassign it
    let rawAddress = document.getElementById("server-url").value.trim();
    
    console.log(rawAddress);

    // Now this reassignment will work without the readonly error
    if (rawAddress.toLowerCase() === "demo") {
        rawAddress = "render-demoaddin-api.onrender.com";
    }
    
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
    const tablesExist = await checkCoreTablesExist();

    // 1. ALWAYS store session info in LocalStorage
    localStorage.setItem("access_token", data.access_token);
    localStorage.setItem("registrar_url", rawAddress);
    localStorage.setItem("instructor_name", data.instructor);
    localStorage.setItem("user_id", data.user_id);
    localStorage.setItem("username", user);

    // 2. ALWAYS Update Settings Sheet in Excel (Crucial fix)
    await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        let settingsSheet = sheets.getItemOrNullObject("Settings");
        await context.sync();

        if (settingsSheet.isNullObject) {
            settingsSheet = sheets.add("Settings");
        }

        const uId = data.user_id ? String(data.user_id) : "";
        const uName = user ? String(user) : "";
        const iName = data.instructor ? String(data.instructor) : "";

        // Update B3, B4, and B5
        settingsSheet.getRange("A3:B5").values = [
            ["userid", uId],
            ["username", uName],
            ["instname", iName]
        ];

        await context.sync();
    });

    // 3. Conditional Download Block
    if (!tablesExist) {
        // REPLACE ALL THOSE setProgress/refreshData lines with this:
        const baseUrl = `https://${rawAddress}`;
        await performFullSync(setProgress, status, baseUrl);
    } else {
        status.innerText = "Welcome back! Opening dashboard...";
        setProgress(100);
    }

            await Excel.run(async (context) => {
                const sheets = context.workbook.worksheets;
                const attendance = sheets.getItemOrNullObject("Attendance");
                const sheet1 = sheets.getItemOrNullObject("Sheet1");
                const notepad = sheets.getItemOrNullObject("Notepad");
                const settings = sheets.getItem("Settings");
                await context.sync();

                settings.visibility = Excel.SheetVisibility.veryHidden;
                //context.workbook.save();
                await context.sync();
            });
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
