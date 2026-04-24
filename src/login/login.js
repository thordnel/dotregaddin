// src/login/login.js
/* global document, Excel, Office */

Office.onReady(async (info) => {
    if (info.host === Office.HostType.Excel) {
        const connectBtn = document.getElementById("connect-button");
        const saveBtn = document.getElementById("save-settings-button");

        connectBtn.onclick = toggleConnection;
        saveBtn.onclick = saveServerToExcel;

        // Auto-load server address from Excel on startup
        await loadSettingsFromExcel();
    }
});

let isConnected = false;

Office.onReady(async (info) => {
    if (info.host === Office.HostType.Excel) {
        const connectBtn = document.getElementById("connect-button");
        const saveBtn = document.getElementById("save-settings-button");

        if (connectBtn) connectBtn.onclick = toggleConnection;
        if (saveBtn) saveBtn.onclick = saveServerToExcel;

        await loadSettingsFromExcel();
    }
});

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
    const serverUrl = document.getElementById("server-url").value.trim();
    const status = document.getElementById("status-message");

    await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        let settingsSheet = sheets.getItemOrNullObject("Settings");
        await context.sync();

        if (settingsSheet.isNullObject) {
            settingsSheet = sheets.add("Settings");
            settingsSheet.visibility = Excel.SheetVisibility.hidden; // Keep it hidden
        }

        const range = settingsSheet.getRange("A1:B2");
        range.values = [
            ["Setting", "Value"],
            ["server", serverUrl]
        ];
        range.format.font.bold = true;

        await context.sync();
        status.innerText = "✅ Address saved to Workbook Settings.";
    });
}

async function loadSettingsFromExcel() {
    await Excel.run(async (context) => {
        const settingsSheet = context.workbook.worksheets.getItemOrNullObject("Settings");
        const range = settingsSheet.getRange("A2:B2");
        range.load("values");
        await context.sync();

        if (!settingsSheet.isNullObject && range.values[0][0] === "server") {
            const savedUrl = range.values[0][1];
            document.getElementById("server-url").value = savedUrl;
            
            // NEW: Automatically trigger the connection toggle
            // We use a small delay to ensure the UI has updated
            setTimeout(() => {
                toggleConnection();
            }, 100);
        }
    }).catch(() => {
        // Ignore if sheet doesn't exist
    });
}