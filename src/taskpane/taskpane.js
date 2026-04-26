/* global document, Excel, Office */

/* global document, Excel, Office */

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    const sideloadMsg = document.getElementById("sideload-msg");
    if (sideloadMsg) sideloadMsg.style.display = "none";

    const appBody = document.getElementById("app-body");
    const getStartedBtn = document.getElementById("get-started");

    // Run compatibility check
    const guard = await checkWorkbookCompatibility();

    // 1. Clean up any existing warnings first to prevent duplicates
    const existingWarning = document.getElementById("compatibility-warning");
    if (existingWarning) {
        existingWarning.remove();
    }

    if (!guard.isCompatible) {
      appBody.style.display = "flex";
      getStartedBtn.disabled = true;
      getStartedBtn.style.opacity = "0.5";

      // 2. Create the warning with a unique ID
      const warning = document.createElement("div");
      warning.id = "compatibility-warning"; // Unique ID for tracking
      warning.className = "ms-MessageBar ms-MessageBar--error";
      warning.style.marginTop = "20px";
      warning.innerHTML = `
        <div class="ms-MessageBar-content">
          <div class="ms-MessageBar-text">
            <b>⚠️ Warning:</b> This workbook is not a valid DOT.REG Class Record. 
            To protect your existing data, add-in features have been disabled.
          </div>
        </div>`;
      
      // Insert it before the button
      getStartedBtn.parentNode.insertBefore(warning, getStartedBtn);
      
    } else {
      // 3. File is compatible (Empty or DotReg)
      appBody.style.display = "flex";
      getStartedBtn.disabled = false;
      getStartedBtn.style.opacity = "1";
      
      getStartedBtn.onclick = () => {
        window.location.href = "login.html";
      };
    }
  }
});

async function checkWorkbookCompatibility() {
  return await Excel.run(async (context) => {
    const workbook = context.workbook;
    const sheets = workbook.worksheets;
    
    // 1. Check for existing DOT.REG metadata first
    sheets.load("items/name");
    let settingsSheet = sheets.getItemOrNullObject("Settings");
    await context.sync();

    if (!settingsSheet.isNullObject) {
      return { isCompatible: true, isDotRegFile: true };
    }

    // 2. Scan sheets for actual content
    let hasData = false;
    for (let sheet of sheets.items) {
      // Get only the range that actually contains values
      const usedRange = sheet.getUsedRangeOrNullObject(); // 'true' ignores formatting/empty cells
      usedRange.load("address, values");
      await context.sync();

      if (!usedRange.isNullObject) {
        // If the used range exists and isn't just a single empty cell at A1
        const cellValue = usedRange.values[0][0];
        if (usedRange.address !== "A1" || (cellValue !== "" && cellValue !== null)) {
          hasData = true;
          break; 
        }
      }
    }

    // A workbook is only compatible if it's brand new (empty) or already a DOT.REG file
    return { isCompatible: !hasData, isDotRegFile: false };
  });
}