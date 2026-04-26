/* global document, Excel, Office */

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) {
    // 1. Always hide the sideload message once Office is ready
    const sideloadMsg = document.getElementById("sideload-msg");
    if (sideloadMsg) sideloadMsg.style.display = "none";

    const appBody = document.getElementById("app-body");
    const getStartedBtn = document.getElementById("get-started");

    // 2. Run the Workbook Guardrail
    const guard = await checkWorkbookCompatibility();

    if (!guard.isCompatible) {
      // CASE A: File is NOT empty AND not a DOT.REG file
      appBody.style.display = "flex";
      getStartedBtn.disabled = true;
      getStartedBtn.style.opacity = "0.5";

      // Create and display professional warning
      const warning = document.createElement("div");
      warning.className = "ms-MessageBar ms-MessageBar--error";
      warning.style.marginTop = "20px";
      warning.innerHTML = `
        <div class="ms-MessageBar-content">
          <div class="ms-MessageBar-text">
            <b>⚠️ Warning:</b> This workbook is not a valid DOT.REG Class Record. 
            To protect your existing data, add-in features have been disabled. 
            Please open the correct class record or start with a blank workbook.
          </div>
        </div>`;
      getStartedBtn.parentNode.insertBefore(warning, getStartedBtn);
    } else {
      // CASE B & C: File is empty OR is a valid DOT.REG file
      appBody.style.display = "flex";
      getStartedBtn.onclick = () => {
        window.location.href = "login.html";
      };
    }
  }
});

async function checkWorkbookCompatibility() {
  return await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    let settingsSheet = sheets.getItemOrNullObject("Settings");
    await context.sync();

    if (!settingsSheet.isNullObject) {
      return { isCompatible: true, isDotRegFile: true };
    }

    let isTotallyEmpty = true;
    for (let sheet of sheets.items) {
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("address, values");
      await context.sync();

      if (!usedRange.isNullObject) {
        if (usedRange.address !== "A1" || (usedRange.values[0][0] !== "" && usedRange.values[0][0] !== null)) {
          isTotallyEmpty = false;
          break;
        }
      }
    }
    return { isCompatible: isTotallyEmpty, isDotRegFile: false };
  });
}