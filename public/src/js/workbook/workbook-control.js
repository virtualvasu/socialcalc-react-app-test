/**
 * @fileoverview Workbook Control - Controls workbook actions (add/del/rename etc) and can appear at the
 * bottom of the screen. Currently appears at the top of the screen as proof of concept.
 * @author Ramu Ramamurthy
 */

/**
 * @namespace SocialCalc
 */
var SocialCalc;
if (!SocialCalc) {
    alert("Main SocialCalc code module needed");
    SocialCalc = {};
}

/**
 * Current workbook control object instance
 * @type {SocialCalc.WorkBookControl|null}
 */
SocialCalc.CurrentWorkbookControlObject = null;

/**
 * Test workbook save string
 * @type {string}
 */
SocialCalc.TestWorkBookSaveStr = "";

/**
 * WorkBook Control constructor - manages workbook operations and UI
 * @class
 * @param {Object} book - The workbook object
 * @param {string} divid - The div ID where the control will be rendered
 * @param {string} defaultsheetname - Default name for new sheets
 */
SocialCalc.WorkBookControl = function (book, divid, defaultsheetname) {
    this.workbook = book;
    this.div = divid;
    this.defaultsheetname = defaultsheetname;

    /** @type {Object.<string, HTMLElement>} Sheet button array */
    this.sheetButtonArr = {};

    /** @type {number} Sheet counter */
    this.sheetCnt = 0;

    /** @type {number} Number of sheets */
    this.numSheets = 0;

    /** @type {HTMLElement|null} Current active sheet button */
    this.currentSheetButton = null;

    /** @type {string} Rename dialog element ID */
    this.renameDialogId = "sheetRenameDialog";

    /** @type {string} Delete dialog element ID */
    this.deleteDialogId = "sheetDeleteDialog";

    /** @type {string} Hide dialog element ID */
    this.hideDialogId = "sheetHideDialog";

    /** @type {string} Unhide dialog element ID */
    this.unhideDialogId = "sheetUnhideDialog";

    /** @type {string} Sheets HTML template */
    this.sheetshtml = '<div id="fooBar" style="background-color:#80A9F3;display:none"></div>';

    // Commented out buttons HTML - preserved for potential future use
    /*
    this.buttonshtml = 
    '<form>'+
    '<div id="workbookControls" style="padding:6px;background-color:#80A9F3;">'+
    '<input type="button" value="add sheet" onclick="SocialCalc.WorkBookControlAddSheet(true)" class="smaller">'+
    '<input type="button" value="delete sheet" onclick="SocialCalc.WorkBookControlDelSheet()" class="smaller">'+
    '<input type="button" value="rename sheet" onclick="SocialCalc.WorkBookControlRenameSheet()" class="smaller">'+
    '<input type="button" value="save workbook" onclick="SocialCalc.WorkBookControlSaveSheet()" class="smaller">'+
    '<input type="button" value="new workbook" onclick="SocialCalc.WorkBookControlNewBook()" class="smaller">'+
    '<input type="button" value="load workbook" onclick="SocialCalc.WorkBookControlLoad()" class="smaller">'+
    '<input type="button" value="copy sheet" onclick="SocialCalc.WorkBookControlCopySheet()" class="smaller">'+
    '<input type="button" value="paste sheet" onclick="SocialCalc.WorkBookControlPasteSheet()" class="smaller">'+
    '</div>'+
    '</form>';
    */

    SocialCalc.CurrentWorkbookControlObject = this;
    this.sheetbar = new SocialCalc.SheetBar();
};

/**
 * Get current workbook control instance
 * @returns {SocialCalc.WorkBookControl} Current workbook control object
 */
SocialCalc.WorkBookControl.prototype.GetCurrentWorkBookControl = function () {
    return SocialCalc.GetCurrentWorkBookControl();
};

/**
 * Initialize the workbook control
 * @returns {void}
 */
SocialCalc.WorkBookControl.prototype.InitializeWorkBookControl = function () {
    return SocialCalc.InitializeWorkBookControl(this);
};

/**
 * Execute workbook control command
 * @param {Object} cmd - Command object containing cmdstr and cmdtype
 * @param {boolean} isremote - Whether the command is from remote source
 * @returns {void}
 */
SocialCalc.WorkBookControl.prototype.ExecuteWorkBookControlCommand = function (cmd, isremote) {
    return SocialCalc.ExecuteWorkBookControlCommand(this, cmd, isremote);
};

/**
 * Execute workbook control command (static method)
 * @param {SocialCalc.WorkBookControl} control - The workbook control instance
 * @param {Object} cmd - Command object with cmdstr, cmdtype, and optional sheetstr
 * @param {boolean} isremote - Whether command originates from remote source
 * @returns {void}
 */
SocialCalc.ExecuteWorkBookControlCommand = function (control, cmd, isremote) {
    // console.log("cmd ", cmd.cmdstr, cmd.cmdtype);

    // Commented out remote check - preserved for potential future use
    // if (!isremote) {
    //   return;
    // }

    if (cmd.cmdtype === "scmd") {
        // dispatch a sheet command
        control.workbook.WorkbookScheduleCommand(cmd, isremote);
        return;
    }

    if (cmd.cmdtype !== "wcmd") {
        return;
    }

    const parseobj = new SocialCalc.Parse(cmd.cmdstr);
    const cmd1 = parseobj.NextToken();

    switch (cmd1) {
        case "addsheet":
            SocialCalc.WorkBookControlAddSheetRemote(null);
            break;

        case "addsheetstr": {
            const sheetstr = cmd.sheetstr;
            SocialCalc.WorkBookControlAddSheetRemote(sheetstr);
            break;
        }

        case "delsheet": {
            const sheetid = parseobj.NextToken();
            SocialCalc.WorkBookControlDelSheetRemote(sheetid);
            break;
        }

        case "rensheet": {
            const sheetid = parseobj.NextToken();
            const oldname = parseobj.NextToken();
            const newname = parseobj.NextToken();
            SocialCalc.WorkBookControlRenameSheetRemote(sheetid, oldname, newname);
            break;
        }

        case "activatesheet": {
            const sheetid = parseobj.NextToken();
            SocialCalc.WorkBookControlActivateSheet(sheetid);
            break;
        }

        case "hidesheet": {
            const sheetid = parseobj.NextToken();
            // Implementation placeholder
            break;
        }

        case "unhidesheet": {
            const sheetid = parseobj.NextToken();
            // Implementation placeholder
            break;
        }
    }
};

/**
 * Get the current workbook control instance
 * @returns {SocialCalc.WorkBookControl} The current workbook control object
 */
SocialCalc.GetCurrentWorkBookControl = function () {
    return SocialCalc.CurrentWorkbookControlObject;
};

/**
 * Initialize workbook control by creating DOM elements and setting up default sheet
 * @param {SocialCalc.WorkBookControl} control - The workbook control instance to initialize
 * @returns {void}
 */
SocialCalc.InitializeWorkBookControl = function (control) {
    const element = document.createElement("div");
    element.innerHTML = control.sheetshtml;

    const containerElement = document.getElementById(control.div);
    containerElement.appendChild(element);

    // Commented out buttons HTML creation - preserved for potential future use
    // const element2 = document.createElement("div");
    // element2.innerHTML = control.buttonshtml;
    // containerElement.appendChild(element2);

    SocialCalc.WorkBookControlAddSheet(false); // Initialize default sheet
};

/**
 * Delete sheet from remote command
 * @param {string} sheetid - ID of the sheet to delete
 * @returns {void}
 */
SocialCalc.WorkBookControlDelSheetRemote = function (sheetid) {
    const control = SocialCalc.GetCurrentWorkBookControl();

    if (sheetid === control.currentSheetButton.id) {
        // The active sheet is being deleted
        SocialCalc.WorkBookControlDelSheet();
        return;
    }

    // Some non-active sheet is being deleted
    const containerElement = document.getElementById("fooBar");
    const deletedButton = document.getElementById(sheetid);

    const buttonId = deletedButton.id;
    const buttonName = deletedButton.value;
    delete control.sheetButtonArr[buttonId];

    containerElement.removeChild(deletedButton);

    const sheetbar = document.getElementById("SocialCalc-sheetbar-buttons");
    const sheetbarButton = document.getElementById(`sbsb-${buttonId}`);
    // TODO: unregister with mouse events etc
    sheetbar.removeChild(sheetbarButton);

    // Delete the sheet
    control.workbook.DeleteWorkBookSheet(buttonId, buttonName);
    control.numSheets = control.numSheets - 1;
};

/**
 * Delete current active sheet with confirmation dialog
 * @returns {void}
 */
SocialCalc.WorkBookControlDelSheet = function () {
    const control = SocialCalc.GetCurrentWorkBookControl();

    if (control.workbook.spreadsheet.editor.state !== "start") {
        // If in edit mode, return
        return;
    }

    if (control.numSheets === 1) {
        // Disallow deletion of last sheet
        const warningMessage =
            '<div style="padding:6px 0px 4px 6px;">' +
            '<span><b>A workbook must contain at least one worksheet</b></span><br/><br/>' +
            '<span>To delete the selected sheet, you must first insert a new sheet.</span><br/>' +
            '</div>' +
            '<div style="width:380px;text-align:right;padding:6px 0px 4px 6px;font-size:small;">' +
            '<input type="button" value="Ok" style="font-size:smaller;" onclick="SocialCalc.WorkBookControlDeleteSheetHide();"></div>';

        const dialogElement = SocialCalc.createDialog(control.deleteDialogId, warningMessage, control.workbook.spreadsheet.spreadsheetDiv);
        return;
    }

    // Check if dialog already exists
    const existingDialog = document.getElementById(control.deleteDialogId);
    if (existingDialog) return;

    const currentSheetName = control.currentSheetButton.value;
    const confirmationMessage =
        '<div style="padding:6px 0px 4px 6px;">' +
        '<span><b>The selected sheet will be permanently deleted.</b></span><br/>' +
        '<span><ul>' +
        '<li>To delete the selected sheet, click OK.</li>' +
        '<li>To cancel the deletion, click cancel.</li>' +
        '</ul></span>' +
        '</div>' +
        '<div style="width:380px;text-align:right;padding:6px 0px 4px 6px;font-size:small;">' +
        '<input type="button" value="Cancel" style="font-size:smaller;" onclick="SocialCalc.WorkBookControlDeleteSheetHide();">&nbsp;' +
        '<input type="button" value="OK" style="font-size:smaller;" onclick="SocialCalc.WorkBookControlDeleteSheetSubmit();"></div>';

    SocialCalc.createDialog(control.deleteDialogId, confirmationMessage, control.workbook.spreadsheet.spreadsheetDiv);
};

/**
 * Create a modal dialog element
 * @param {string} dialogId - ID for the dialog element
 * @param {string} content - HTML content for the dialog body
 * @param {HTMLElement} parentElement - Parent element to append dialog to
 * @returns {HTMLElement} The created dialog element
 */
SocialCalc.createDialog = function (dialogId, content, parentElement) {
    const dialogElement = document.createElement("div");
    dialogElement.id = dialogId;
    dialogElement.style.position = "absolute";

    const viewport = SocialCalc.GetViewportInfo();
    dialogElement.style.top = `${viewport.height / 3}px`;
    dialogElement.style.left = `${viewport.width / 3}px`;
    dialogElement.style.zIndex = "100";
    dialogElement.style.backgroundColor = "#FFF";
    dialogElement.style.border = "1px solid black";
    dialogElement.style.width = "400px";

    dialogElement.innerHTML =
        '<table cellspacing="0" cellpadding="0" style="border-bottom:1px solid black;"><tr>' +
        '<td style="font-size:10px;cursor:default;width:100%;background-color:#999;color:#FFF;">&nbsp;</td>' +
        '<td style="font-size:10px;cursor:default;color:#666;" onclick="SocialCalc.WorkBookControlDeleteSheetHide();">&nbsp;X&nbsp;</td>' +
        '</tr></table>' +
        `<div style="background-color:#DDD;">${content}</div>`;

    // Register drag functionality
    SocialCalc.DragRegister(
        dialogElement.firstChild.firstChild.firstChild.firstChild,
        true,
        true,
        {
            MouseDown: SocialCalc.DragFunctionStart,
            MouseMove: SocialCalc.DragFunctionPosition,
            MouseUp: SocialCalc.DragFunctionPosition,
            Disabled: null,
            positionobj: dialogElement
        }
    );

    parentElement.appendChild(dialogElement);
    return dialogElement;
};

/**
 * Hide the delete sheet confirmation dialog
 * @returns {void}
 */
SocialCalc.WorkBookControlDeleteSheetHide = function () {
    const control = SocialCalc.GetCurrentWorkBookControl();
    const dialogElement = document.getElementById(control.deleteDialogId);

    if (!dialogElement) return;

    dialogElement.innerHTML = "";
    SocialCalc.DragUnregister(dialogElement);
    SocialCalc.KeyboardFocus();

    if (dialogElement.parentNode) {
        dialogElement.parentNode.removeChild(dialogElement);
    }
};
/**
 * Submit sheet deletion after user confirmation
 * Removes the current active sheet, updates UI, broadcasts command, and activates another sheet
 * @returns {void}
 */
SocialCalc.WorkBookControlDeleteSheetSubmit = function () {
    const control = SocialCalc.GetCurrentWorkBookControl();
    SocialCalc.WorkBookControlDeleteSheetHide();

    const containerElement = document.getElementById("fooBar");
    const currentButton = document.getElementById(control.currentSheetButton.id);

    const sheetId = currentButton.id;
    const currentSheetName = control.currentSheetButton.value;
    delete control.sheetButtonArr[sheetId];

    containerElement.removeChild(currentButton);

    const sheetbar = document.getElementById("SocialCalc-sheetbar-buttons");
    const sheetbarButton = document.getElementById(`sbsb-${currentButton.id}`);
    // TODO: unregister with mouse events etc
    sheetbar.removeChild(sheetbarButton);

    control.currentSheetButton = null;
    // Delete the sheet from workbook
    control.workbook.DeleteWorkBookSheet(sheetId, currentSheetName);
    control.numSheets = control.numSheets - 1;

    const commandString = `delsheet ${sheetId}`;
    SocialCalc.Callbacks.broadcast('execute', {
        cmdtype: "wcmd",
        id: "0",
        cmdstr: commandString
    });

    // Reset current sheet - activate the first available sheet
    for (const sheetKey in control.sheetButtonArr) {
        if (sheetKey != null) {
            control.currentSheetButton = control.sheetButtonArr[sheetKey];
            control.currentSheetButton.setAttribute("style", "background-color:lightgreen");
            SocialCalc.SheetBarButtonActivate(control.currentSheetButton.id, true);
            break;
        }
    }

    if (control.currentSheetButton != null) {
        control.workbook.ActivateWorkBookSheet(control.currentSheetButton.id, null);
    }
};

/**
 * Hide current active sheet with confirmation dialog
 * Shows warning if only one sheet exists, otherwise shows confirmation dialog
 * @returns {void}
 */
SocialCalc.WorkBookControlHideSheet = function () {
    const control = SocialCalc.GetCurrentWorkBookControl();

    if (control.workbook.spreadsheet.editor.state !== "start") {
        // If in edit mode, return
        return;
    }

    if (control.numSheets === 1) {
        // Disallow hiding the last sheet
        const warningMessage =
            '<div style="padding:6px 0px 4px 6px;">' +
            '<span><b>A workbook must contain at least one worksheet</b></span><br/><br/>' +
            '<span>Before hiding the selected sheet, you must first insert a new sheet.</span><br/>' +
            '</div>' +
            '<div style="width:380px;text-align:right;padding:6px 0px 4px 6px;font-size:small;">' +
            '<input type="button" value="Ok" style="font-size:smaller;" onclick="SocialCalc.WorkBookControlHideSheetHide();"></div>';

        SocialCalc.createDialog(control.hideDialogId, warningMessage, control.workbook.spreadsheet.spreadsheetDiv);
        return;
    }

    // Check if dialog already exists
    const existingDialog = document.getElementById(control.hideDialogId);
    if (existingDialog) return;

    const currentSheetName = control.currentSheetButton.value;
    const confirmationMessage =
        '<div style="padding:6px 0px 4px 6px;">' +
        '<span><b>The selected sheet will be hidden.</b></span><br/>' +
        '<span><ul>' +
        '<li>To hide the selected sheet, click OK.</li>' +
        '<li>To cancel the hiding, click cancel.</li>' +
        '</ul></span>' +
        '</div>' +
        '<div style="width:380px;text-align:right;padding:6px 0px 4px 6px;font-size:small;">' +
        '<input type="button" value="Cancel" style="font-size:smaller;" onclick="SocialCalc.WorkBookControlHideSheetHide();">&nbsp;' +
        '<input type="button" value="OK" style="font-size:smaller;" onclick="SocialCalc.WorkBookControlHideSheetSubmit();"></div>';

    SocialCalc.createDialog(control.hideDialogId, confirmationMessage, control.workbook.spreadsheet.spreadsheetDiv);
};
/**
 * Hide the sheet hide/unhide dialog and clean up
 * @returns {void}
 */
SocialCalc.WorkBookControlHideSheetHide = function () {
    const control = SocialCalc.GetCurrentWorkBookControl();
    const spreadsheet = control.workbook.spreadsheet;

    const dialogElement = document.getElementById(control.hideDialogId);

    if (!dialogElement) return;

    dialogElement.innerHTML = "";
    SocialCalc.DragUnregister(dialogElement);
    SocialCalc.KeyboardFocus();

    if (dialogElement.parentNode) {
        dialogElement.parentNode.removeChild(dialogElement);
    }
};

/**
 * Submit sheet hiding after user confirmation
 * Hides the current active sheet, updates UI, broadcasts command, and activates another visible sheet
 * @returns {void}
 */
SocialCalc.WorkBookControlHideSheetSubmit = function () {
    const control = SocialCalc.GetCurrentWorkBookControl();
    SocialCalc.WorkBookControlHideSheetHide();

    const containerElement = document.getElementById("fooBar");
    const currentButton = document.getElementById(control.currentSheetButton.id);

    const sheetId = currentButton.id;
    const currentSheetName = control.currentSheetButton.value;

    const sheetbar = document.getElementById("SocialCalc-sheetbar-buttons");
    const sheetbarButton = document.getElementById(`sbsb-${currentButton.id}`);

    // TODO: unregister with mouse events etc
    SocialCalc.SheetBarButtonActivate(control.currentSheetButton.id, false);
    sheetbarButton.style.display = "none";
    control.currentSheetButton = null;

    control.numSheets = control.numSheets - 1;

    const commandString = `hidesheet ${sheetId}`;
    SocialCalc.Callbacks.broadcast('execute', {
        cmdtype: "wcmd",
        id: "0",
        cmdstr: commandString
    });

    // Reset current sheet - activate the first available visible sheet
    for (const sheetKey in control.sheetButtonArr) {
        if (sheetKey != null && document.getElementById(`sbsb-${sheetKey}`).style.display !== "none") {
            control.currentSheetButton = control.sheetButtonArr[sheetKey];
            break;
        }
    }

    if (control.currentSheetButton != null) {
        control.currentSheetButton.setAttribute("style", "background-color:lightgreen");
        SocialCalc.SheetBarButtonActivate(control.currentSheetButton.id, true);
        control.workbook.ActivateWorkBookSheet(control.currentSheetButton.id, null);
    }
};

/**
 * Display unhide sheet dialog with list of hidden sheets
 * Shows warning if no sheets are hidden, otherwise shows selection dialog
 * @returns {void}
 */
SocialCalc.WorkBookControlUnhideSheet = function () {
    const control = SocialCalc.GetCurrentWorkBookControl();

    if (control.workbook.spreadsheet.editor.state !== "start") {
        // If in edit mode, return
        return;
    }

    // Count hidden sheets
    let hiddenCount = 0;
    for (const sheetKey in control.sheetButtonArr) {
        if (document.getElementById(`sbsb-${sheetKey}`).style.display === "none") {
            hiddenCount++;
        }
    }

    if (hiddenCount === 0) {
        // No hidden sheets, show error message
        const noHiddenSheetsMessage =
            '<div style="padding:6px 0px 4px 6px;">' +
            '<span><b>There are no hidden worksheets.</b></span><br/><br/>' +
            '<span>Before unhiding any sheets, you must first hide a sheet.</span><br/>' +
            '</div>' +
            '<div style="width:380px;text-align:right;padding:6px 0px 4px 6px;font-size:small;">' +
            '<input type="button" value="Ok" style="font-size:smaller;" onclick="SocialCalc.WorkBookControlUnhideSheetHide();"></div>';

        SocialCalc.createDialog(control.unhideDialogId, noHiddenSheetsMessage, control.workbook.spreadsheet.spreadsheetDiv);
        return;
    }

    // Check if dialog already exists
    const existingDialog = document.getElementById(control.unhideDialogId);
    if (existingDialog) return;

    const currentSheetName = control.currentSheetButton.value;

    // Build dialog content with hidden sheets list
    let dialogContent =
        '<div style="padding:6px 0px 4px 6px;">' +
        '<span><b>The following sheets are hidden.</b></span><br/>' +
        '<form id="unhidesheetform"><ul>' +
        '<input type="hidden" name="unhidesheet" value=""/>';

    // Add radio buttons for each hidden sheet
    for (const sheetKey in control.sheetButtonArr) {
        if (document.getElementById(`sbsb-${sheetKey}`).style.display === "none") {
            const sheetName = control.sheetButtonArr[sheetKey].value;
            dialogContent +=
                `<input type="radio" value="${sheetKey}" ` +
                `onclick="document.getElementById('unhidesheetform').unhidesheet.value='${sheetKey}';"/> ` +
                `${sheetName}<br/>`;
        }
    }

    dialogContent +=
        '</ul></form>' +
        '<span><ul>' +
        '<li>To unhide the selected sheet, click OK.</li>' +
        '<li>To cancel the unhiding, click cancel.</li>' +
        '</ul></span>' +
        '</div>' +
        '<div style="width:380px;text-align:right;padding:6px 0px 4px 6px;font-size:small;">' +
        '<input type="button" value="Cancel" style="font-size:smaller;" onclick="SocialCalc.WorkBookControlUnhideSheetHide();">&nbsp;' +
        '<input type="button" value="OK" style="font-size:smaller;" ' +
        'onclick="SocialCalc.WorkBookControlUnhideSheetSubmit(document.getElementById(\'unhidesheetform\').unhidesheet.value);"></div>';

    SocialCalc.createDialog(control.unhideDialogId, dialogContent, control.workbook.spreadsheet.spreadsheetDiv);
};
/**
 * Hide the unhide sheet dialog and clean up
 * @returns {void}
 */
SocialCalc.WorkBookControlUnhideSheetHide = function () {
    const control = SocialCalc.GetCurrentWorkBookControl();
    const spreadsheet = control.workbook.spreadsheet;

    const dialogElement = document.getElementById(control.unhideDialogId);

    if (!dialogElement) return;

    dialogElement.innerHTML = "";
    SocialCalc.DragUnregister(dialogElement);
    SocialCalc.KeyboardFocus();

    if (dialogElement.parentNode) {
        dialogElement.parentNode.removeChild(dialogElement);
    }
};

/**
 * Submit unhide sheet operation for specified sheet
 * @param {string} name - ID of the sheet to unhide
 * @returns {void}
 */
SocialCalc.WorkBookControlUnhideSheetSubmit = function (name) {
    const control = SocialCalc.GetCurrentWorkBookControl();
    SocialCalc.WorkBookControlUnhideSheetHide();

    const currentButton = document.getElementById(control.currentSheetButton.id);
    const currentId = currentButton.id;
    const currentName = control.currentSheetButton.value;

    control.currentSheetButton.setAttribute("style", "");
    const previousSheetId = control.currentSheetButton.id;
    console.log(previousSheetId);
    SocialCalc.SheetBarButtonActivate(previousSheetId, false);

    const sheetbarButton = document.getElementById(`sbsb-${name}`);
    // Unhide the button
    sheetbarButton.style.display = "inline";
    control.currentSheetButton = null;

    control.numSheets = control.numSheets + 1;

    const commandString = `unhidesheet ${name}`;
    SocialCalc.Callbacks.broadcast('execute', {
        cmdtype: "wcmd",
        id: "0",
        cmdstr: commandString
    });

    // Reset current sheet - activate the first available visible sheet
    for (const sheetKey in control.sheetButtonArr) {
        if (sheetKey != null && document.getElementById(`sbsb-${sheetKey}`).style.display !== "none") {
            control.currentSheetButton = control.sheetButtonArr[sheetKey];
            break;
        }
    }

    if (control.currentSheetButton != null) {
        control.currentSheetButton.setAttribute("style", "background-color:lightgreen");
        SocialCalc.SheetBarButtonActivate(control.currentSheetButton.id, true);
        control.workbook.ActivateWorkBookSheet(control.currentSheetButton.id, null);
    }
};

/**
 * Create and add a sheet button to the workbook control
 * @param {string|null} sheetname - Name for the sheet (null for auto-generated)
 * @param {string|null} sheetid - ID for the sheet (null for auto-generated)
 * @returns {HTMLElement} The created button element
 */
SocialCalc.WorkBookControlAddSheetButton = function (sheetname, sheetid) {
    const control = SocialCalc.GetCurrentWorkBookControl();

    // Create button element dynamically
    const element = document.createElement("input");

    let sheetIdentifier = null;

    if (sheetid != null) {
        sheetIdentifier = sheetid;
    } else {
        sheetIdentifier = `sheet${(control.sheetCnt + 1).toString()}`;
        control.sheetCnt = control.sheetCnt + 1;
    }

    // Assign attributes to the element
    element.setAttribute("type", "button");
    element.setAttribute("value", sheetname || sheetIdentifier);
    element.setAttribute("id", sheetIdentifier);
    element.setAttribute("name", sheetIdentifier);

    const functionName = `SocialCalc.WorkBookControlActivateSheet('${sheetIdentifier}')`;
    element.setAttribute("onclick", functionName);

    control.sheetButtonArr[sheetIdentifier] = element;

    const containerElement = document.getElementById("fooBar");

    // Append the element to the page
    containerElement.appendChild(element);

    control.numSheets = control.numSheets + 1;

    const sheetBarButton = new SocialCalc.SheetBarSheetButton(
        `sbsb-${sheetIdentifier}`,
        sheetname || sheetIdentifier,
        document.getElementById("SocialCalc-sheetbar-buttons"),
        {
            // Commented styles preserved for potential future use
            // normalstyle: "border:1px solid #000;backgroundColor:#FFF;",
            // downstyle: "border:1px solid #000;backgroundColor:#CCC;",
            // hoverstyle: "border:1px solid #000;backgroundColor:#FFF;"
        },
        {
            MouseDown: function () { SocialCalc.SheetBarSheetButtonPress(sheetIdentifier); },
            Repeat: function () { },
            Disabled: function () { }
        }
    );

    return element;
};

/**
 * Add a new sheet to the workbook with optional name
 * @param {boolean} addworksheet - Whether to actually create the worksheet
 * @param {string|null} sheetname - Optional name for the new sheet
 * @returns {void}
 */
SocialCalc.WorkBookControlAddSheet = function (addworksheet, sheetname) {
    const control = SocialCalc.GetCurrentWorkBookControl();

    if (control.workbook.spreadsheet.editor.state !== "start") {
        // If in edit mode, return
        return;
    }

    // First add the button
    const element = SocialCalc.WorkBookControlAddSheetButton(sheetname);

    // Then change the highlight
    let previousSheetId = "sheet1";

    if (control.currentSheetButton != null) {
        control.currentSheetButton.setAttribute("style", "");
        previousSheetId = control.currentSheetButton.id;
        SocialCalc.SheetBarButtonActivate(previousSheetId, false);
    }

    element.setAttribute("style", "background-color:lightgreen");
    control.currentSheetButton = element;

    const newSheetId = element.id;
    SocialCalc.SheetBarButtonActivate(newSheetId, true);

    // Create the sheet
    if (addworksheet) {
        control.workbook.AddNewWorkBookSheet(newSheetId, previousSheetId, false);

        // Broadcast an add command
        const commandString = "addsheet";
        SocialCalc.Callbacks.broadcast('execute', {
            cmdtype: "wcmd",
            id: "0",
            cmdstr: commandString
        });
    }
};

/**
 * Add sheet from remote command without switching to it
 * @param {string|null} savestr - Saved sheet data string
 * @returns {void}
 */
SocialCalc.WorkBookControlAddSheetRemote = function (savestr) {
    const control = SocialCalc.GetCurrentWorkBookControl();

    // First add the button
    const element = SocialCalc.WorkBookControlAddSheetButton();

    // Add the sheet, don't switch to it
    control.workbook.AddNewWorkBookSheetNoSwitch(
        element.id,
        element.value,
        savestr
    );
};

/**
 * Activate specified sheet by name/ID
 * @param {string} name - ID/name of the sheet to activate
 * @returns {void}
 */
SocialCalc.WorkBookControlActivateSheet = function (name) {
    // alert("in activate sheet=" + name)

    const control = SocialCalc.GetCurrentWorkBookControl();

    const targetButton = document.getElementById(name);
    targetButton.setAttribute("style", "background-color:lightgreen;");
    SocialCalc.SheetBarButtonActivate(name, true);

    const previousSheetId = control.currentSheetButton.id;

    if (control.currentSheetButton.id !== targetButton.id) {
        control.currentSheetButton.setAttribute("style", "");
        SocialCalc.SheetBarButtonActivate(previousSheetId, false);
    }

    control.currentSheetButton = targetButton;
    control.workbook.ActivateWorkBookSheet(name, previousSheetId);
};

/**
 * HTTP request object for workbook control operations
 * @type {XMLHttpRequest|null}
 */
SocialCalc.WorkBookControlHttpRequest = null;

/**
 * Alert contents from HTTP request response
 * Handles the response from workbook control HTTP requests
 * @returns {void}
 */
SocialCalc.WorkBookControlAlertContents = function () {
    let loadedString = "";
    const httpRequest = SocialCalc.WorkBookControlHttpRequest;

    if (httpRequest.readyState === 4) {
        // addmsg("received:" + httpRequest.responseText.length + " chars");
        try {
            if (httpRequest.status === 200) {
                loadedString = httpRequest.responseText || "";
                httpRequest = null;
            }
            else {
                // Handle non-200 status codes
            }
        }
        catch (e) {
            // Handle request errors
        }

        // Process loaded string
        // alert("loaded=" + loadedString);
        SocialCalc.TestWorkBookSaveStr = loadedString;
        SocialCalc.Clipboard.clipboard = loadedString;
    }
};

/**
 * Make an AJAX call to the server with workbook data
 * @param {string} url - The URL to send the request to (currently unused)
 * @param {string} contents - The data to send in the POST request
 * @returns {boolean} True if request was successfully initiated, false otherwise
 */
SocialCalc.WorkBookControlAjaxCall = function (url, contents) {
    let httpRequest = null;

    alert("in ajax");

    if (window.XMLHttpRequest) {
        // Modern browsers (Mozilla, Safari, Chrome, etc.)
        httpRequest = new XMLHttpRequest();
    }
    else if (window.ActiveXObject) {
        // Internet Explorer
        try {
            httpRequest = new ActiveXObject("Msxml2.XMLHTTP");
        }
        catch (e) {
            try {
                httpRequest = new ActiveXObject("Microsoft.XMLHTTP");
            }
            catch (e) {
                // Failed to create ActiveXObject
            }
        }
    }

    if (!httpRequest) {
        alert('Giving up :( Cannot create an XMLHTTP instance');
        return false;
    }

    // Make the actual request
    SocialCalc.WorkBookControlHttpRequest = httpRequest;

    httpRequest.onreadystatechange = SocialCalc.WorkBookControlAlertContents;
    httpRequest.open('POST', document.URL, true); // async
    httpRequest.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
    httpRequest.send(contents);

    return true;
};

/**
 * Save the current workbook sheets to a JSON string
 * @returns {string} JSON string representation of the workbook data
 */
SocialCalc.WorkBookControlSaveSheet = function () {
    const control = SocialCalc.GetCurrentWorkBookControl();

    const sheetSaveData = {};

    sheetSaveData.numsheets = control.numSheets;
    sheetSaveData.currentid = control.currentSheetButton.id;
    sheetSaveData.currentname = control.currentSheetButton.value;

    sheetSaveData.sheetArr = {};

    for (const sheetKey in control.sheetButtonArr) {
        const sheetString = control.workbook.SaveWorkBookSheet(sheetKey);
        sheetSaveData.sheetArr[sheetKey] = {};
        sheetSaveData.sheetArr[sheetKey].sheetstr = sheetString;
        sheetSaveData.sheetArr[sheetKey].name = control.sheetButtonArr[sheetKey].value;
        sheetSaveData.sheetArr[sheetKey].hidden =
            document.getElementById(`sbsb-${sheetKey}`).style.display === "none" ? "1" : "0";
    }

    // Save the editable cells if specified
    if (SocialCalc.EditableCells && SocialCalc.EditableCells.allow) {
        sheetSaveData.EditableCells = {};
        for (const key in SocialCalc.EditableCells) {
            sheetSaveData.EditableCells[key] = SocialCalc.EditableCells[key];
        }
    }

    SocialCalc.TestWorkBookSaveStr = JSON.stringify(sheetSaveData);
    // alert(SocialCalc.TestWorkBookSaveStr);

    // Send to backend - commented out for now
    // SocialCalc.WorkBookControlAjaxCall("/", "&sheetdata=" + encodeURIComponent(SocialCalc.TestWorkBookSaveStr));

    return SocialCalc.TestWorkBookSaveStr;
};

/**
 * Insert another workbook into an existing workbook
 * Assumes at least 1 sheet exists in existing workbook
 * Sheets with same names will be overwritten
 * @param {string} savestr - JSON string containing workbook data to insert
 * @returns {void}
 */
SocialCalc.WorkBookControlInsertWorkbook = function (savestr) {
    let sheetSaveData;

    if (savestr) {
        sheetSaveData = JSON.parse(savestr);
    }

    const control = SocialCalc.GetCurrentWorkBookControl();

    for (const sheetKey in sheetSaveData.sheetArr) {
        let sheetDataString = sheetSaveData.sheetArr[sheetKey].sheetstr.savestr;
        const parts = control.workbook.spreadsheet.DecodeSpreadsheetSave(sheetDataString);

        if (parts && parts.sheet) {
            sheetDataString = sheetDataString.substring(parts.sheet.start, parts.sheet.end);
        }

        // Check if sheet name already exists
        const sheetName = sheetSaveData.sheetArr[sheetKey].name;
        let sheetId = control.workbook.SheetNameExistsInWorkBook(sheetName);

        if (sheetId) {
            // Sheet exists, overwrite it
            // alert(sheetName + "exists");
            control.workbook.LoadRenameWorkBookSheet(sheetId, sheetDataString, sheetName);
        }
        else {
            // Create new sheet
            sheetId = `sheet${(control.sheetCnt + 1).toString()}`;
            control.sheetCnt = control.sheetCnt + 1;
            SocialCalc.WorkBookControlAddSheetButton(sheetSaveData.sheetArr[sheetKey].name, sheetId);
            // Create the sheet
            control.workbook.AddNewWorkBookSheetNoSwitch(sheetId, sheetSaveData.sheetArr[sheetKey].name, sheetDataString);
        }
    }
};

/**
 * Load a workbook from saved data
 * @param {string} savestr - JSON string containing workbook data to load
 * @returns {void}
 */
SocialCalc.WorkBookControlLoad = function (savestr) {
    let sheetSaveData;

    if (savestr === "") return;

    if (savestr) {
        sheetSaveData = JSON.parse(savestr);
    } else {
        sheetSaveData = JSON.parse(SocialCalc.TestWorkBookSaveStr);
    }

    // alert(sheetSaveData.currentid + "," + sheetSaveData.currentname);

    // First create a new workbook
    const control = SocialCalc.GetCurrentWorkBookControl();

    SocialCalc.WorkBookControlCreateNewBook();

    // At this point there is one sheet and 1 button
    // Create the sequence of buttons and sheets
    let firstRun = true;
    let newButtonCount = 0;
    let sheetId = null;
    let currentSheetId = sheetSaveData.currentid;

    // alert("button=" + newButtonCount);

    for (const sheetKey in sheetSaveData.sheetArr) {
        // alert(sheetKey);
        if (newButtonCount > sheetSaveData.numsheets) {
            break;
        }

        // alert("button=" + newButtonCount);
        let sheetDataString = sheetSaveData.sheetArr[sheetKey].sheetstr.savestr;
        const parts = control.workbook.spreadsheet.DecodeSpreadsheetSave(sheetDataString);

        if (parts && parts.sheet) {
            sheetDataString = sheetDataString.substring(parts.sheet.start, parts.sheet.end);
        }

        if (firstRun) {
            firstRun = false;
            // Set the first button's name correctly
            sheetId = control.currentSheetButton.id;
            control.currentSheetButton.value = sheetSaveData.sheetArr[sheetKey].name;
            SocialCalc.SheetBarButtonSetName(sheetId, sheetSaveData.sheetArr[sheetKey].name);
            // Set the sheet data for the first sheet which already exists
            control.workbook.LoadRenameWorkBookSheet(sheetId, sheetDataString, control.currentSheetButton.value);
            // Need to also set the formula cache
            currentSheetId = sheetId;
        }
        else {
            sheetId = `sheet${(control.sheetCnt + 1).toString()}`;
            control.sheetCnt = control.sheetCnt + 1;
            SocialCalc.WorkBookControlAddSheetButton(sheetSaveData.sheetArr[sheetKey].name, sheetId);
            // Create the sheet
            control.workbook.AddNewWorkBookSheetNoSwitch(sheetId, sheetSaveData.sheetArr[sheetKey].name, sheetDataString);
        }

        if (sheetSaveData.sheetArr[sheetKey].hidden === "1") {
            // Hide the sheet button
            const sheetbarButton = document.getElementById(`sbsb-${sheetId}`);
            sheetbarButton.style.display = "none";
            SocialCalc.SheetBarButtonActivate(sheetKey, false);
            newButtonCount = newButtonCount - 1;
        }

        if (sheetKey === sheetSaveData.currentid) {
            currentSheetId = sheetId;
        }

        newButtonCount = newButtonCount + 1;
    }

    // Save the user script data
    if (sheetSaveData.EditableCells) {
        SocialCalc.EditableCells = {};
        for (const key in sheetSaveData.EditableCells) {
            SocialCalc.EditableCells[key] = sheetSaveData.EditableCells[key];
        }
    }

    const timeoutFunction = function () {
        SocialCalc.WorkBookControlActivateSheet(currentSheetId);
    };
    window.setTimeout(timeoutFunction, 200);
};

/**
 * Rename the current active sheet
 * Shows rename dialog if not in edit mode
 * @returns {void}
 */
SocialCalc.WorkBookControlRenameSheet = function () {
    const control = SocialCalc.GetCurrentWorkBookControl();

    if (control.workbook.spreadsheet.editor.state !== "start") {
        // If in edit mode, return
        return;
    }

    // Create popup dialog to get the new name of the sheet
    // The popup has an input element with submit and cancel buttons
    const existingDialog = document.getElementById(control.renameDialogId);
    if (existingDialog) return;

    const currentSheetName = control.currentSheetButton.value;
    const dialogContent =
        '<div style="padding:6px 0px 4px 6px;">' +
        `<span style="font-size:smaller;">Rename - ${currentSheetName}</span><br>` +
        '<span style="font-size:smaller;">Please ensure that you DO NOT have ANY spaces in the sheet name.</span>' +
        `<input type="text" id="newSheetName" style="width:380px;" value="${currentSheetName}"><br>` +
        '</div>' +
        '<div style="width:380px;text-align:right;padding:6px 0px 4px 6px;font-size:small;">' +
        '<input type="button" value="Submit" style="font-size:smaller;" onclick="SocialCalc.WorkBookControlRenameSheetSubmit();">&nbsp;' +
        '<input type="button" value="Cancel" style="font-size:smaller;" onclick="SocialCalc.WorkBookControlRenameSheetHide();"></div>';

    SocialCalc.createDialog(control.renameDialogId, dialogContent, control.workbook.spreadsheet.spreadsheetDiv);

    const inputElement = document.getElementById("newSheetName");
    inputElement.focus();
    SocialCalc.CmdGotFocus(inputElement);
};

/**
 * Hide the rename sheet dialog and clean up
 * @returns {void}
 */
SocialCalc.WorkBookControlRenameSheetHide = function () {
    const control = SocialCalc.GetCurrentWorkBookControl();
    const spreadsheet = control.workbook.spreadsheet;

    const dialogElement = document.getElementById(control.renameDialogId);

    if (!dialogElement) return;

    dialogElement.innerHTML = "";
    SocialCalc.DragUnregister(dialogElement);
    SocialCalc.KeyboardFocus();

    if (dialogElement.parentNode) {
        dialogElement.parentNode.removeChild(dialogElement);
    }
};

/**
 * Submit the sheet rename operation with validation
 * Validates input, checks for duplicates, and performs the rename
 * @returns {void}
 */
SocialCalc.WorkBookControlRenameSheetSubmit = function () {
    const inputElement = document.getElementById("newSheetName");
    // console.log(inputElement.value);

    const control = SocialCalc.GetCurrentWorkBookControl();

    if (inputElement.value.length === 0) {
        inputElement.focus();
        return;
    }

    const oldName = control.currentSheetButton.value;
    const newName = inputElement.value;

    if (newName.indexOf(" ") !== -1) {
        alert("A space was found in the new name. Please ensure that the new name has no spaces");
        return;
    }

    SocialCalc.WorkBookControlRenameSheetHide();

    // Verify newname does not clash with any existing sheet name
    const normalizedNewName = newName.toLowerCase(); // Converting to lower case to normalize
    // console.log(normalizedNewName + " old " + inputElement.value);

    for (const sheetKey in workbook.sheetArr) {
        console.log(workbook.sheetArr[sheetKey].sheet.sheetname); // Checking in sheetArr for repeated names
        if (workbook.sheetArr[sheetKey].sheet.sheetname === normalizedNewName) {
            alert(`${newName} already exists`);
            return;
        }
    } // Variation of case in letters of a sheet name will give an error if normalizedNewName is used.

    control.currentSheetButton.value = normalizedNewName;
    SocialCalc.SheetBarButtonSetName(control.currentSheetButton.id, newName);

    // Perform a rename for formula references to this sheet in all the 
    // sheets in the workbook
    control.workbook.RenameWorkBookSheet(oldName, normalizedNewName, control.currentSheetButton.id);

    const commandString = `rensheet ${control.currentSheetButton.id} ${oldName} ${newName}`;
    // console.log(commandString);
    SocialCalc.Callbacks.broadcast('execute', {
        cmdtype: "wcmd",
        id: "0",
        cmdstr: commandString
    });
};
/**
 * Handle remote sheet rename operation
 * Updates button name and sheet references from remote command
 * @param {string} sheetid - ID of the sheet to rename
 * @param {string} oldname - Current name of the sheet
 * @param {string} newname - New name for the sheet
 * @returns {void}
 */
SocialCalc.WorkBookControlRenameSheetRemote = function(sheetid, oldname, newname) {
    // console.log("rename sheet ", sheetid, oldname, newname);
    const control = SocialCalc.GetCurrentWorkBookControl();

    const containerElement = document.getElementById("fooBar");
    const renameButton = document.getElementById(sheetid);
    
    renameButton.value = newname;
    SocialCalc.SheetBarButtonSetName(sheetid, newname);
    control.workbook.RenameWorkBookSheet(oldname, newname, sheetid);
};

/**
 * Create a new workbook by resetting to single default sheet
 * Deletes all sheets except current one and resets it to default state
 * @returns {void}
 */
SocialCalc.WorkBookControlCreateNewBook = function() {
    const control = SocialCalc.GetCurrentWorkBookControl();
    
    // Delete all the sheets except current one
    for (const sheetKey in control.sheetButtonArr) {
        if (sheetKey !== control.currentSheetButton.id) {
            control.workbook.DeleteWorkBookSheet(
                control.sheetButtonArr[sheetKey].id, 
                control.sheetButtonArr[sheetKey].value
            );
        }
    }
    
    // Reset that one remaining sheet
    control.workbook.LoadRenameWorkBookSheet(
        control.currentSheetButton.id, 
        "", 
        control.workbook.defaultsheetname
    );
    
    // Delete all the buttons except current one
    for (const sheetKey in control.sheetButtonArr) {
        if (sheetKey !== control.currentSheetButton.id) {
            const containerElement = document.getElementById("fooBar");
            const buttonElement = document.getElementById(control.sheetButtonArr[sheetKey].id);
            
            const buttonId = buttonElement.id;
            delete control.sheetButtonArr[buttonId];
            
            containerElement.removeChild(buttonElement);

            const sheetbar = document.getElementById("SocialCalc-sheetbar-buttons");
            const sheetbarButton = document.getElementById(`sbsb-${buttonId}`);
            // TODO: unregister with mouse events etc
            sheetbar.removeChild(sheetbarButton);

            control.numSheets = control.numSheets - 1;
        }
    }
    
    // Rename the remaining button
    control.currentSheetButton.value = control.workbook.defaultsheetname;   
    // alert("done new workbook");
};

/**
 * Create new workbook and render it
 * @returns {void}
 */
SocialCalc.WorkBookControlNewBook = function() {
    const control = SocialCalc.GetCurrentWorkBookControl();
    SocialCalc.WorkBookControlCreateNewBook();
    control.workbook.RenderWorkBookSheet();
};

/**
 * Move current sheet in specified direction (left or right)
 * @param {string} direction - Direction to move ("left" or "right")
 * @returns {void}
 */
SocialCalc.WorkBookControlMove = function(direction) {
    const control = SocialCalc.GetCurrentWorkBookControl();
    
    if (control.workbook.spreadsheet.editor.state !== "start") {
        return;
    }
    
    const sheetButtonArray = control.sheetButtonArr;
    const newSheetArray = {};
    const currentSheetId = control.currentSheetButton.id;

    const currentButton = document.getElementById(currentSheetId);
    const currentSheetBarButton = document.getElementById(`sbsb-${currentSheetId}`);
    let siblingButton = null;
    let siblingSheetBarButton = null;
    
    if (direction === "left") {
        siblingButton = currentButton.previousSibling;
        siblingSheetBarButton = currentSheetBarButton.previousSibling;
        if (!siblingSheetBarButton) {
            alert("Cannot move leftmost Sheet further to the left");
            return;
        }
    }
    else {
        siblingButton = currentButton.nextSibling;
        siblingSheetBarButton = currentSheetBarButton.nextSibling;   
        if (!siblingSheetBarButton) {
            alert("Cannot move rightmost Sheet further to the right");
            return;
        }  
    }
    
    const currentId = currentSheetId;
    const siblingId = siblingButton.id;
    const parentElement = currentButton.parentNode;
    const sheetBarParent = currentSheetBarButton.parentNode;

    const clonedButtons = {};
    const clonedSheetBarButtons = {};
    
    // Remove all buttons temporarily
    for (const buttonKey in sheetButtonArray) {
        clonedSheetBarButtons[buttonKey] = document.getElementById(`sbsb-${buttonKey}`);
        clonedButtons[buttonKey] = document.getElementById(buttonKey);
        sheetBarParent.removeChild(document.getElementById(`sbsb-${buttonKey}`));
        parentElement.removeChild(document.getElementById(buttonKey));
    }
    
    // Re-add buttons in new order
    for (const buttonKey in sheetButtonArray) {
        if (buttonKey !== currentId && buttonKey !== siblingId) {
            newSheetArray[buttonKey] = sheetButtonArray[buttonKey];
            sheetBarParent.appendChild(clonedSheetBarButtons[buttonKey]);
            parentElement.appendChild(clonedButtons[buttonKey]);
        }
        else if (buttonKey === currentId) {
            if (direction === "left") {
                newSheetArray[currentId] = sheetButtonArray[currentId];
                newSheetArray[siblingId] = sheetButtonArray[siblingId];
                sheetBarParent.appendChild(clonedSheetBarButtons[currentId]);
                parentElement.appendChild(clonedButtons[currentId]);
                sheetBarParent.appendChild(clonedSheetBarButtons[siblingId]);
                parentElement.appendChild(clonedButtons[siblingId]);
            } else {
                newSheetArray[siblingId] = sheetButtonArray[siblingId];
                newSheetArray[currentId] = sheetButtonArray[currentId];
                sheetBarParent.appendChild(clonedSheetBarButtons[siblingId]);
                parentElement.appendChild(clonedButtons[siblingId]);
                sheetBarParent.appendChild(clonedSheetBarButtons[currentId]);
                parentElement.appendChild(clonedButtons[currentId]);
            }
        }
    }
    
    control.sheetButtonArr = newSheetArray;
    SocialCalc.SheetBarButtonActivate(currentId, true);
};

/**
 * Move current sheet to the left
 * @returns {void}
 */
SocialCalc.WorkBookControlMoveLeft = function() {
    SocialCalc.WorkBookControlMove("left");
};

/**
 * Move current sheet to the right
 * @returns {void}
 */
SocialCalc.WorkBookControlMoveRight = function() {
    SocialCalc.WorkBookControlMove("right");
};

/**
 * Copy the current active sheet to clipboard
 * @returns {void}
 */
SocialCalc.WorkBookControlCopySheet = function() {
    // alert("in copy");

    const control = SocialCalc.GetCurrentWorkBookControl();

    if (control.workbook.spreadsheet.editor.state !== "start") {
        // If in edit mode, return
        return;
    }   

    control.workbook.CopyWorkBookSheet(control.currentSheetButton.id);
    alert(`copied sheet: ${control.currentSheetButton.value}`);
};

/**
 * Paste copied sheet as new sheet in workbook
 * @returns {void}
 */
SocialCalc.WorkBookControlPasteSheet = function() {
    // alert("in paste");

    const control = SocialCalc.GetCurrentWorkBookControl();

    if (control.workbook.spreadsheet.editor.state !== "start") {
        // If in edit mode, return
        return;
    }   

    const oldSheetId = control.currentSheetButton.id;
    
    SocialCalc.WorkBookControlAddSheet(false);
    
    const newSheetId = control.currentSheetButton.id;
    
    // alert(newSheetId + oldSheetId);
    
    control.workbook.PasteWorkBookSheet(newSheetId, oldSheetId);

    const commandString = "addsheetstr";
    SocialCalc.Callbacks.broadcast('execute', { 
        cmdtype: "wcmd", 
        id: "0", 
        cmdstr: commandString, 
        sheetstr: control.workbook.clipsheet.savestr 
    });
};

/**
 * Sheet Bar constructor - creates and manages the sheet tab bar UI
 * @class
 */
SocialCalc.SheetBar = function() {
    /** @type {HTMLElement} Base container div for the sheet bar */
    this.baseDiv = document.getElementById("SocialCalc-sheetbar");

    /** @type {HTMLElement} Pre-buttons spacing div */
    this.prebuttonsDiv = document.createElement("div");
    this.prebuttonsDiv.style.cssText = "display:inline;"; 
    this.prebuttonsDiv.innerHTML = "&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp";
    this.prebuttonsDiv.id = "SocialCalc-sheetbar-prebuttons";

    /** @type {HTMLElement} Container for sheet buttons */
    this.buttonsDiv = document.createElement("div");
    this.buttonsDiv.id = "SocialCalc-sheetbar-buttons";
    this.buttonsDiv.style.cssText = "display:inline;"; 

    /** @type {HTMLElement} Container for action buttons */
    this.buttonActionsDiv = document.createElement("div");
    this.buttonActionsDiv.id = "SocialCalc-sheetbar-buttonactions";
    this.buttonActionsDiv.style.display = "inline"; 
    
    const addButton = new SocialCalc.SheetBarSheetButton(
        "sbsba-add", 
        "sbsba-add", 
        this.buttonActionsDiv, 
        {}, 
        {
            MouseDown: function() { 
                const result = SocialCalc.WorkBookControlAddSheet(true); 
            }
        },
        "add-2.png"
    );

    this.baseDiv.appendChild(this.prebuttonsDiv);
    this.baseDiv.appendChild(this.buttonsDiv);
    this.baseDiv.appendChild(this.buttonActionsDiv);
};
/**
 * Sheet Bar Sheet Button constructor - creates interactive sheet tab buttons
 * @class
 * @param {string} id - Unique identifier for the button
 * @param {string} name - Display name for the button
 * @param {HTMLElement} parentdiv - Parent container element
 * @param {Object} params - Button styling parameters
 * @param {Object} functions - Event handler functions
 * @param {string} [img] - Optional image filename for icon-only buttons
 */
SocialCalc.SheetBarSheetButton = function(id, name, parentdiv, params, functions, img) {
    /** @type {HTMLElement} The button element */
    this.ele = document.createElement("div");
    this.ele.id = id;
    this.ele.name = name;
    
    if (!img) {
        // Text button with dropdown arrow
        this.ele.innerHTML = name;
        this.ele.style.cssText = "font-size:small;display:inline;padding:5px 5px 2px 5px;border:1px solid #000;";
        
        const imageElement = document.createElement("img");
        imageElement.id = `${id}-img`;
        imageElement.src = `${SocialCalc.Constants.defaultImagePrefix}menu-dropdown.png`;
        imageElement.style.cssText = "padding:0px 2px;width:16px;height:16px;vertical-align:middle;";
        
        this.ele.appendChild(imageElement);
        SocialCalc.ButtonRegister(this.ele, params, functions);
        SocialCalc.ButtonRegister(imageElement, params, functions);
    } else {
        // Icon-only button
        const imageElement = document.createElement("img");
        imageElement.src = `${SocialCalc.Constants.defaultImagePrefix}${img}`;
        imageElement.style.cssText = "width:16px;height:16px;vertical-align:middle;";
        
        this.ele.appendChild(imageElement);
        this.ele.style.cssText = "display:inline;padding:5px 5px 2px 5px;";
        SocialCalc.ButtonRegister(imageElement, params, functions);
    }
    
    parentdiv.appendChild(this.ele);
};

/**
 * Activate or deactivate a sheet bar button
 * @param {string} id - ID of the sheet to activate/deactivate
 * @param {boolean} active - Whether to activate (true) or deactivate (false) the button
 * @returns {void}
 */
SocialCalc.SheetBarButtonActivate = function(id, active) {
    const sheetButton = document.getElementById(`sbsb-${id}`);
    sheetButton.isactive = active;
    
    if (active) {
        sheetButton.style.backgroundColor = "#FFF";
        let imageElement = document.getElementById(`sbsb-${id}-img`);
        
        if (!imageElement) {
            imageElement = document.createElement("img");
            imageElement.id = `sbsb-${id}-img`;
            imageElement.src = `${SocialCalc.Constants.defaultImagePrefix}menu-dropdown.png`;
            imageElement.style.cssText = "padding:0px 2px;width:16px;height:16px;vertical-align:middle;";
        }
        
        sheetButton.appendChild(imageElement);
        SocialCalc.ButtonRegister(imageElement, {}, {
            MouseDown: function() { SocialCalc.SheetBarSheetButtonPress(id); },
            Repeat: function() {},
            Disabled: function() {}
        });
    } 
    else {
        sheetButton.style.backgroundColor = "#CCC";
        const imageElement = document.getElementById(`sbsb-${id}-img`);
        if (imageElement) {
            sheetButton.removeChild(imageElement);
        }
    }
    
    const menu = document.getElementById("sbsb-menu");
    if (menu && menu.style.display !== "none") {
        menu.style.display = "none";
    }
};

/**
 * Set the display name for a sheet bar button
 * @param {string} id - ID of the sheet button
 * @param {string} name - New name to display
 * @returns {void}
 */
SocialCalc.SheetBarButtonSetName = function(id, name) {
    const sheetButton = document.getElementById(`sbsb-${id}`);
    sheetButton.name = name;
    sheetButton.innerHTML = name;
    
    if (sheetButton.isactive) {
        SocialCalc.SheetBarButtonActivate(id, true);
    }
};

/**
 * Handle sheet bar button press events
 * Shows context menu for active buttons, activates sheet for inactive buttons
 * @param {string} id - ID of the pressed sheet button
 * @returns {void}
 */
SocialCalc.SheetBarSheetButtonPress = function(id) {
    // console.log("button press");
    const sheetButton = document.getElementById(`sbsb-${id}`);
    
    if (sheetButton && sheetButton.isactive) {
        let menu = document.getElementById("sbsb-menu");
        
        if (!menu) {
            const sheetButtonMenu = new SocialCalc.SheetBarSheetButtonMenu("sbsb-menu", id);        
        } else {
            menu.clickedsheetid = id;
            if (menu.style.display === "none") {
                menu.style.display = "inline";
                SocialCalc.SheetBarSheetButtonMenuPosition(menu, id);
            } else {
                menu.style.display = "none";
            }
        }
    } else if (sheetButton) {
        SocialCalc.WorkBookControlActivateSheet(id);
    }    
};

/**
 * Sheet Bar Sheet Button Menu Item constructor - creates context menu items
 * @class
 * @param {string} id - Unique identifier for the menu item
 * @param {string} text - Display text for the menu item
 * @returns {HTMLElement} The created menu item element
 */
SocialCalc.SheetBarSheetButtonMenuItem = function(id, text) {
    /** @type {HTMLElement} The menu item element */
    this.ele = document.createElement("div");
    this.ele.id = id;
    this.ele.innerHTML = text;
    this.ele.className = "";
    this.ele.style.cssText = "padding:3px 4px;width:100px;height:20px;background-color:#FFF;";

    const buttonParams = {
        normalstyle: "backgroundColor:#FFF;",
        downstyle: "backgroundColor:#CCC;",
        hoverstyle: "backgroundColor:#CCC;"
    };
    
    const buttonFunctions = {
        MouseDown: function() { SocialCalc.SheetBarMenuItemPress(id); },
        Repeat: function() {},
        Disabled: function() {}
    };

    SocialCalc.ButtonRegister(this.ele, buttonParams, buttonFunctions);
    SocialCalc.TouchRegister(this.ele, { SingleTap: buttonFunctions.MouseDown });

    return this.ele;
};

/**
 * Handle menu item press events and execute corresponding actions
 * @param {string} id - ID of the pressed menu item
 * @returns {void}
 */
SocialCalc.SheetBarMenuItemPress = function(id) {
    const menu = document.getElementById("sbsb-menu");
    if (!menu) return;

    const clickedSheetId = menu.clickedsheetid;

    switch (id) {
        case "sbsb_deletesheet":
            // console.log("delete " + clickedSheetId);
            SocialCalc.WorkBookControlDelSheet();
            break;
            
        case "sbsb_hidesheet":
            // console.log("hide" + clickedSheetId);
            SocialCalc.WorkBookControlHideSheet();
            break;
            
        case "sbsb_unhidesheet":
            // console.log("unhide" + clickedSheetId);
            SocialCalc.WorkBookControlUnhideSheet();
            break;
            
        case "sbsb_copysheet":
            // console.log("copy " + clickedSheetId);
            SocialCalc.WorkBookControlCopySheet();
            break;
            
        case "sbsb_moveleft":
            // console.log("move left " + clickedSheetId);    
            SocialCalc.WorkBookControlMoveLeft();
            break;
            
        case "sbsb_moveright":
            // console.log("move right " + clickedSheetId);    
            SocialCalc.WorkBookControlMoveRight();
            break;
            
        case "sbsb_pastesheet":
            // console.log("paste " + clickedSheetId);
            SocialCalc.WorkBookControlPasteSheet();
            break;
            
        case "sbsb_renamesheet":
            // console.log("rename " + clickedSheetId);    
            SocialCalc.WorkBookControlRenameSheet();
            break;
            
        case "sbsb_closemenu":
            // console.log("close menu " + clickedSheetId);    
            menu.style.display = "none";
            break;
            
        default:
            break;
    }
    
    menu.style.display = "none";
};

/**
 * Sheet Bar Sheet Button Menu constructor - creates context menu for sheet operations
 * @class
 * @param {string} id - Unique identifier for the menu
 * @param {string} clickedsheetid - ID of the sheet that was clicked to open this menu
 */
SocialCalc.SheetBarSheetButtonMenu = function(id, clickedsheetid) {
    /** @type {HTMLElement} The menu container element */
    this.ele = document.createElement("div");
    this.ele.id = id;
    this.ele.className = "";
    this.ele.clickedsheetid = clickedsheetid;
    this.ele.style.cssText = "border:1px solid #000;position:absolute;top:200px;left:0px;width=100px;z-index:120";

    // Create and append menu items
    const menuItems = [
        { id: "sbsb_deletesheet", text: " Delete Sheet" },
        { id: "sbsb_hidesheet", text: " Hide Sheet " },
        { id: "sbsb_unhidesheet", text: " Unhide Sheet " },
        { id: "sbsb_renamesheet", text: " Rename Sheet " },
        { id: "sbsb_moveleft", text: " Move Left " },
        { id: "sbsb_moveright", text: " Move Right " },
        { id: "sbsb_copysheet", text: " Copy Sheet " },
        { id: "sbsb_pastesheet", text: " Paste Sheet " },
        { id: "sbsb_closemenu", text: " Cancel" }
    ];

    menuItems.forEach(item => {
        const menuItemElement = new SocialCalc.SheetBarSheetButtonMenuItem(item.id, item.text);
        this.ele.appendChild(menuItemElement);
    });
    
        SocialCalc.SheetBarSheetButtonMenuPosition(this.ele, clickedsheetid);

    // Commented out positioning code - preserved for reference
    // const clickedsheet = document.getElementById(clickedsheetid);
    // const position = SocialCalc.GetElementPosition(clickedsheet);
    // console.log(clickedsheet.offsetHeight, clickedsheet.offsetWidth, clickedsheet.offsetLeft, clickedsheet.offsetTop);

    const control = SocialCalc.GetCurrentWorkBookControl();
    control.workbook.spreadsheet.editor.toplevel.appendChild(this.ele);
};

/**
 * Position the sheet context menu relative to the clicked button
 * @param {HTMLElement} menu - The menu element to position
 * @param {string} clickedsheetid - ID of the sheet button that was clicked
 * @returns {void}
 */
SocialCalc.SheetBarSheetButtonMenuPosition = function(menu, clickedsheetid) {
    const horizontalLessButton = document.getElementById("te_lessbuttonh");
    
    // console.log(horizontalLessButton.style.top, horizontalLessButton.style.left);

    const sheetButton = document.getElementById(`sbsb-${clickedsheetid}`);    

    // console.log(sheetButton.offsetLeft, clickedsheetid);

    const topPosition = horizontalLessButton.style.top.slice(0, -2) - 220;
    const leftPosition = sheetButton.offsetLeft + 7;

    menu.style.top = `${topPosition}px`;
    menu.style.left = `${leftPosition}px`;

    // console.log(menu.style.top, menu.style.left);    
};

/**
 * Script information storage and handling
 * @namespace
 * @property {Object} scripts - Collection of user scripts by coordinate
 * @property {number|null} handle - Timer handle for script evaluation
 */
SocialCalc.ScriptInfo = {
    scripts: {},
    handle: null
};

/**
 * Check cell text for embedded scripts and queue them for evaluation
 * @param {string} sheetid - ID of the sheet containing the cell
 * @param {string} coord - Cell coordinate (e.g., "A1")
 * @param {string} text - Cell text content to check for scripts
 * @returns {void}
 */
SocialCalc.ScriptCheck = function(sheetid, coord, text) {
    const commentStart = text.indexOf("<!--script");
    const commentEnd = text.indexOf("script-->");
    
    if ((commentStart !== -1) && (commentEnd !== -1)) {
        const script = text.slice(commentStart + 10, commentEnd);
        // alert(script);
        SocialCalc.ScriptInfo.scripts[coord] = script;
        
        if (SocialCalc.ScriptInfo.handle === null) {
            SocialCalc.ScriptInfo.handle = window.setTimeout(SocialCalc.EvalUserScripts, 500);
        }
        // alert(coord + "-" + sheetid);
    }
};

/**
 * Evaluate a single user script by injecting it into the document head
 * @param {string} data - JavaScript code to evaluate
 * @returns {void}
 */
SocialCalc.EvalUserScript = function(data) {
    const head = document.getElementsByTagName("head")[0] || document.documentElement;

    if (data === "") return;

    const scriptElement = document.createElement("script");
    scriptElement.type = "text/javascript";
    
    try {
        // Standard approach - doesn't work on older IE
        scriptElement.appendChild(document.createTextNode(data));      
    } catch (e) {
        // IE has different script node handling
        scriptElement.text = data;
    }

    head.insertBefore(scriptElement, head.firstChild);
    head.removeChild(scriptElement);   
};

/**
 * Evaluate all queued user scripts and clear the queue
 * @returns {void}
 */
SocialCalc.EvalUserScripts = function() {
    for (const coordinate in SocialCalc.ScriptInfo.scripts) {
        SocialCalc.EvalUserScript(SocialCalc.ScriptInfo.scripts[coordinate]);
        // console.log(coordinate, SocialCalc.ScriptInfo.scripts[coordinate]);
    }
    
    SocialCalc.ScriptInfo.handle = null;
    SocialCalc.ScriptInfo.scripts = {};
};

/**
 * Callback function called when rendering cells to check for embedded scripts
 * @param {Object} sheetobj - Sheet object containing cell data
 * @param {string} value - Rendered value of the cell
 * @param {string} cr - Cell coordinate
 * @returns {void}
 */
SocialCalc.CallOutOnRenderCell = function(sheetobj, value, cr) {
    const cell = sheetobj.cells[cr];
    if (!cell) return;
    
    const valueType = cell.valuetype || ""; // get type of value to determine formatting
    const valueSubtype = valueType.substring(1);
    const sheetAttributes = sheetobj.attribs;
    const mainValueType = valueType.charAt(0);
    
    if (mainValueType === "t") {
        const valueFormat = sheetobj.valueformats[cell.textvalueformat - 0] || 
                           sheetobj.valueformats[sheetAttributes.defaulttextvalueformat - 0] || 
                           "";
        
        if (valueFormat === "text-html") {
            SocialCalc.ScriptCheck(sheetobj.sheetid, cr, value);
        }
    }
};

/**
 * Get the data value from a cell, supporting cross-sheet references
 * @param {string} coord - Cell coordinate, optionally prefixed with sheet name (e.g., "Sheet1!A1" or "A1")
 * @returns {string|number} The cell's data value, or "0" if cell doesn't exist
 */
SocialCalc.GetCellDataValue = function(coord) {
    let sheetName = null;
    let sheetId = "";
    
    const bangIndex = coord.indexOf("!");
    if (bangIndex !== -1) {
        sheetName = coord.slice(0, bangIndex);
        coord = coord.slice(bangIndex + 1);
        // console.log(sheetName, coord);
    }
    
    const control = SocialCalc.GetCurrentWorkBookControl();
    
    if (sheetName === null) {
        sheetId = control.currentSheetButton.id;
    } else {
        sheetId = control.workbook.SheetNameExistsInWorkBook(sheetName);
    }
    
    if ((sheetId === null) || (sheetId === "")) {
        return "0";
    }
    
    const sheetObject = control.workbook.sheetArr[sheetId].sheet;
    const cell = sheetObject.cells[coord];
    
    if (cell) {
        return cell.datavalue;
    } else {
        return 0;
    }
};

/**
 * Get data values from multiple cells as an array
 * @param {string} coordstr - Comma-separated list of cell coordinates
 * @param {string|null} sheetname - Optional sheet name to prefix to all coordinates
 * @returns {Array} Array of cell data values
 */
SocialCalc.GetCellDataArray = function(coordstr, sheetname) {
    const values = [];
    const coordinates = coordstr.split(",");
    
    let sheetPrefix = "";
    if (sheetname !== null) {
        sheetPrefix = `${sheetname}!`;
    }
    
    for (const coordinate of coordinates) {
        values.push(SocialCalc.GetCellDataValue(sheetPrefix + coordinate));
    }
    
    return values;
};

/**
 * Storage for user script data
 * @type {Object}
 */
SocialCalc.UserScriptData = {};

/**
 * Information for workbook-wide recalculation process
 * @namespace
 * @property {Array} sheets - List of sheet IDs to recalculate
 * @property {Array} calcorder - Sheets in reverse order for calculation
 * @property {number} current - Current position in recalculation process
 * @property {number} pass - Current pass number
 */
SocialCalc.WorkBookRecalculateInfo = {
    sheets: [],
    calcorder: [],
    current: 0,
    pass: 0
};

/**
 * Initiate recalculation of all sheets in the workbook
 * Processes sheets in reverse order using recalc-done signal to trigger next sheet
 * @returns {void}
 */
SocialCalc.WorkBookRecalculateAll = function() {
    // If already in the middle of a recalculate-all, ignore this request
    if ((SocialCalc.WorkBookRecalculateInfo.current !== 0) ||
        (SocialCalc.WorkBookRecalculateInfo.calcorder.length !== 0) ||
        (SocialCalc.WorkBookRecalculateInfo.sheets.length !== 0)) {
        return;
    }

    const control = SocialCalc.GetCurrentWorkBookControl();

    if (control.workbook.spreadsheet.editor.state !== "start") {
        // If in edit mode, return
        return;
    }   

    SocialCalc.WorkBookRecalculateInfo.current = 0;

    // Collect all sheet IDs
    for (const sheetKey in control.workbook.sheetArr) {
        SocialCalc.WorkBookRecalculateInfo.sheets.push(sheetKey);
    }
    
    // Create reverse order for calculation
    let index = 0;
    for (let c = SocialCalc.WorkBookRecalculateInfo.sheets.length; c > 0; c--) {
        SocialCalc.WorkBookRecalculateInfo.calcorder[index] = 
            SocialCalc.WorkBookRecalculateInfo.sheets[c - 1];
        index++;
    }
    
    window.setTimeout(SocialCalc.WorkBookRecalculateStep, 500);
};
/**
 * Execute one step of the workbook recalculation process
 * Processes sheets sequentially and handles completion or multi-pass scenarios
 * @returns {void}
 */
SocialCalc.WorkBookRecalculateStep = function() {
    if (SocialCalc.WorkBookRecalculateInfo.current === 
        SocialCalc.WorkBookRecalculateInfo.calcorder.length) {
        
        SocialCalc.WorkBookRecalculateInfo.current = 0;
        SocialCalc.WorkBookRecalculateInfo.calcorder = [];
        SocialCalc.WorkBookRecalculateInfo.sheets = [];
        
        if (SocialCalc.WorkBookRecalculateInfo.pass === 1) {
            SocialCalc.WorkBookRecalculateInfo.pass = 0;
            SocialCalc.SpinnerWaitHide();
            // alert("load done");
            return;
        } else {
            SocialCalc.WorkBookRecalculateInfo.pass++;
            SocialCalc.WorkBookRecalculateAll();
            return;
        }
    }
    
    const control = SocialCalc.GetCurrentWorkBookControl();
    
    // alert("recalculate " + 
    //     SocialCalc.WorkBookRecalculateInfo.calcorder[
    //         SocialCalc.WorkBookRecalculateInfo.current]
    // );
    
    const sheetId = SocialCalc.WorkBookRecalculateInfo.calcorder[
        SocialCalc.WorkBookRecalculateInfo.current
    ];
    
    SocialCalc.WorkBookControlActivateSheet(sheetId);
    SocialCalc.WorkBookRecalculateInfo.current++;

    window.setTimeout(SocialCalc.WorkBookRecalculateStep, 1000);    
};

/**
 * Create a loading spinner overlay for workbook operations
 * @returns {void}
 */
SocialCalc.SpinnerWaitCreate = function() {
    // If the div exists already, just use it
    const existingSpinner = document.getElementById("waitloadingspinner");
    if (existingSpinner) {
        return;
    }
    
    const spinnerElement = document.createElement("div");
    spinnerElement.id = "waitloadingspinner";
    spinnerElement.style.position = "absolute";

    const viewport = SocialCalc.GetViewportInfo();
    spinnerElement.style.top = `${viewport.height / 2}px`;
    spinnerElement.style.left = `${viewport.width / 2}px`;
    spinnerElement.style.zIndex = "110";
    spinnerElement.style.width = "50px";
    spinnerElement.style.height = "50px";
    spinnerElement.innerHTML = '<img src="/src/images/spinner.gif" alt="Loading..." />';

    const control = SocialCalc.GetCurrentWorkBookControl();    
    control.workbook.spreadsheet.spreadsheetDiv.appendChild(spinnerElement);
};

/**
 * Hide and remove the loading spinner overlay
 * @returns {void}
 */
SocialCalc.SpinnerWaitHide = function() {
    const spinnerElement = document.getElementById("waitloadingspinner");
    
    if (spinnerElement) {
        spinnerElement.innerHTML = "";

        if (spinnerElement.parentNode) {
            spinnerElement.parentNode.removeChild(spinnerElement);
        }
    }
};

/**
 * Configuration for editable cells functionality
 * @namespace
 * @property {boolean} allow - Whether to restrict cell editing
 * @property {Object} cells - Map of editable cell references
 * @property {Object} constraints - Map of cell validation constraints
 */
SocialCalc.EditableCells = {};
SocialCalc.EditableCells.allow = false;
SocialCalc.EditableCells.cells = {};
SocialCalc.EditableCells.constraints = {};

/**
 * Callback to determine if a cell is editable
 * @param {Object} editor - The spreadsheet editor object
 * @returns {boolean} True if cell can be edited, false otherwise
 */
SocialCalc.Callbacks.IsCellEditable = function(editor) {
    const cellName = `${editor.workingvalues.currentsheet}!${editor.ecell.coord}`;

    if (!SocialCalc.EditableCells.allow) {
        // By default all cells are editable
        return true;
    }
    
    if (SocialCalc.EditableCells.cells[cellName]) {
        // Cell is explicitly marked as editable
        return true;
    }    
    
    return false;
};

/**
 * Callback to validate cell value against defined constraints
 * @param {Object} editor - The spreadsheet editor object
 * @param {string} value - The value to validate
 * @returns {boolean} True if value passes validation, false otherwise
 */
SocialCalc.Callbacks.CheckConstraints = function(editor, value) {
    const cellName = `${editor.workingvalues.currentsheet}!${editor.ecell.coord}`;
    const constraint = SocialCalc.EditableCells.constraints[cellName];
    
    // alert(JSON.stringify(constraint));
    
    if (constraint != null) {
        switch (constraint[0]) {
            case "numeric":
                // Check that value is numeric
                const numericValue = parseInt(value);
                // alert(numericValue);
                
                if (isNaN(numericValue)) {
                    alert("please input a number");
                    return false;
                }
                break;
                
            case "percent":
                // Percent validation logic would go here
                break;
        }
    }
    
    return true;
};

/**
 * Toggle cell content between checkmark and empty space
 * Used for creating interactive checkbox-like functionality
 * @returns {void}
 */
SocialCalc.Callbacks.ToggleCell = function() {
    // Check what the current ecell is
    const control = SocialCalc.GetCurrentWorkBookControl();
    const editor = control.workbook.spreadsheet.editor;
    const cellCoordinate = editor.ecell.coord;
    const currentValue = SocialCalc.GetCellDataValue(cellCoordinate);   

    console.log(cellCoordinate);
    console.log(currentValue);

    let commandString = "";
    
    if (currentValue.indexOf("&nbsp;") !== -1) {
        // Set the value to checkmark
        // commandString = "set " + cellCoordinate + ' text t <div onclick="SocialCalc.Callbacks.ToggleCell();"><img src="http://www.ftsign.com/images/checkmark__16_black_T.png"></img></div>' + "\n";
        commandString = `set ${cellCoordinate} text t <div onclick="SocialCalc.Callbacks.ToggleCell();">&#10004</div>\n`;   
    } else {
        // Set the value to a space
        commandString = `set ${cellCoordinate} text t <div onclick="SocialCalc.Callbacks.ToggleCell();">&nbsp;</div>\n`;
    }

    const sheetId = control.currentSheetButton.id;
    console.log(sheetId);
    console.log(commandString);
    
    const command = {
        cmdtype: "scmd", 
        id: sheetId, 
        cmdstr: commandString, 
        saveundo: false
    };
    
    control.ExecuteWorkBookControlCommand(command, false);  
};

/**
 * Create HTML representation of workbook sheets
 * @param {Object|null} sheetlist - Optional list of specific sheets to render, or null for current sheet
 * @returns {string} HTML string representation of the sheet(s)
 */
SocialCalc.WorkbookControlCreateSheetHTML = function(sheetlist) {
    let context, renderElement;
    let result = "";
    
    const control = SocialCalc.GetCurrentWorkBookControl();      
    const containerDiv = document.createElement("div");

    if (!sheetlist) {
        context = new SocialCalc.RenderContext(spreadsheet.sheet);
        renderElement = context.RenderSheet(null, { type: "html" });
        containerDiv.appendChild(renderElement);
        delete context;
    } else {
        for (const sheetId in sheetlist) {
            context = control.workbook.sheetArr[sheetId].context;
            renderElement = context.RenderSheet(null, { type: "html" });
            containerDiv.appendChild(renderElement);
        }
    }

    result = containerDiv.innerHTML;
    delete renderElement;
    delete containerDiv;
    
    return result;
};


