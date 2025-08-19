//
// Workbook is a collection of sheets that are worked upon together
// 
// The WorkBook class models and manages the collection of sheets
//
// Author: Ramu Ramamurthy
//
//

if (!window.SocialCalc) {
	alert("Main SocialCalc code module needed");
    window.SocialCalc = {};
}
// const SocialCalc = window.SocialCalc;

/**
 * WorkBook is a collection of sheets that are worked upon together
 * The WorkBook class models and manages the collection of sheets
 * @class
 */
SocialCalc.WorkBook = class {
    /**
     * Creates a new WorkBook instance
     * @param {Object} spread - The spreadsheet control
     */
    constructor(spread) {
        this.spreadsheet = spread; // this is the spreadsheet control
        this.defaultsheetname = null;
        this.sheetArr = {};  // misnomer, this is not really an array
        this.clipsheet = {}; // for copy paste of sheets
    }

    /**
     * Initializes the WorkBook with a default sheet
     * @param {string} defaultsheet - The name of the default sheet
     * @returns {*} Result of the initialization
     */
    InitializeWorkBook(defaultsheet) {
        return SocialCalc.InitializeWorkBook(this, defaultsheet);
    }

    /**
     * Adds a new sheet to the workbook without switching to it
     * @param {string} sheetid - The sheet ID
     * @param {string} sheetname - The sheet name
     * @param {string} savestr - Saved sheet data
     * @returns {*} Result of the operation
     */
    AddNewWorkBookSheetNoSwitch(sheetid, sheetname, savestr) {
        return SocialCalc.AddNewWorkBookSheetNoSwitch(this, sheetid, sheetname, savestr);
    }

    /**
     * Adds a new sheet to the workbook and switches to it
     * @param {string} sheetname - The new sheet name
     * @param {string} oldsheetname - The old sheet name
     * @param {boolean} fromclip - Whether the sheet is from clipboard
     * @param {Object} spread - Optional existing spreadsheet object
     * @returns {*} Result of the operation
     */
    AddNewWorkBookSheet(sheetname, oldsheetname, fromclip, spread) {
        return SocialCalc.AddNewWorkBookSheet(this, sheetname, oldsheetname, fromclip, spread);
    }

    /**
     * Activates a specific sheet in the workbook
     * @param {string} sheetname - The sheet to activate
     * @param {string} oldsheetname - The previously active sheet
     * @returns {*} Result of the activation
     */
    ActivateWorkBookSheet(sheetname, oldsheetname) {
        return SocialCalc.ActivateWorkBookSheet(this, sheetname, oldsheetname);
    }

    /**
     * Deletes a sheet from the workbook
     * @param {string} sheetname - The sheet to delete
     * @param {string} cursheetname - The current sheet name
     * @returns {*} Result of the deletion
     */
    DeleteWorkBookSheet(sheetname, cursheetname) {
        return SocialCalc.DeleteWorkBookSheet(this, sheetname, cursheetname);
    }

    /**
     * Saves a workbook sheet and returns its data
     * @param {string} sheetid - The sheet ID to save
     * @returns {Object} Object containing the saved sheet data
     */
    SaveWorkBookSheet(sheetid) {
        return SocialCalc.SaveWorkBookSheet(this, sheetid);
    }

    /**
     * Loads and renames a workbook sheet
     * @param {string} sheetid - The sheet ID
     * @param {string} savestr - The saved sheet data
     * @param {string} newname - The new name for the sheet
     * @returns {*} Result of the operation
     */
    LoadRenameWorkBookSheet(sheetid, savestr, newname) {
        return SocialCalc.LoadRenameWorkBookSheet(this, sheetid, savestr, newname);
    }

    /**
     * Renames a workbook sheet and updates all formula references
     * @param {string} oldname - The current name of the sheet
     * @param {string} newname - The new name for the sheet
     * @param {string} sheetid - The sheet ID
     * @returns {*} Result of the rename operation
     */
    RenameWorkBookSheet(oldname, newname, sheetid) {
        return SocialCalc.RenameWorkBookSheet(this, oldname, newname, sheetid);
    }

    /**
     * Copies a workbook sheet to the clipboard
     * @param {string} sheetid - The sheet ID to copy
     * @returns {*} Result of the copy operation
     */
    CopyWorkBookSheet(sheetid) {
        return SocialCalc.CopyWorkBookSheet(this, sheetid);
    }

    /**
     * Pastes a workbook sheet from the clipboard
     * @param {string} newid - The new sheet ID
     * @param {string} oldid - The old sheet ID
     * @returns {*} Result of the paste operation
     */
    PasteWorkBookSheet(newid, oldid) {
        return SocialCalc.PasteWorkBookSheet(this, newid, oldid);
    }

    /**
     * Renders the current workbook sheet
     * @returns {*} Result of the render operation
     */
    RenderWorkBookSheet() {
        return SocialCalc.RenderWorkBookSheet(this);
    }

    /**
     * Checks if a sheet name exists in the workbook
     * @param {string} name - The sheet name to check
     * @returns {string|null} The sheet ID if exists, null otherwise
     */
    SheetNameExistsInWorkBook(name) {
        return SocialCalc.SheetNameExistsInWorkBook(this, name);
    }

    /**
     * Schedules a command for the workbook
     * @param {Object} cmd - The command object
     * @param {boolean} isremote - Whether the command is remote
     * @returns {*} Result of the command scheduling
     */
    WorkbookScheduleCommand(cmd, isremote) {
        return SocialCalc.WorkbookScheduleCommand(this, cmd, isremote);
    }

    /**
     * Schedules a sheet command for the workbook
     * @param {Object} cmd - The command object
     * @param {boolean} isremote - Whether the command is remote
     * @returns {*} Result of the sheet command scheduling
     */
    WorkbookScheduleSheetCommand(cmd, isremote) {
        return SocialCalc.WorkbookScheduleSheetCommand(this, cmd, isremote);
    }
}

/**
 * Schedules some command - could be for sheet or for the workbook itself
 * @param {Object} workbook - The workbook instance
 * @param {Object} cmd - The command object with cmdtype property
 * @param {boolean} isremote - Whether the command is remote
 */
SocialCalc.WorkbookScheduleCommand = (workbook, cmd, isremote) => {
    
    //console.log(`cmd ${cmd.cmdstr} ${cmd.cmdtype}`);

    if (cmd.cmdtype === "scmd") {
	workbook.WorkbookScheduleSheetCommand(cmd, isremote);
    }
};

/**
 * Schedules a sheet command for the workbook
 * @param {Object} workbook - The workbook instance
 * @param {Object} cmd - The command object with id, cmdstr, and saveundo properties
 * @param {boolean} isremote - Whether the command is remote
 */
SocialCalc.WorkbookScheduleSheetCommand = (workbook, cmd, isremote) => {
    //console.log(cmd.cmdtype,cmd.id,cmd.cmdstr);
  
    // check if sheet exists first
    if (workbook.sheetArr[cmd.id]) {
	workbook.sheetArr[cmd.id].sheet.ScheduleSheetCommands(
	    cmd.cmdstr,
	    cmd.saveundo,
	    isremote);
    }
};


/**
 * Initializes the WorkBook with a default sheet
 * @param {Object} workbook - The workbook instance to initialize
 * @param {string} defaultsheet - The name of the default sheet
 */
SocialCalc.InitializeWorkBook = (workbook, defaultsheet) => {

   	workbook.defaultsheetname = defaultsheet;
	
	const { spreadsheet } = workbook;
	const { defaultsheetname } = workbook;
	
    // Initialize the Spreadsheet Control and display it

	SocialCalc.Formula.SheetCache.sheets[defaultsheetname] = {sheet: spreadsheet.sheet, name: defaultsheetname}; 

        spreadsheet.sheet.sheetid = defaultsheetname;
        spreadsheet.sheet.sheetname = defaultsheetname;

   	workbook.sheetArr[defaultsheetname] = {};
   	workbook.sheetArr[defaultsheetname].sheet = spreadsheet.sheet;
   	workbook.sheetArr[defaultsheetname].context = spreadsheet.context;
	
	// if these were properties of the sheet, then we wouldnt need to do this !
   	workbook.sheetArr[defaultsheetname].editorprop = {};
   	workbook.sheetArr[defaultsheetname].editorprop.ecell = null;
   	workbook.sheetArr[defaultsheetname].editorprop.range = null;
   	workbook.sheetArr[defaultsheetname].editorprop.range2 = null;
	
	workbook.clipsheet.savestr = null;
	workbook.clipsheet.copiedfrom = null;
	workbook.clipsheet.editorprop = {};

	spreadsheet.editor.workingvalues.currentsheet = spreadsheet.sheet.sheetname;
	spreadsheet.editor.workingvalues.startsheet = spreadsheet.editor.workingvalues.currentsheet;
        spreadsheet.editor.workingvalues.currentsheetid = spreadsheet.sheet.sheetid;
	
}

/**
 * Adds a new sheet to the workbook without switching to it
 * @param {Object} workbook - The workbook instance
 * @param {string} sheetid - The sheet ID
 * @param {string} sheetname - The sheet name
 * @param {string} savestr - Saved sheet data
 */
SocialCalc.AddNewWorkBookSheetNoSwitch = (workbook, sheetid, sheetname, savestr) => {

	//alert(`${sheetid},${sheetname},${savestr}`);
	
	const { spreadsheet } = workbook;
	
	const newsheet = new SocialCalc.Sheet();
	
	SocialCalc.Formula.SheetCache.sheets[sheetname] = {
		sheet: newsheet,
		name: sheetname
	};

	newsheet.sheetid = sheetid;
	newsheet.sheetname = sheetname;

        if (savestr) {
	    newsheet.ParseSheetSave(savestr);
	}
	
	workbook.sheetArr[sheetid] = {};
	workbook.sheetArr[sheetid].sheet = newsheet;
	workbook.sheetArr[sheetid].context = null;

	if (workbook.sheetArr[sheetid].sheet.attribs) {
	    workbook.sheetArr[sheetid].sheet.attribs.needsrecalc = "yes";
	}
	
	workbook.sheetArr[sheetid].editorprop = {};
	workbook.sheetArr[sheetid].editorprop.ecell = {
			coord: "A1",
			row: 1,
			col: 1
		};
	workbook.sheetArr[sheetid].editorprop.range = null;
	workbook.sheetArr[sheetid].editorprop.range2 = null;
	
}

/**
 * Adds a new sheet to the workbook and switches to it
 * @param {Object} workbook - The workbook instance
 * @param {string} sheetid - The new sheet ID
 * @param {string} oldsheetid - The old sheet ID
 * @param {boolean} fromclip - Whether the sheet is from clipboard
 * @param {Object} spread - Optional existing spreadsheet object
 */
SocialCalc.AddNewWorkBookSheet = (workbook, sheetid, oldsheetid, fromclip, spread) => {

	const { spreadsheet } = workbook;
	
	//alert(`create new sheet ${sheetid} old=${oldsheetid} def=${workbook.defaultsheetname}`);
	
	if (spread == null) {
		spreadsheet.sheet = new SocialCalc.Sheet();
		SocialCalc.Formula.SheetCache.sheets[sheetid] = {
			sheet: spreadsheet.sheet,
			name: sheetid
		}
	        spreadsheet.sheet.sheetid = sheetid;
		spreadsheet.sheet.sheetname = sheetid;
	}
	else {
		//alert("existing spread")
		spreadsheet.sheet = spread
	}

	spreadsheet.context = new SocialCalc.RenderContext(spreadsheet.sheet);
	
	spreadsheet.sheet.statuscallback = SocialCalc.EditorSheetStatusCallback;
	spreadsheet.sheet.statuscallbackparams = spreadsheet.editor;
	
	workbook.sheetArr[sheetid] = {};
	workbook.sheetArr[sheetid].sheet = spreadsheet.sheet;
	workbook.sheetArr[sheetid].context = spreadsheet.context;
	
	workbook.sheetArr[sheetid].editorprop = {};
	workbook.sheetArr[sheetid].editorprop.ecell = null;
	workbook.sheetArr[sheetid].editorprop.range = null;
	workbook.sheetArr[sheetid].editorprop.range2 = null;
	
	if (oldsheetid != null) {
		workbook.sheetArr[oldsheetid].editorprop.ecell = spreadsheet.editor.ecell;
		workbook.sheetArr[oldsheetid].editorprop.range = spreadsheet.editor.range;
		workbook.sheetArr[oldsheetid].editorprop.range2 = spreadsheet.editor.range2;
	}
	
	
	spreadsheet.context.showGrid = true;
	spreadsheet.context.showRCHeaders = true;
	spreadsheet.editor.context = spreadsheet.context;
	
	if (!fromclip) {
		spreadsheet.editor.ecell = {
			coord: "A1",
			row: 1,
			col: 1
		};
		
		spreadsheet.editor.range = {
			hasrange: false
		};
		spreadsheet.editor.range2 = {
			hasrange: false
		};
	}
	
	// set highlights
	spreadsheet.context.highlights[spreadsheet.editor.ecell.coord] = "cursor";
	
	if (fromclip) {
		// this is the result of a paste sheet
		//alert("from clip");
		
		if (workbook.clipsheet.savestr !== null) {
			//alert(`sheetdata = ${workbook.clipsheet.savestr}`);
			spreadsheet.sheet.ParseSheetSave(workbook.clipsheet.savestr);
		}
		
		spreadsheet.editor.ecell = workbook.clipsheet.editorprop.ecell;
		spreadsheet.context.highlights[spreadsheet.editor.ecell.coord] = "cursor";
		
		// range is not pasted ??!??
	
	}

	spreadsheet.editor.workingvalues.currentsheet = spreadsheet.sheet.sheetname;
	spreadsheet.editor.workingvalues.startsheet = spreadsheet.editor.workingvalues.currentsheet;
        spreadsheet.editor.workingvalues.currentsheetid = spreadsheet.sheet.sheetid;

	spreadsheet.editor.FitToEditTable();
	spreadsheet.editor.ScheduleRender();
	//spreadsheet.ExecuteCommand('recalc', '');
	
}

/**
 * Activates a specific sheet in the workbook
 * @param {Object} workbook - The workbook instance
 * @param {string} sheetnamestr - The sheet to activate
 * @param {string} oldsheetnamestr - The previously active sheet
 */
SocialCalc.ActivateWorkBookSheet = (workbook, sheetnamestr, oldsheetnamestr) => {

	const { spreadsheet } = workbook;
	
	//alert(`activate ${sheetnamestr} old=${oldsheetnamestr}`);
	
	spreadsheet.sheet = workbook.sheetArr[sheetnamestr].sheet;
	spreadsheet.context = workbook.sheetArr[sheetnamestr].context;

	if (spreadsheet.context == null) {
		//alert("context null")
		//for (const sheet in workbook.sheetArr) alert(`${sheet}${spreadsheet.sheet}`);
		workbook.AddNewWorkBookSheet(sheetnamestr, oldsheetnamestr, false, spreadsheet.sheet)
		return
	}

	spreadsheet.editor.context = spreadsheet.context;

	if (oldsheetnamestr != null) {
		workbook.sheetArr[oldsheetnamestr].editorprop.ecell = spreadsheet.editor.ecell;
	}
	spreadsheet.editor.ecell = workbook.sheetArr[sheetnamestr].editorprop.ecell;
	
	if (oldsheetnamestr != null) {
		workbook.sheetArr[oldsheetnamestr].editorprop.range = spreadsheet.editor.range;
	}
	spreadsheet.editor.range = workbook.sheetArr[sheetnamestr].editorprop.range;
			   
	if (oldsheetnamestr != null) {
		workbook.sheetArr[oldsheetnamestr].editorprop.range2 = spreadsheet.editor.range2;
	}
	spreadsheet.editor.range2 = workbook.sheetArr[sheetnamestr].editorprop.range2;

	spreadsheet.sheet.statuscallback = SocialCalc.EditorSheetStatusCallback;
    spreadsheet.sheet.statuscallbackparams = spreadsheet.editor;
			   	
	// reset highlights ??
	
	//spreadsheet.editor.FitToEditTable();				   

	spreadsheet.editor.workingvalues.currentsheet = spreadsheet.sheet.sheetname;
        spreadsheet.editor.workingvalues.currentsheetid = spreadsheet.sheet.sheetid;

	if (spreadsheet.editor.state !== "start" && spreadsheet.editor.inputBox) 
	  spreadsheet.editor.inputBox.element.focus();

	if (spreadsheet.editor.state === "start") {
	    spreadsheet.editor.workingvalues.startsheet = spreadsheet.editor.workingvalues.currentsheet;
	}
	
	//spreadsheet.editor.ScheduleRender();
	
        if (spreadsheet.editor.state !== "start" && spreadsheet.editor.inputBox) {
	    spreadsheet.editor.ScheduleRender();
	} else {
	    if (spreadsheet.sheet.attribs) {
	      spreadsheet.sheet.attribs.needsrecalc = "yes";
	    } else {
	      spreadsheet.sheet.attribs = {}
	      spreadsheet.sheet.attribs.needsrecalc = "yes";
	    }

	    spreadsheet.ExecuteCommand('redisplay','');
	} 

	
}   

/**
 * Deletes a sheet from the workbook
 * @param {Object} workbook - The workbook instance
 * @param {string} oldname - The sheet to delete
 * @param {string} curname - The current sheet name
 */
SocialCalc.DeleteWorkBookSheet = (workbook, oldname, curname) => {
	try {
		//alert(`delete ${oldname},${curname}`);
		
		if (!workbook.sheetArr[oldname]) {
			throw new Error(`Sheet ${oldname} not found`);
		}
		
		delete workbook.sheetArr[oldname].context;
		delete workbook.sheetArr[oldname].sheet;
		delete workbook.sheetArr[oldname];
		// take sheet out of the formula cache
		delete SocialCalc.Formula.SheetCache.sheets[curname];
	} catch (error) {
		console.error('Error in DeleteWorkBookSheet:', error);
		throw error;
	}
}


/**
 * Saves a workbook sheet and returns its data
 * @param {Object} workbook - The workbook instance
 * @param {string} sheetid - The sheet ID to save
 * @returns {Object} Object containing the saved sheet data
 */
SocialCalc.SaveWorkBookSheet = (workbook, sheetid) => {
	try {
		if (!workbook.sheetArr[sheetid]) {
			throw new Error(`Sheet with ID ${sheetid} not found`);
		}
		
		const sheetstr = {};
		sheetstr.savestr = workbook.sheetArr[sheetid].sheet.CreateSheetSave();
		return sheetstr;
	} catch (error) {
		console.error('Error in SaveWorkBookSheet:', error);
		throw error;
	}
} 

/**
 * Loads and renames a workbook sheet
 * @param {Object} workbook - The workbook instance
 * @param {string} sheetid - The sheet ID
 * @param {string} savestr - The saved sheet data
 * @param {string} newname - The new name for the sheet
 */
SocialCalc.LoadRenameWorkBookSheet = (workbook, sheetid, savestr, newname) => {
	try {
		if (!workbook.sheetArr[sheetid]) {
			throw new Error(`Sheet with ID ${sheetid} not found`);
		}
		
		workbook.sheetArr[sheetid].sheet.ResetSheet();
		workbook.sheetArr[sheetid].sheet.ParseSheetSave(savestr);

		if (workbook.sheetArr[sheetid].sheet.attribs) {
		    workbook.sheetArr[sheetid].sheet.attribs.needsrecalc = "yes";
		}
		
		delete SocialCalc.Formula.SheetCache.sheets[workbook.sheetArr[sheetid].sheet.sheetname];
		workbook.sheetArr[sheetid].sheet.sheetname = newname;
		SocialCalc.Formula.SheetCache.sheets[newname] = {sheet: workbook.sheetArr[sheetid].sheet, name: newname};
	} catch (error) {
		console.error('Error in LoadRenameWorkBookSheet:', error);
		throw error;
	}
}

/**
 * Renders the current workbook sheet
 * @param {Object} workbook - The workbook instance
 */
SocialCalc.RenderWorkBookSheet = (workbook) => {
	workbook.spreadsheet.editor.ScheduleRender();
};

/**
 * Renames sheet references in a formula cell
 * @param {string} formula - The formula to update
 * @param {string} oldname - The old sheet name
 * @param {string} newname - The new sheet name
 * @returns {string} The updated formula
 */
SocialCalc.RenameWorkBookSheetCell = (formula, oldname, newname) => {
	try {
	 	let ttype, ttext, i, newcr;
	   	let updatedformula = "";
	   	let sheetref = false;
	   	const scf = SocialCalc.Formula;
	   	if (!scf) {
	   		throw new Error("SocialCalc.Formula is required");
	    }
   	const { TokenType, TokenOpExpansion } = scf;
   	const { op: token_op, string: token_string, coord: token_coord } = TokenType;

   	const parseinfo = SocialCalc.Formula.ParseFormulaIntoTokens(formula);

   	for (i = 0; i < parseinfo.length; i++) {
   		const { type: ttype, text: ttext } = parseinfo[i];
		//alert(`${ttype},${ttext}`);
		//console.log(`${scf.NormalizeSheetName(ttext)}   ${oldname}`);
		if ((ttype === TokenType.name) && (scf.NormalizeSheetName(ttext) === oldname) && (i < parseinfo.length)) {
   			if ((parseinfo[i + 1].type === token_op) && (parseinfo[i + 1].text === "!")) {
				updatedformula += newname;//console.log (updatedformula);
			} else {
				updatedformula += ttext;//console.log (updatedformula);
			}
	  	} else {
			updatedformula += ttext;
		}
   	}
		//alert(updatedformula);
		return updatedformula;
	} catch (error) {
		console.error('Error in RenameWorkBookSheetCell:', error);
		return formula; // Return original formula if error occurs
	}
}

/**
 * Renames a workbook sheet and updates all formula references
 * @param {Object} workbook - The workbook instance
 * @param {string} oldname - The current name of the sheet
 * @param {string} newname - The new name for the sheet
 * @param {string} sheetid - The sheet ID
 */
SocialCalc.RenameWorkBookSheet = (workbook, oldname, newname, sheetid) => {

	// for each sheet, fix up all the formula references
	//
	//alert (sheetid);
	const oldsheet = SocialCalc.Formula.SheetCache.sheets[oldname].sheet;
	delete SocialCalc.Formula.SheetCache.sheets[oldname];
	//alert (newname); // to check the newname
	SocialCalc.Formula.SheetCache.sheets[newname] = {sheet: oldsheet, name: newname};
	workbook.sheetArr[sheetid].sheet.sheetname = newname
	//
	// fix up formulas for sheet rename
	// if formulas should not be fixed up upon sheet rename, then comment out the following
	// block
	//
	for (const sheet of Object.keys(workbook.sheetArr)) {
		//alert(`found sheet-${sheet}`);
		const { cells } = workbook.sheetArr[sheet].sheet;
		for (const cr of Object.keys(cells)) { // update cell references to sheet name
			//alert(cr);
			const cell = cells[cr];
			//if (cell) alert(cell.datatype)
			if (cell && cell.datatype === "f") {
				cell.formula = SocialCalc.RenameWorkBookSheetCell(cell.formula, oldname, newname);
				if (cell.parseinfo) {
					delete cell.parseinfo;
				}
			}
		}
	}
	// recalculate
	workbook.spreadsheet.ExecuteCommand('recalc', '');
}

/**
 * Copies a workbook sheet to the clipboard
 * @param {Object} workbook - The workbook instance
 * @param {string} sheetid - The sheet ID to copy
 */
SocialCalc.CopyWorkBookSheet = (workbook, sheetid) => {

	//alert(`in copy ${sheetid}`);
    workbook.clipsheet.savestr = workbook.sheetArr[sheetid].sheet.CreateSheetSave();
	//alert(`in copy save=${workbook.clipsheet.savestr}`);
    workbook.clipsheet.copiedfrom = sheetid;
    workbook.clipsheet.editorprop = {};
    workbook.clipsheet.editorprop.ecell = workbook.spreadsheet.editor.ecell;
    //workbook.clipsheet.editorprop.range = workbook.spreadsheet.editor.range;
	//workbook.clipsheet.editorprop.range2 = workbook.spreadsheet.editor.range2;
	//workbook.clipsheet.highlights = workbook.spreadsheet.context.highlights;
	
	//alert(`copied ${sheetid}`);
}

/**
 * Pastes a workbook sheet from the clipboard
 * @param {Object} workbook - The workbook instance
 * @param {string} newsheetid - The new sheet ID
 * @param {string} oldsheetid - The old sheet ID
 */
SocialCalc.PasteWorkBookSheet = (workbook, newsheetid, oldsheetid) => {
	
	//alert(`${newsheetid}${oldsheetid}`);
	workbook.AddNewWorkBookSheet(newsheetid, oldsheetid, true);
	
	// clear the clip ?
	
}


/**
 * Checks if a sheet name exists in the workbook
 * @param {Object} workbook - The workbook instance
 * @param {string} name - The sheet name to check
 * @returns {string|null} The sheet ID if exists, null otherwise
 */
SocialCalc.SheetNameExistsInWorkBook = (workbook, name) => {
    for (const sheet of Object.keys(workbook.sheetArr)) {    
	if (workbook.sheetArr[sheet].sheet.sheetname === name) {
	    return sheet;
	}
    }
    return null;
}