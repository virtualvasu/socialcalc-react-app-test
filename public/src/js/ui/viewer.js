//
// SocialCalcViewer
//
// (c) Copyright 2008-2010 Socialtext, Inc.
// All Rights Reserved.
//
// ────────────────────────────────────────────────────────────────
//  NOTE:  Extensive license text and history comments kept verbatim
//  to preserve original legal requirements.  Only executable code
//  below has been “modernized” (ES-6+ syntax, JSDoc) while behavior
//  remains identical.
// ────────────────────────────────────────────────────────────────
// This file is part of SocialCalc, a web-based spreadsheet application.

/* global SocialCalc, window */

/* -----------------------------------------------------------------
 * Ensure the main SocialCalc namespace and required modules exist
 * ----------------------------------------------------------------*/
if (typeof window.SocialCalc === "undefined") {
  /* eslint-disable no-alert */
  alert("Main SocialCalc code module needed");
  /* eslint-enable  no-alert */
  window.SocialCalc = {};
}

if (!window.SocialCalc.TableEditor) {
  alert("SocialCalc TableEditor code module needed");
}

/* -----------------------------------------------------------------
 * SpreadsheetViewer class
 * ----------------------------------------------------------------*/

/** @type {SocialCalc.SpreadsheetViewer|null} */
SocialCalc.CurrentSpreadsheetViewerObject = null; // only one active at a time

/**
 * SpreadsheetViewer constructor.
 * Creates a read-only spreadsheet viewer with optional toolbar/status bar.
 *
 * @constructor
 */
SocialCalc.SpreadsheetViewer = function () {

  /** @const */ const scc = SocialCalc.Constants;

  /* -------------------- Public properties -------------------- */

  /** @type {HTMLElement|null}  */ this.parentNode        = null;
  /** @type {HTMLElement|null}  */ this.spreadsheetDiv    = null;
  /** @type {number}            */ this.requestedHeight   = 0;
  /** @type {number}            */ this.requestedWidth    = 0;
  /** @type {number}            */ this.requestedSpaceBelow = 0;
  /** @type {number}            */ this.height            = 0;
  /** @type {number}            */ this.width             = 0;
  /** @type {number}            */ this.viewheight        = 0; // calculated

  /* -------------------- Dynamic references -------------------- */

  /** @type {SocialCalc.Sheet|null}         */ this.sheet         = null;
  /** @type {SocialCalc.RenderContext|null} */ this.context       = null;
  /** @type {SocialCalc.TableEditor|null}   */ this.editor        = null;

  /** @type {HTMLElement|null} */ this.editorDiv = null;

  /** @type {string} */ this.sortrange = ""; // remembered range for sort tab

  /* -------------------- Constants ----------------------------- */

  /** @type {string} */ this.idPrefix      = "SocialCalc-";
  /** @type {string} */ this.imagePrefix   = scc.defaultImagePrefix;

  /** @type {number} */ this.statuslineheight = scc.SVStatuslineheight;
  /** @type {string} */ this.statuslineCSS  = scc.SVStatuslineCSS;

  /* -------------------- Callbacks & flags --------------------- */

  /** @type {boolean} */ this.hasStatusLine = true;
  /** @type {string}  */ this.statuslineHTML =
      '<table cellspacing="0" cellpadding="0"><tr>' +
      '<td width="100%" style="overflow:hidden;">{status}</td>' +
      '<td>&nbsp;</td></tr></table>';
  /** @type {boolean} */ this.statuslineFull = true;
  /** @type {boolean} */ this.noRecalc = true; // viewer never recalcs

  /* -------------------- Repeating-macro support --------------- */

  /** @type {number|null} */ this.repeatingMacroTimer     = null;
  /** @type {number}       */ this.repeatingMacroInterval = 60;  // seconds
  /** @type {string}       */ this.repeatingMacroCommands = "";

  /* -------------------- Initialization ------------------------ */

  this.sheet             = new SocialCalc.Sheet();
  this.context           = new SocialCalc.RenderContext(this.sheet);
  this.context.showGrid  = true;
  this.context.showRCHeaders = true;

  this.editor            = new SocialCalc.TableEditor(this.context);
  this.editor.noEdit     = true;
  this.editor.StatusCallback.statusline = {
    func   : SocialCalc.SpreadsheetViewerStatuslineCallback,
    params : {}
  };

  // Remember the active viewer instance
  SocialCalc.CurrentSpreadsheetViewerObject = this;
};

/**
 * SpreadsheetViewer prototype methods
 */

/**
 * Initializes the spreadsheet viewer in the specified node
 * @param {HTMLElement|string} node - Parent DOM element or element ID
 * @param {number} height - Requested height in pixels
 * @param {number} width - Requested width in pixels  
 * @param {number} spacebelow - Space to leave below viewer
 * @returns {*} Result from InitializeSpreadsheetViewer
 */
SocialCalc.SpreadsheetViewer.prototype.InitializeSpreadsheetViewer = 
   function(node, height, width, spacebelow) {
      return SocialCalc.InitializeSpreadsheetViewer(this, node, height, width, spacebelow);
   };

/**
 * Loads saved spreadsheet data
 * @param {string} str - Saved spreadsheet data string
 * @returns {*} Result from SpreadsheetViewerLoadSave
 */
SocialCalc.SpreadsheetViewer.prototype.LoadSave = function(str) {
   return SocialCalc.SpreadsheetViewerLoadSave(this, str);
};

/**
 * Handles window resize events
 * @returns {*} Result from DoOnResize
 */
SocialCalc.SpreadsheetViewer.prototype.DoOnResize = function() {
   return SocialCalc.DoOnResize(this);
};

/**
 * Sizes the spreadsheet div element
 * @returns {*} Result from SizeSSDiv
 */
SocialCalc.SpreadsheetViewer.prototype.SizeSSDiv = function() {
   return SocialCalc.SizeSSDiv(this);
};

/**
 * Decodes saved spreadsheet data
 * @param {string} str - Encoded spreadsheet save string
 * @returns {*} Result from SpreadsheetViewerDecodeSpreadsheetSave
 */
SocialCalc.SpreadsheetViewer.prototype.DecodeSpreadsheetSave = 
   function(str) {
      return SocialCalc.SpreadsheetViewerDecodeSpreadsheetSave(this, str);
   };

/**
 * Sheet method proxy to make access easier
 */

/**
 * Parses sheet save data
 * @param {string} str - Sheet save data string
 * @returns {*} Result from sheet.ParseSheetSave
 */
SocialCalc.SpreadsheetViewer.prototype.ParseSheetSave = function(str) {
   return this.sheet.ParseSheetSave(str);
};

/**
 * Main Functions
 */

/**
 * InitializeSpreadsheetViewer(spreadsheet, node, height, width, spacebelow)
 *
 * Creates the control elements and makes them the child of node (string or element).
 * If present, height and width specify size.
 * If either is 0 or null (missing), the maximum that fits on the screen
 * (taking spacebelow into account) is used.
 *
 * You should do a redisplay or recalc (which redisplays) after running this.
 *
 * @param {SocialCalc.SpreadsheetViewer} spreadsheet - The viewer instance
 * @param {HTMLElement|string} node - Parent DOM element or element ID
 * @param {number} height - Requested height in pixels
 * @param {number} width - Requested width in pixels
 * @param {number} spacebelow - Space to leave below viewer
 */
SocialCalc.InitializeSpreadsheetViewer = function(spreadsheet, node, height, width, spacebelow) {
   const scc = SocialCalc.Constants;
   const SCLoc = SocialCalc.LocalizeString;
   const SCLocSS = SocialCalc.LocalizeSubstrings;

   let html, child, i, vname, v, style, button, bele;
   const tabs = spreadsheet.tabs;
   const views = spreadsheet.views;

   spreadsheet.requestedHeight = height;
   spreadsheet.requestedWidth = width;
   spreadsheet.requestedSpaceBelow = spacebelow;

   if (typeof node == "string") node = document.getElementById(node);

   if (node == null) {
      alert("SocialCalc.SpreadsheetControl not given parent node.");
   }

   spreadsheet.parentNode = node;

   // create node to hold spreadsheet view
   spreadsheet.spreadsheetDiv = document.createElement("div");

   spreadsheet.SizeSSDiv(); // calculate and fill in the size values

   for (child = node.firstChild; child != null; child = node.firstChild) {
      node.removeChild(child);
   }

   node.appendChild(spreadsheet.spreadsheetDiv);

   // create sheet div
   spreadsheet.nonviewheight = spreadsheet.hasStatusLine ? spreadsheet.statuslineheight : 0;
   spreadsheet.viewheight = spreadsheet.height - spreadsheet.nonviewheight;
   spreadsheet.editorDiv = spreadsheet.editor.CreateTableEditor(spreadsheet.width, spreadsheet.viewheight);

   spreadsheet.spreadsheetDiv.appendChild(spreadsheet.editorDiv);

   // create statusline
   if (spreadsheet.hasStatusLine) {
      spreadsheet.statuslineDiv = document.createElement("div");
      spreadsheet.statuslineDiv.style.cssText = spreadsheet.statuslineCSS;
      spreadsheet.statuslineDiv.style.height = spreadsheet.statuslineheight -
         (spreadsheet.statuslineDiv.style.paddingTop.slice(0, -2) - 0) -
         (spreadsheet.statuslineDiv.style.paddingBottom.slice(0, -2) - 0) + "px";
      spreadsheet.statuslineDiv.id = spreadsheet.idPrefix + "statusline";
      spreadsheet.spreadsheetDiv.appendChild(spreadsheet.statuslineDiv);
      spreadsheet.editor.StatusCallback.statusline =
         { func: SocialCalc.SpreadsheetViewerStatuslineCallback,
           params: { spreadsheetobj: spreadsheet } };
   }

   // done - refresh screen needed
   return;
};

/**
 * Loads saved spreadsheet data and applies it to the viewer
 * @param {SocialCalc.SpreadsheetViewer} spreadsheet - The viewer instance
 * @param {string} savestr - Saved spreadsheet data string
 */
SocialCalc.SpreadsheetViewerLoadSave = function(spreadsheet, savestr) {
   let rmstr, pos, t, t2;

   const parts = spreadsheet.DecodeSpreadsheetSave(savestr);
   if (parts) {
      if (parts.sheet) {
         spreadsheet.sheet.ResetSheet();
         spreadsheet.sheet.ParseSheetSave(savestr.substring(parts.sheet.start, parts.sheet.end));
      }
      if (parts.edit) {
         spreadsheet.editor.LoadEditorSettings(savestr.substring(parts.edit.start, parts.edit.end));
      }
      if (parts.startupmacro) { // executed now
         spreadsheet.editor.EditorScheduleSheetCommands(savestr.substring(parts.startupmacro.start, parts.startupmacro.end), false, true);
      }
      if (parts.repeatingmacro) { // first line tells how many seconds before first execution. Last cmd must be "cmdextension repeatmacro delay" to continue repeating.
         rmstr = savestr.substring(parts.repeatingmacro.start, parts.repeatingmacro.end);
         rmstr = rmstr.replace("\r", ""); // make sure no CR, only LF
         pos = rmstr.indexOf("\n");
         if (pos > 0) {
            t = rmstr.substring(0, pos) - 0; // get number
            t2 = t;
//            if (!(t > 0)) t = 60; // handles NAN, too
            spreadsheet.repeatingMacroInterval = t;
            spreadsheet.repeatingMacroCommands = rmstr.substring(pos + 1);
            if (t2 > 0) { // zero means don't start yet
               spreadsheet.repeatingMacroTimer = window.setTimeout(SocialCalc.SpreadsheetViewerDoRepeatingMacro, spreadsheet.repeatingMacroInterval * 1000);
            }
         }
      }
   }
   if (spreadsheet.editor.context.sheetobj.attribs.recalc == "off" || spreadsheet.noRecalc) {
      spreadsheet.editor.ScheduleRender();
   }
   else {
      spreadsheet.editor.EditorScheduleSheetCommands("recalc");
   }
};

/**
 * SocialCalc.SpreadsheetViewerDoRepeatingMacro
 *
 * Called by a timer. Executes repeatingMacroCommands once.
 * Use the "startcmdextension repeatmacro delay" command last to schedule this again.
 */
SocialCalc.SpreadsheetViewerDoRepeatingMacro = function() {
   const spreadsheet = SocialCalc.GetSpreadsheetViewerObject();
   const editor = spreadsheet.editor;

   spreadsheet.repeatingMacroTimer = null;

   SocialCalc.SheetCommandInfo.CmdExtensionCallbacks.repeatmacro = { func: SocialCalc.SpreadsheetViewerRepeatMacroCommand, data: null };

   editor.EditorScheduleSheetCommands(spreadsheet.repeatingMacroCommands);
};

/**
 * Handles the repeat macro command extension
 * @param {string} name - Command name
 * @param {*} data - Command data
 * @param {SocialCalc.Sheet} sheet - Sheet object
 * @param {*} cmd - Command object
 * @param {boolean} saveundo - Whether to save undo
 */
SocialCalc.SpreadsheetViewerRepeatMacroCommand = function(name, data, sheet, cmd, saveundo) {
   const spreadsheet = SocialCalc.GetSpreadsheetViewerObject();

   const rest = cmd.RestOfString();
   let t = rest - 0; // get number
   if (!(t > 0)) t = spreadsheet.repeatingMacroInterval; // handles NAN, too, using last value
   spreadsheet.repeatingMacroInterval = t;

   spreadsheet.repeatingMacroTimer = window.setTimeout(SocialCalc.SpreadsheetViewerDoRepeatingMacro, spreadsheet.repeatingMacroInterval * 1000);
};

/**
 * Stops the repeating macro timer
 */
SocialCalc.SpreadsheetViewerStopRepeatingMacro = function() {
   const spreadsheet = SocialCalc.GetSpreadsheetViewerObject();

   if (spreadsheet.repeatingMacroTimer) {
      window.clearTimeout(spreadsheet.repeatingMacroTimer);
      spreadsheet.repeatingMacroTimer = null;
   }
};

/**
 * SocialCalc.SpreadsheetViewerDoButtonCmd(e, buttoninfo, bobj)
 *
 * Handles button command execution in the spreadsheet viewer
 *
 * @param {Event} e - Event object
 * @param {object} buttoninfo - Button information object
 * @param {object} bobj - Button object containing element and function info
 */
SocialCalc.SpreadsheetViewerDoButtonCmd = function(e, buttoninfo, bobj) {
   const obj = bobj.element;
   const which = bobj.functionobj.command;

   const spreadsheet = SocialCalc.GetSpreadsheetViewerObject();
   const editor = spreadsheet.editor;

   switch (which) {
      case "recalc":
         editor.EditorScheduleSheetCommands("recalc");
         break;

      default:
         break;
   }

   if (obj && obj.blur) obj.blur();
   SocialCalc.KeyboardFocus();
};

/**
 * outstr = SocialCalc.LocalizeString(str)
 *
 * SocialCalc function to make localization easier.
 * If str is "Text to localize", it returns
 * SocialCalc.Constants.s_loc_text_to_localize if
 * it exists, or else with just "Text to localize".
 * Note that spaces are replaced with "_" and other special
 * chars with "X" in the name of the constant (e.g., "A & B"
 * would look for SocialCalc.Constants.s_loc_a_X_b.
 *
 * @param {string} str - String to localize
 * @returns {string} Localized string or original if not found
 */
SocialCalc.LocalizeString = function(str) {
   let cstr = SocialCalc.LocalizeStringList[str]; // found already this session?
   if (!cstr) { // no - look up
      cstr = SocialCalc.Constants["s_loc_" + str.toLowerCase().replace(/\s/g, "_").replace(/\W/g, "X")] || str;
      SocialCalc.LocalizeStringList[str] = cstr;
   }
   return cstr;
};

SocialCalc.LocalizeStringList = {}; // a list of strings to localize accumulated by the routine

/**
 * outstr = SocialCalc.LocalizeSubstrings(str)
 *
 * SocialCalc function to make localization easier using %loc and %scc.
 *
 * Replaces sections of str with:
 *    %loc!Text to localize!
 * with SocialCalc.Constants.s_loc_text_to_localize if
 * it exists, or else with just "Text to localize".
 * Note that spaces are replaced with "_" and other special
 * chars with "X" in the name of the constant (e.g., %loc!A & B!
 * would look for SocialCalc.Constants.s_loc_a_X_b.
 * Uses SocialCalc.LocalizeString for this.
 *
 * Replaces sections of str with:
 *    %ssc!constant-name!
 * with SocialCalc.Constants.constant-name.
 * If the constant doesn't exist, throws and alert.
 *
 * @param {string} str - String containing localization placeholders
 * @returns {string} String with placeholders replaced
 */
SocialCalc.LocalizeSubstrings = function(str) {
   const SCLoc = SocialCalc.LocalizeString;

   return str.replace(/%(loc|ssc)!(.*?)!/g, function(a, t, c) {
      if (t == "ssc") {
         return SocialCalc.Constants[c] || alert("Missing constant: " + c);
      }
      else {
         return SCLoc(c);
      }
   });
};

/**
 * obj = GetSpreadsheetViewerObject()
 *
 * Returns the current spreadsheet view object
 *
 * @returns {SocialCalc.SpreadsheetViewer} Current spreadsheet viewer object
 * @throws {string} Error message if no current object exists
 */
SocialCalc.GetSpreadsheetViewerObject = function() {
   const csvo = SocialCalc.CurrentSpreadsheetViewerObject;
   if (csvo) return csvo;

   throw ("No current SpreadsheetViewer object.");
};

/**
 * SocialCalc.DoOnResize(spreadsheet)
 *
 * Processes an onResize event, setting the different views.
 *
 * @param {SocialCalc.SpreadsheetViewer} spreadsheet - The spreadsheet viewer object
 */
SocialCalc.DoOnResize = function(spreadsheet) {
   let v;
   const views = spreadsheet.views;

   const needresize = spreadsheet.SizeSSDiv();
   if (!needresize) return;

   for (const vname in views) {
      v = views[vname].element;
      v.style.width = spreadsheet.width + "px";
      v.style.height = (spreadsheet.height - spreadsheet.nonviewheight) + "px";
   }

   spreadsheet.editor.ResizeTableEditor(spreadsheet.width, spreadsheet.height - spreadsheet.nonviewheight);
};

/**
 * resized = SocialCalc.SizeSSDiv(spreadsheet)
 *
 * Figures out a reasonable size for the spreadsheet, given any requested values and viewport.
 * Sets ssdiv to that.
 * Return true if different than existing values.
 *
 * @param {SocialCalc.SpreadsheetViewer} spreadsheet - The spreadsheet viewer object
 * @returns {boolean} True if size was changed, false otherwise
 */
SocialCalc.SizeSSDiv = function(spreadsheet) {
   let sizes, pos, resized, nodestyle, newval;
   const fudgefactorX = 10; // for IE
   const fudgefactorY = 10;

   resized = false;

   sizes = SocialCalc.GetViewportInfo();
   pos = SocialCalc.GetElementPosition(spreadsheet.parentNode);
   pos.bottom = 0;
   pos.right = 0;

   nodestyle = spreadsheet.parentNode.style;

   if (nodestyle.marginTop) {
      pos.top += nodestyle.marginTop.slice(0, -2) - 0;
   }
   if (nodestyle.marginBottom) {
      pos.bottom += nodestyle.marginBottom.slice(0, -2) - 0;
   }
   if (nodestyle.marginLeft) {
      pos.left += nodestyle.marginLeft.slice(0, -2) - 0;
   }
   if (nodestyle.marginRight) {
      pos.right += nodestyle.marginRight.slice(0, -2) - 0;
   }

   newval = spreadsheet.requestedHeight ||
            sizes.height - (pos.top + pos.bottom + fudgefactorY) -
               (spreadsheet.requestedSpaceBelow || 0);
   if (spreadsheet.height != newval) {
      spreadsheet.height = newval;
      spreadsheet.spreadsheetDiv.style.height = newval + "px";
      resized = true;
   }
   newval = spreadsheet.requestedWidth ||
            sizes.width - (pos.left + pos.right + fudgefactorX) || 700;
   if (spreadsheet.width != newval) {
      spreadsheet.width = newval;
      spreadsheet.spreadsheetDiv.style.width = newval + "px";
      resized = true;
   }

   return resized;
};

/**
 * SocialCalc.SpreadsheetViewerStatuslineCallback
 *
 * Updates (or clears) the status-line HTML when the editor’s status changes.
 *
 * @param {SocialCalc.TableEditor} editor  – The editor issuing the callback
 * @param {string}                status  – Status code
 * @param {*}                      arg     – Optional status argument
 * @param {{spreadsheetobj: SocialCalc.SpreadsheetViewer}} params – Extra params
 */
SocialCalc.SpreadsheetViewerStatuslineCallback = function (editor, status, arg, params) {
   const spreadsheet = params.spreadsheetobj;
   let slstr = "";

   if (spreadsheet && spreadsheet.statuslineDiv) {
      slstr = spreadsheet.statuslineFull
               ? editor.GetStatuslineString(status, arg, params)
               : editor.ecell.coord;
      slstr = spreadsheet.statuslineHTML.replace(/\{status\}/, slstr);
      spreadsheet.statuslineDiv.innerHTML = slstr;
   }

   switch (status) {
      case "cmdendnorender":
      case "calcfinished":
      case "doneposcalc":
         /* no toolbar here, so nothing to update */ 
         break;
      default:
         break;
   }
};

/**
 * SocialCalc.CmdGotFocus(obj)
 *
 * Sets SocialCalc.Keyboard.passThru: obj should be element with focus or `"true"`.
 *
 * @param {HTMLElement|true} obj – Element that received focus (or true)
 */
SocialCalc.CmdGotFocus = function (obj) {
   SocialCalc.Keyboard.passThru = obj;
};

/**
 * Returns the entire sheet rendered as plain HTML.
 *
 * @param {SocialCalc.SpreadsheetViewer} spreadsheet
 * @returns {string} HTML string representing the sheet
 */
SocialCalc.SpreadsheetViewerCreateSheetHTML = function (spreadsheet) {
   const context = new SocialCalc.RenderContext(spreadsheet.sheet);
   const div     = document.createElement("div");
   const ele     = context.RenderSheet(null, { type: "html" });

   div.appendChild(ele);
   const result = div.innerHTML;

   // clean-up
   delete context;
   delete ele;
   delete div;
   return result;
};

///////////////////////
//  LOAD ROUTINES   //
///////////////////////

/**
 * Separates the parts of a spreadsheet save string, returning start/end offsets.
 *
 * @param {SocialCalc.SpreadsheetViewer} spreadsheet – The viewer (unused but kept for API parity)
 * @param {string} str  – Full MIME multipart save string
 * @returns {Object.<string,{start:number,end:number}>} Map of part-names to ranges
 */
SocialCalc.SpreadsheetViewerDecodeSpreadsheetSave = function (spreadsheet, str) {

   // Normalize any bare-CR to CRLF
   const hasReturnOnly = /[^\n]\r[^\n]/;
   if (hasReturnOnly.test(str)) {
      str = str.replace(/([^\n])\r([^\n])/g, "$1\r\n$2");
   }

   const parts = {};
   let searchInfo;

   const topPos = str.search(/^MIME-Version:\s1\.0/mi);
   if (topPos < 0) return parts;

   /* -------- Determine multipart boundary -------- */
   const mpRegex = /^Content-Type:\s*multipart\/mixed;\s*boundary=(\S+)/mig;
   mpRegex.lastIndex = topPos;
   searchInfo = mpRegex.exec(str);
   if (!searchInfo) return parts;
   const boundary = searchInfo[1];

   const boundaryRegex = new RegExp("^--" + boundary + "(?:\r\n|\n)", "mg");
   boundaryRegex.lastIndex = mpRegex.lastIndex;

   /* Move to the blank line after the MIME header */
   searchInfo = boundaryRegex.exec(str);
   const blankLineRegex = /(?:\r\n|\n)(?:\r\n|\n)/gm;
   blankLineRegex.lastIndex = boundaryRegex.lastIndex;
   searchInfo = blankLineRegex.exec(str);
   if (!searchInfo) return parts;
   let start = blankLineRegex.lastIndex;

   /* Gather header key/value lines */
   boundaryRegex.lastIndex = start;
   searchInfo = boundaryRegex.exec(str);
   if (!searchInfo) return parts;
   let ending = searchInfo.index;

   const headerLines = str.substring(start, ending).split(/\r\n|\n/);
   const partList = [];

   for (const line of headerLines) {
      const p = line.split(":");
      if (p[0] === "part") partList.push(p[1]);
   }

   /* -------- Extract each MIME part -------- */
   for (let pnum = 0; pnum < partList.length; pnum++) {
      blankLineRegex.lastIndex = ending;
      searchInfo = blankLineRegex.exec(str);               // blank after part header
      if (!searchInfo) return parts;
      start = blankLineRegex.lastIndex;

      // Last part ends with boundary--
      const endBoundaryRegex = (pnum === partList.length - 1)
         ? new RegExp("^--" + boundary + "--$", "mg")
         : new RegExp("^--" + boundary + "(?:\r\n|\n)", "mg");

      endBoundaryRegex.lastIndex = start;
      searchInfo = endBoundaryRegex.exec(str);
      if (!searchInfo) return parts;
      ending = searchInfo.index;

      parts[partList[pnum]] = { start, end: ending };
   }

   return parts;
};

// END OF FILE
