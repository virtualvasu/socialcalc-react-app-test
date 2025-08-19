/**
 * @fileoverview SocialCalc Table Editor
 * The code module of the SocialCalc package that displays a scrolling grid with panes
 * and handles keyboard and mouse I/O.
 * 
 * @version 1.0.0
 * @copyright Copyright 2008, 2009, 2010 Socialtext, Inc. All Rights Reserved.
 * @license Common Public Attribution License Version 1.0
 * @see {@link http://socialcalc.org} For license details
 * 
 * @description
 * LEGAL NOTICES REQUIRED BY THE COMMON PUBLIC ATTRIBUTION LICENSE:
 * 
 * EXHIBIT A. Common Public Attribution License Version 1.0.
 * 
 * The contents of this file are subject to the Common Public Attribution License Version 1.0 (the 
 * "License"); you may not use this file except in compliance with the License. You may obtain a copy 
 * of the License at http://socialcalc.org. The License is based on the Mozilla Public License Version 1.1 but 
 * Sections 14 and 15 have been added to cover use of software over a computer network and provide for 
 * limited attribution for the Original Developer. In addition, Exhibit A has been modified to be 
 * consistent with Exhibit B.
 * 
 * Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF ANY 
 * KIND, either express or implied. See the License for the specific language governing rights and 
 * limitations under the License.
 * 
 * The Original Code is SocialCalc JavaScript TableEditor.
 * The Original Developer is the Initial Developer.
 * The Initial Developer of the Original Code is Socialtext, Inc. All portions of the code written by 
 * Socialtext, Inc., are Copyright (c) Socialtext, Inc. All Rights Reserved.
 * 
 * Contributor: Dan Bricklin.
 * 
 * EXHIBIT B. Attribution Information
 * 
 * When the TableEditor is producing and/or controlling the display the Graphic Image must be
 * displayed on the screen visible to the user in a manner comparable to that in the 
 * Original Code. The Attribution Phrase must be displayed as a "tooltip" or "hover-text" for
 * that image. The image must be linked to the Attribution URL so as to access that page
 * when clicked. If the user interface includes a prominent "about" display which includes
 * factual prominent attribution in a form similar to that in the "about" display included
 * with the Original Code, including Socialtext copyright notices and URLs, then the image
 * need not be linked to the Attribution URL but the "tool-tip" is still required.
 * 
 * Attribution Copyright Notice:
 * Copyright (C) 2010 Socialtext, Inc. All Rights Reserved.
 * 
 * Attribution Phrase (not exceeding 10 words): SocialCalc
 * Attribution URL: http://www.socialcalc.org/xoattrib
 * 
 * Graphic Image: The contents of the sc-logo.gif file in the Original Code or
 * a suitable replacement from http://www.socialcalc.org/licenses specified as
 * being for SocialCalc.
 * 
 * Display of Attribution Information is required in Larger Works which are defined 
 * in the CPAL as a work which combines Covered Code or portions thereof with code 
 * not governed by the terms of the CPAL.
 * 
 * Some of the other files in the SocialCalc package are licensed under
 * different licenses. Please note the licenses of the modules you use.
 * 
 * Code History:
 * Initially coded by Dan Bricklin of Software Garden, Inc., for Socialtext, Inc.
 * Based in part on the SocialCalc 1.1.0 code written in Perl.
 * The SocialCalc 1.1.0 code was:
 *    Portions (c) Copyright 2005, 2006, 2007 Software Garden, Inc. All Rights Reserved.
 *    Portions (c) Copyright 2007 Socialtext, Inc. All Rights Reserved.
 * The Perl SocialCalc started as modifications to the wikiCalc(R) program, version 1.0.
 * wikiCalc 1.0 was written by Software Garden, Inc.
 * Unless otherwise specified, referring to "SocialCalc" in comments refers to this
 * JavaScript version of the code, not the SocialCalc Perl code.
 * 
 * See the comments in the main SocialCalc code module file of the SocialCalc package.
 */

/**
 * @namespace SocialCalc
 * @description The main SocialCalc namespace
 */
var SocialCalc;
if (!SocialCalc) { // created here, too, in case load order is wrong, but main routines are required
   SocialCalc = {};
}

/**
 * @class TableEditor
 * @memberof SocialCalc
 * @description Table Editor class for displaying a scrolling grid with panes and handling keyboard/mouse I/O
 * @param {Object} context - The editing context
 */
SocialCalc.TableEditor = function(context) {
   const scc = SocialCalc.Constants;

   /**
    * @member {Object} context - Editing context
    */
   this.context = context;
   
   /**
    * @member {HTMLElement|null} toplevel - Top level HTML element for this table editor
    */
   this.toplevel = null;
   
   /**
    * @member {Object|null} fullgrid - Rendered editing context
    */
   this.fullgrid = null;

   /**
    * @member {boolean} noEdit - If true, disable all edit UI and make read-only
    */
   this.noEdit = false;

   /**
    * @member {number|null} width - Editor width
    */
   this.width = null;
   
   /**
    * @member {number|null} tablewidth - Table width
    */
   this.tablewidth = null;
   
   /**
    * @member {number|null} height - Editor height
    */
   this.height = null;
   
   /**
    * @member {number|null} tableheight - Table height
    */
   this.tableheight = null;

   /**
    * @member {HTMLElement|null} inputBox - Input box element
    */
   this.inputBox = null;
   
   /**
    * @member {HTMLElement|null} inputEcho - Input echo element
    */
   this.inputEcho = null;
   
   /**
    * @member {HTMLElement|null} verticaltablecontrol - Vertical table control element
    */
   this.verticaltablecontrol = null;
   
   /**
    * @member {HTMLElement|null} horizontaltablecontrol - Horizontal table control element
    */
   this.horizontaltablecontrol = null;

   /**
    * @member {HTMLElement|null} logo - Logo element
    */
   this.logo = null;

   /**
    * @member {Object|null} cellhandles - Cell handles object
    */
   this.cellhandles = null;

   // Dynamic properties:

   /**
    * @member {number|null} timeout - Timer id for position calculations (if non-null)
    */
   this.timeout = null;
   
   /**
    * @member {boolean} busy - True when executing command, calculating, etc.
    */
   this.busy = false;
   
   /**
    * @member {boolean} ensureecell - If true, ensure ecell is visible after timeout
    */
   this.ensureecell = false;
   
   /**
    * @member {Array<Object>} deferredCommands - Commands to execute after busy, in form: {cmdstr: "cmds", saveundo: t/f}
    */
   this.deferredCommands = [];

   /**
    * @member {Object|null} gridposition - Screen coords of full grid
    */
   this.gridposition = null;
   
   /**
    * @member {Object|null} headposition - Screen coords of upper left of grid within header rows
    */
   this.headposition = null;
   
   /**
    * @member {number|null} firstscrollingrow - Row number of top row in last (the scrolling) pane
    */
   this.firstscrollingrow = null;
   
   /**
    * @member {number|null} firstscrollingrowtop - Position of top row in last (the scrolling) pane
    */
   this.firstscrollingrowtop = null;
   
   /**
    * @member {number|null} lastnonscrollingrow - Row number of last displayed row in last non-scrolling pane, or zero (for thumb position calculations)
    */
   this.lastnonscrollingrow = null;
   
   /**
    * @member {number|null} lastvisiblerow - Used for paging down
    */
   this.lastvisiblerow = null;
   
   /**
    * @member {number|null} firstscrollingcol - Column number of top col in last (the scrolling) pane
    */
   this.firstscrollingcol = null;
   
   /**
    * @member {number|null} firstscrollingcolleft - Position of top col in last (the scrolling) pane
    */
   this.firstscrollingcolleft = null;
   
   /**
    * @member {number|null} lastnonscrollingcol - Col number of last displayed column in last non-scrolling pane, or zero (for thumb position calculations)
    */
   this.lastnonscrollingcol = null;
   
   /**
    * @member {number|null} lastvisiblecol - Used for paging right
    */
   this.lastvisiblecol = null;

   /**
    * @member {Array<number>} rowpositions - Screen positions of the top of some rows
    */
   this.rowpositions = [];
   
   /**
    * @member {Array<number>} colpositions - Screen positions of the left side of some rows
    */
   this.colpositions = [];
   
   /**
    * @member {Array<number>} rowheight - Size in pixels of each row when last checked, or null/undefined, for page up
    */
   this.rowheight = [];
   
   /**
    * @member {Array<number>} colwidth - Size in pixels of each column when last checked, or null/undefined, for page left
    */
   this.colwidth = [];

   /**
    * @member {Object|null} ecell - Either null or {coord: c, row: r, col: c}
    */
   this.ecell = null;
   
   /**
    * @member {string} state - The keyboard states: see EditorProcessKey
    */
   this.state = "start";

   /**
    * @member {Object} workingvalues - Values used during keyboard editing, etc.
    */
   this.workingvalues = {};

   // Constants:

   /**
    * @member {string} imageprefix - URL prefix for images (e.g., "/images/sc")
    */
   this.imageprefix = scc.defaultImagePrefix;
   
   /**
    * @member {string} idPrefix - ID prefix for elements
    */
   this.idPrefix = scc.defaultTableEditorIDPrefix;
   
   /**
    * @member {number} pageUpDnAmount - Number of rows to move cursor on PgUp/PgDn keys (numeric)
    */
   this.pageUpDnAmount = scc.defaultPageUpDnAmount;

   // Callbacks

   /**
    * @member {Function} recalcFunction - If present, function(editor) {...}, called to do a recalc
    * @description Default (sheet.RecalcSheet) does all the right stuff.
    * @param {Object} editor - The editor instance
    * @returns {*} Result of recalculation or null
    */
   this.recalcFunction = (editor) => {
      if (editor.context.sheetobj.RecalcSheet) {
         editor.context.sheetobj.RecalcSheet(SocialCalc.EditorSheetStatusCallback, editor);
      } else {
         return null;
      }
   };

   /**
    * @member {Function} ctrlkeyFunction - If present, function(editor, charname) {...}, called to handle ctrl-V, etc., at top level
    * @description Returns true (pass through for continued processing) or false (stop processing this key).
    * @param {Object} editor - The editor instance
    * @param {string} charname - The character name
    * @returns {boolean} True for continued processing, false to stop processing this key
    */
   this.ctrlkeyFunction = (editor, charname) => {
      let ta, ha, cell, position, cmd, sel, cliptext;

      switch (charname) {
         case "[ctrl-c]":
         case "[ctrl-x]":
            ta = editor.pasteTextarea;
            ta.value = "";
            cell = SocialCalc.GetEditorCellElement(editor, editor.ecell.row, editor.ecell.col);
            if (cell) {
               position = SocialCalc.GetElementPosition(cell.element);
               ta.style.left = `${position.left - 1}px`;
               ta.style.top = `${position.top - 1}px`;
            }
            if (editor.range.hasrange) {
               sel = `${SocialCalc.crToCoord(editor.range.left, editor.range.top)}:${SocialCalc.crToCoord(editor.range.right, editor.range.bottom)}`;
            } else {
               sel = editor.ecell.coord;
            }

            // get what to copy to clipboard
            cliptext = SocialCalc.ConvertSaveToOtherFormat(SocialCalc.CreateSheetSave(editor.context.sheetobj, sel), "tab");

            if (charname === "[ctrl-c]" || editor.noEdit || (SocialCalc.Callbacks.IsCellEditable && (!SocialCalc.Callbacks.IsCellEditable(editor)))) { // if copy or cut but in no edit
               cmd = `copy ${sel} formulas`;
            } else { // [ctrl-x]
               cmd = `cut ${sel} formulas`;
            }
            editor.EditorScheduleSheetCommands(cmd, true, false); // queue up command to put on SocialCalc clipboard

            /* Copy as HTML: This fails rather badly as it won't paste into Notepad as tab-delimited text. Oh well.

                ha = editor.pasteHTMLarea;
                if (editor.range.hasrange) {
                    cell = SocialCalc.GetEditorCellElement(editor, editor.range.top, editor.range.left);
                }
                else {
                    cell = SocialCalc.GetEditorCellElement(editor, editor.ecell.row, editor.ecell.col);
                }
                if (cell) position = SocialCalc.GetElementPosition(cell.element);
             
                if (ha) {
                    if (position) {
                        ha.style.left = (position.left-1)+"px";
                        ha.style.top = (position.top-1)+"px";
                    }
                    ha.style.visibility="visible";
                    cliptext = SocialCalc.ConvertSaveToOtherFormat(SocialCalc.CreateSheetSave(editor.context.sheetobj, sel), "html");
                    ha.innerHTML = cliptext.replace(/<tr\b[^>]*>[\d\D]*?<\/tr\b[^>]*>/i, '');
                    ha.focus();

                    var range = document.body.createControlRange();
                    range.addElement(ha.childNodes[0]);
                    range.select();
                }
            */
            ta.style.display = "block";
            ta.value = cliptext; // must follow "block" setting for Webkit
            ta.focus();
            ta.select();
            window.setTimeout(() => {
               if (!SocialCalc.GetSpreadsheetControlObject) return; // in case not loaded
               const s = SocialCalc.GetSpreadsheetControlObject();
               if (!s) return;
               const editor = s.editor;
               /*
               var ha = editor.pasteHTMLarea;
               if (ha) {
                 ha.blur();
                 ha.innerHTML = '';
                 ha.style.visibility = 'hidden';
               }
               */
               const ta = editor.pasteTextarea;
               ta.blur();
               ta.style.display = "none";
               SocialCalc.KeyboardFocus();
            }, 200);

            return true;

         case "[ctrl-v]":
            if (editor.noEdit) return true; // not if no edit
            if (SocialCalc.Callbacks && SocialCalc.Callbacks.IsCellEditable) {
               if (!SocialCalc.Callbacks.IsCellEditable(editor)) {
                  return true;
               } 
            }

            const showPasteTextArea = () => {
               ta = editor.pasteTextarea;
               ta.value = "";

               cell = SocialCalc.GetEditorCellElement(editor, editor.ecell.row, editor.ecell.col);
               if (cell) {
                  position = SocialCalc.GetElementPosition(cell.element);
                  ta.style.left = `${position.left - 1}px`;
                  ta.style.top = `${position.top - 1}px`;
               }
               ta.style.display = "block";
               ta.value = "";  // must follow "block" setting for Webkit
               ta.focus();
            };

            ha = editor.pasteHTMLarea;
            if (ha) {
               /* Pasting via HTML - Currently IE only */
               ha.style.visibility = "visible";
               ha.focus();
            } else {
               showPasteTextArea();
            }
            window.setTimeout(() => {
               if (!SocialCalc.GetSpreadsheetControlObject) return;
               const s = SocialCalc.GetSpreadsheetControlObject();
               if (!s) return;
               const editor = s.editor;
               let value = null;
               let isPasteSameAsClipboard = false;

               ha = editor.pasteHTMLarea;
               if (ha) {
                  /* IE: We append a U+FFFC to every TD that's not the last of its row,
                   *     then we obtain innerText, then turn U+FFFC back to \t,
                   *     thereby preserving the cell separations (which gets discarded
                   *     if we simply paste via textarea.
                   */
                  const _ObjectReplacementCharacter_ = String.fromCharCode(0xFFFC);
                  const html = ha.innerHTML;
               if (html.search(/<(?![Bb][Rr])[A-Za-z]/) >= 0) {
                  /* HTML Paste: Mark TDs with U+FFFC accordingly.. */
                  ha.innerHTML = html.replace(
                     /(?:<\/[Tt][Dd]>)/g,
                     _ObjectReplacementCharacter_
                  );
               } else {
                  /* Text Paste: In IE, \t is transformed into &nbsp;, so replace them with U+FFFC. */
                  ha.innerHTML = html.replace(
                     /&[Nn][Bb][Ss][Pp];/g,
                     _ObjectReplacementCharacter_
                  );
               }

               value = ha.innerText.replace(new RegExp(_ObjectReplacementCharacter_, 'g'), '\t');

               ha.innerHTML = '';
               ha.blur();
               ha.style.visibility = "hidden";
            } else {
               const ta = editor.pasteTextarea;
               value = ta.value;
               ta.blur();
               ta.style.display = "none";
            }

            value = value.replace(/\r\n/g, "\n").replace(/\n?$/, '\n');
            const clipstr = SocialCalc.ConvertSaveToOtherFormat(SocialCalc.Clipboard.clipboard, "tab");
            if (value === clipstr || (value.length - clipstr.length === 1 && value.substring(0, value.length - 1) === clipstr)) {
               isPasteSameAsClipboard = true;
            }

            let cmd = "";
            // pastes SocialCalc clipboard if did a Ctrl-C and contents still the same
            // Webkit adds an extra blank line, so need to allow for that
            if (!isPasteSameAsClipboard) {
               cmd = `loadclipboard ${SocialCalc.encodeForSave(SocialCalc.ConvertOtherFormatToSave(value, "tab"))}\n`;
            }
            let cr;
            if (editor.range.hasrange) {
               cr = SocialCalc.crToCoord(editor.range.left, editor.range.top);
            } else {
               cr = editor.ecell.coord;
            }
            cmd += `paste ${cr} formulas`;
            editor.EditorScheduleSheetCommands(cmd, true, false);
            SocialCalc.KeyboardFocus();
         }, 200);
         return true;

      case "[ctrl-z]":
         editor.EditorScheduleSheetCommands("undo", true, false);
         return false;

      case "[ctrl-s]": // !!!! temporary hack
         window.setTimeout(() => {
            if (!SocialCalc.GetSpreadsheetControlObject) return;
            const s = SocialCalc.GetSpreadsheetControlObject();
            if (!s) return;
            const editor = s.editor;
            const sheet = editor.context.sheetobj;
            const cell = sheet.GetAssuredCell(editor.ecell.coord);
            const ntvf = cell.nontextvalueformat ? sheet.valueformats[cell.nontextvalueformat - 0] || "" : "";
            const newntvf = window.prompt("Advanced Feature:\n\nCustom Numeric Format or Command", ntvf);
            if (newntvf !== null) { // not cancelled
               if (newntvf.match(/^cmd:/)) {
                  cmd = newntvf.substring(4); // execute as command
               } else if (newntvf.match(/^edit:/)) {
                  cmd = newntvf.substring(5); // execute as command
                  if (SocialCalc.CtrlSEditor) {
                     SocialCalc.CtrlSEditor(cmd);
                  }
                  return;
               } else {
                  if (editor.range.hasrange) {
                     sel = `${SocialCalc.crToCoord(editor.range.left, editor.range.top)}:${SocialCalc.crToCoord(editor.range.right, editor.range.bottom)}`;
                  } else {
                     sel = editor.ecell.coord;
                  }
                  cmd = `set ${sel} nontextvalueformat ${newntvf}`;
               }
               editor.EditorScheduleSheetCommands(cmd, true, false);
            }
         }, 200);
         return false;

      default:
         break;
      }
      return true;
   };

   /**
    * Set sheet's status callback
    */
   context.sheetobj.statuscallback = SocialCalc.EditorSheetStatusCallback;
   context.sheetobj.statuscallbackparams = this; // this object: the table editor object

   /**
    * @description StatusCallback: all values are called at appropriate times, add with unique name, delete when done
    * 
    * Each value must be an object in the form of:
    * 
    *    func: function(editor, status, arg, params) {...},
    *    params: params value to call func with
    * 
    * The values for status and arg are:
    * 
    *    all the SocialCalc RecalcSheet statuscallbacks, including:
    * 
    *       calccheckdone, calclist length
    *       calcorder, {coord: coord, total: celllist length, count: count}
    *       calcstep, {coord: coord, total: calclist length, count: count}
    *       calcfinished, time in milliseconds
    * 
    *    the command callbacks, like cmdstart and cmdend
    *    cmdendnorender
    * 
    *    calcstart, null
    *    moveecell, new ecell coord
    *    rangechange, "coord:coord" or "coord" or ""
    *    specialkey, keyname ("[esc]")
    * 
    * @member {Object} StatusCallback
    */
   this.StatusCallback = {};

   /**
    * @member {Object} MoveECellCallback - All values are called with editor as arg; add with unique name, delete when done
    */
   this.MoveECellCallback = {};
   
   /**
    * @member {Object} RangeChangeCallback - All values are called with editor as arg; add with unique name, delete when done
    */
   this.RangeChangeCallback = {};
   
   /**
    * @member {Object} SettingsCallbacks - See SocialCalc.SaveEditorSettings
    */
   this.SettingsCallbacks = {};

   /**
    * Set initial cursor
    */
   this.ecell = {coord: "A1", row: 1, col: 1};
   context.highlights[this.ecell.coord] = "cursor";

   /**
    * Initialize range data
    * Range has at least hasrange (true/false).
    * It may also have: anchorcoord, anchorrow, anchorcol, top, bottom, left, and right.
    * @member {Object} range
    */
   this.range = {hasrange: false};

   /**
    * Initialize range2 data (used to show selections, such as for move)
    * Range2 has at least hasrange (true/false).
    * It may also have: top, bottom, left, and right.
    * @member {Object} range2
    */
   this.range2 = {hasrange: false};
};

/**
 * @description Methods for SocialCalc.TableEditor prototype
 */

/**
 * @method CreateTableEditor
 * @memberof SocialCalc.TableEditor
 * @param {number} width - Editor width
 * @param {number} height - Editor height
 * @returns {*} Result of SocialCalc.CreateTableEditor
 */
SocialCalc.TableEditor.prototype.CreateTableEditor = function(width, height) {
   return SocialCalc.CreateTableEditor(this, width, height);
};

/**
 * @method ResizeTableEditor
 * @memberof SocialCalc.TableEditor
 * @param {number} width - New editor width
 * @param {number} height - New editor height
 * @returns {*} Result of SocialCalc.ResizeTableEditor
 */
SocialCalc.TableEditor.prototype.ResizeTableEditor = function(width, height) {
   return SocialCalc.ResizeTableEditor(this, width, height);
};

/**
 * @method SaveEditorSettings
 * @memberof SocialCalc.TableEditor
 * @returns {*} Result of SocialCalc.SaveEditorSettings
 */
SocialCalc.TableEditor.prototype.SaveEditorSettings = function() {
   return SocialCalc.SaveEditorSettings(this);
};

/**
 * @method LoadEditorSettings
 * @memberof SocialCalc.TableEditor
 * @param {string} str - Settings string
 * @param {*} flags - Settings flags
 * @returns {*} Result of SocialCalc.LoadEditorSettings
 */
SocialCalc.TableEditor.prototype.LoadEditorSettings = function(str, flags) {
   return SocialCalc.LoadEditorSettings(this, str, flags);
};

/**
 * @method EditorRenderSheet
 * @memberof SocialCalc.TableEditor
 */
SocialCalc.TableEditor.prototype.EditorRenderSheet = function() {
   SocialCalc.EditorRenderSheet(this);
};

/**
 * @method EditorScheduleSheetCommands
 * @memberof SocialCalc.TableEditor
 * @param {string} cmdstr - Command string
 * @param {boolean} saveundo - Whether to save undo
 * @param {boolean} ignorebusy - Whether to ignore busy state
 */
SocialCalc.TableEditor.prototype.EditorScheduleSheetCommands = function(cmdstr, saveundo, ignorebusy) {
   SocialCalc.EditorScheduleSheetCommands(this, cmdstr, saveundo, ignorebusy);
};

/**
 * @method ScheduleSheetCommands
 * @memberof SocialCalc.TableEditor
 * @param {string} cmdstr - Command string
 * @param {boolean} saveundo - Whether to save undo
 */
SocialCalc.TableEditor.prototype.ScheduleSheetCommands = function(cmdstr, saveundo) {
   this.context.sheetobj.ScheduleSheetCommands(cmdstr, saveundo);
};

/**
 * @method SheetUndo
 * @memberof SocialCalc.TableEditor
 */
SocialCalc.TableEditor.prototype.SheetUndo = function() {
   this.context.sheetobj.SheetUndo();
};

/**
 * @method SheetRedo
 * @memberof SocialCalc.TableEditor
 */
SocialCalc.TableEditor.prototype.SheetRedo = function() {
   this.context.sheetobj.SheetRedo();
};

/**
 * @method EditorStepSet
 * @memberof SocialCalc.TableEditor
 * @param {string} status - Status value
 * @param {*} arg - Argument value
 */
SocialCalc.TableEditor.prototype.EditorStepSet = function(status, arg) {
   SocialCalc.EditorStepSet(this, status, arg);
};

/**
 * @method GetStatuslineString
 * @memberof SocialCalc.TableEditor
 * @param {string} status - Status value
 * @param {*} arg - Argument value
 * @param {*} params - Parameters
 * @returns {string} Status line string
 */
SocialCalc.TableEditor.prototype.GetStatuslineString = function(status, arg, params) {
   return SocialCalc.EditorGetStatuslineString(this, status, arg, params);
};

/**
 * @method EditorMouseRegister
 * @memberof SocialCalc.TableEditor
 * @returns {*} Result of SocialCalc.EditorMouseRegister
 */
SocialCalc.TableEditor.prototype.EditorMouseRegister = function() {
   return SocialCalc.EditorMouseRegister(this);
};

/**
 * @method EditorMouseUnregister
 * @memberof SocialCalc.TableEditor
 * @returns {*} Result of SocialCalc.EditorMouseUnregister
 */
SocialCalc.TableEditor.prototype.EditorMouseUnregister = function() {
   return SocialCalc.EditorMouseUnregister(this);
};

/**
 * @method EditorMouseRange
 * @memberof SocialCalc.TableEditor
 * @param {string} coord - Cell coordinate
 * @returns {*} Result of SocialCalc.EditorMouseRange
 */
SocialCalc.TableEditor.prototype.EditorMouseRange = function(coord) {
   return SocialCalc.EditorMouseRange(this, coord);
};

/**
 * @method EditorProcessKey
 * @memberof SocialCalc.TableEditor
 * @param {string} ch - Character
 * @param {Event} e - Event object
 * @returns {*} Result of SocialCalc.EditorProcessKey
 */
SocialCalc.TableEditor.prototype.EditorProcessKey = function(ch, e) {
   return SocialCalc.EditorProcessKey(this, ch, e);
};

/**
 * @method EditorAddToInput
 * @memberof SocialCalc.TableEditor
 * @param {string} str - String to add
 * @param {string} prefix - Prefix string
 * @returns {*} Result of SocialCalc.EditorAddToInput
 */
SocialCalc.TableEditor.prototype.EditorAddToInput = function(str, prefix) {
   return SocialCalc.EditorAddToInput(this, str, prefix);
};

/**
 * @method DisplayCellContents
 * @memberof SocialCalc.TableEditor
 * @returns {*} Result of SocialCalc.EditorDisplayCellContents
 */
SocialCalc.TableEditor.prototype.DisplayCellContents = function() {
   return SocialCalc.EditorDisplayCellContents(this);
};

/**
 * @method EditorSaveEdit
 * @memberof SocialCalc.TableEditor
 * @param {string} text - Text to save
 * @returns {*} Result of SocialCalc.EditorSaveEdit
 */
SocialCalc.TableEditor.prototype.EditorSaveEdit = function(text) {
   return SocialCalc.EditorSaveEdit(this, text);
};

/**
 * @method EditorApplySetCommandsToRange
 * @memberof SocialCalc.TableEditor
 * @param {string} cmdline - Command line
 * @param {string} type - Command type
 * @returns {*} Result of SocialCalc.EditorApplySetCommandsToRange
 */
SocialCalc.TableEditor.prototype.EditorApplySetCommandsToRange = function(cmdline, type) {
   return SocialCalc.EditorApplySetCommandsToRange(this, cmdline, type);
};

/**
 * @method MoveECellWithKey
 * @memberof SocialCalc.TableEditor
 * @param {string} ch - Character key
 * @returns {*} Result of SocialCalc.MoveECellWithKey
 */
SocialCalc.TableEditor.prototype.MoveECellWithKey = function(ch) {
   return SocialCalc.MoveECellWithKey(this, ch);
};

/**
 * @method MoveECell
 * @memberof SocialCalc.TableEditor
 * @param {string} newcell - New cell coordinate
 * @returns {*} Result of SocialCalc.MoveECell
 */
SocialCalc.TableEditor.prototype.MoveECell = function(newcell) {
   return SocialCalc.MoveECell(this, newcell);
};

/**
 * @method ReplaceCell
 * @memberof SocialCalc.TableEditor
 * @param {HTMLElement} cell - Cell element
 * @param {number} row - Row number
 * @param {number} col - Column number
 */
SocialCalc.TableEditor.prototype.ReplaceCell = function(cell, row, col) {
   SocialCalc.ReplaceCell(this, cell, row, col);
};

/**
 * @method UpdateCellCSS
 * @memberof SocialCalc.TableEditor
 * @param {HTMLElement} cell - Cell element
 * @param {number} row - Row number
 * @param {number} col - Column number
 */
SocialCalc.TableEditor.prototype.UpdateCellCSS = function(cell, row, col) {
   SocialCalc.UpdateCellCSS(this, cell, row, col);
};

/**
 * @method SetECellHeaders
 * @memberof SocialCalc.TableEditor
 * @param {boolean} selected - Whether selected
 */
SocialCalc.TableEditor.prototype.SetECellHeaders = function(selected) {
   SocialCalc.SetECellHeaders(this, selected);
};

/**
 * @method EnsureECellVisible
 * @memberof SocialCalc.TableEditor
 */
SocialCalc.TableEditor.prototype.EnsureECellVisible = function() {
   SocialCalc.EnsureECellVisible(this);
};

/**
 * @method RangeAnchor
 * @memberof SocialCalc.TableEditor
 * @param {string} coord - Cell coordinate
 */
SocialCalc.TableEditor.prototype.RangeAnchor = function(coord) {
   SocialCalc.RangeAnchor(this, coord);
};

/**
 * @method RangeExtend
 * @memberof SocialCalc.TableEditor
 * @param {string} coord - Cell coordinate
 */
SocialCalc.TableEditor.prototype.RangeExtend = function(coord) {
   SocialCalc.RangeExtend(this, coord);
};

/**
 * @method RangeRemove
 * @memberof SocialCalc.TableEditor
 */
SocialCalc.TableEditor.prototype.RangeRemove = function() {
   SocialCalc.RangeRemove(this);
};

/**
 * @method Range2Remove
 * @memberof SocialCalc.TableEditor
 */
SocialCalc.TableEditor.prototype.Range2Remove = function() {
   SocialCalc.Range2Remove(this);
};

/**
 * @method FitToEditTable
 * @memberof SocialCalc.TableEditor
 */
SocialCalc.TableEditor.prototype.FitToEditTable = function() {
   SocialCalc.FitToEditTable(this);
};

/**
 * @method CalculateEditorPositions
 * @memberof SocialCalc.TableEditor
 */
SocialCalc.TableEditor.prototype.CalculateEditorPositions = function() {
   SocialCalc.CalculateEditorPositions(this);
};

/**
 * @method ScheduleRender
 * @memberof SocialCalc.TableEditor
 */
SocialCalc.TableEditor.prototype.ScheduleRender = function() {
   SocialCalc.ScheduleRender(this);
};

/**
 * @method DoRenderStep
 * @memberof SocialCalc.TableEditor
 */
SocialCalc.TableEditor.prototype.DoRenderStep = function() {
   SocialCalc.DoRenderStep(this);
};

/**
 * @method SchedulePositionCalculations
 * @memberof SocialCalc.TableEditor
 */
SocialCalc.TableEditor.prototype.SchedulePositionCalculations = function() {
   SocialCalc.SchedulePositionCalculations(this);
};

/**
 * @method DoPositionCalculations
 * @memberof SocialCalc.TableEditor
 */
SocialCalc.TableEditor.prototype.DoPositionCalculations = function() {
   SocialCalc.DoPositionCalculations(this);
};

/**
 * @method CalculateRowPositions
 * @memberof SocialCalc.TableEditor
 * @param {number} panenum - Pane number
 * @param {Array} positions - Positions array
 * @param {Array} sizes - Sizes array
 * @returns {*} Result of SocialCalc.CalculateRowPositions
 */
SocialCalc.TableEditor.prototype.CalculateRowPositions = function(panenum, positions, sizes) {
   return SocialCalc.CalculateRowPositions(this, panenum, positions, sizes);
};

/**
 * @method CalculateColPositions
 * @memberof SocialCalc.TableEditor
 * @param {number} panenum - Pane number
 * @param {Array} positions - Positions array
 * @param {Array} sizes - Sizes array
 * @returns {*} Result of SocialCalc.CalculateColPositions
 */
SocialCalc.TableEditor.prototype.CalculateColPositions = function(panenum, positions, sizes) {
   return SocialCalc.CalculateColPositions(this, panenum, positions, sizes);
};

/**
 * @method ScrollRelative
 * @memberof SocialCalc.TableEditor
 * @param {boolean} vertical - Whether vertical scroll
 * @param {number} amount - Scroll amount
 */
SocialCalc.TableEditor.prototype.ScrollRelative = function(vertical, amount) {
   SocialCalc.ScrollRelative(this, vertical, amount);
};

/**
 * @method ScrollRelativeBoth
 * @memberof SocialCalc.TableEditor
 * @param {number} vamount - Vertical amount
 * @param {number} hamount - Horizontal amount
 */
SocialCalc.TableEditor.prototype.ScrollRelativeBoth = function(vamount, hamount) {
   SocialCalc.ScrollRelativeBoth(this, vamount, hamount);
};

/**
 * @method PageRelative
 * @memberof SocialCalc.TableEditor
 * @param {boolean} vertical - Whether vertical page
 * @param {number} direction - Page direction
 */
SocialCalc.TableEditor.prototype.PageRelative = function(vertical, direction) {
   SocialCalc.PageRelative(this, vertical, direction);
};

/**
 * @method LimitLastPanes
 * @memberof SocialCalc.TableEditor
 */
SocialCalc.TableEditor.prototype.LimitLastPanes = function() {
   SocialCalc.LimitLastPanes(this);
};

/**
 * @method ScrollTableUpOneRow
 * @memberof SocialCalc.TableEditor
 * @returns {*} Result of SocialCalc.ScrollTableUpOneRow
 */
SocialCalc.TableEditor.prototype.ScrollTableUpOneRow = function() {
   return SocialCalc.ScrollTableUpOneRow(this);
};

/**
 * @method ScrollTableDownOneRow
 * @memberof SocialCalc.TableEditor
 * @returns {*} Result of SocialCalc.ScrollTableDownOneRow
 */
SocialCalc.TableEditor.prototype.ScrollTableDownOneRow = function() {
   return SocialCalc.ScrollTableDownOneRow(this);
};

/**
 * @method ScrollTableLeftOneCol
 * @memberof SocialCalc.TableEditor
 * @returns {*} Result of SocialCalc.ScrollTableLeftOneCol
 */
SocialCalc.TableEditor.prototype.ScrollTableLeftOneCol = function() {
   return SocialCalc.ScrollTableLeftOneCol(this);
};

/**
 * @method ScrollTableRightOneCol
 * @memberof SocialCalc.TableEditor
 * @returns {*} Result of SocialCalc.ScrollTableRightOneCol
 */
SocialCalc.TableEditor.prototype.ScrollTableRightOneCol = function() {
   return SocialCalc.ScrollTableRightOneCol(this);
};

/**
 * @namespace Functions
 * @description Main functions for SocialCalc
 */

/**
 * @function CreateTableEditor
 * @memberof SocialCalc
 * @param {Object} editor - Editor instance
 * @param {number} width - Editor width
 * @param {number} height - Editor height
 * @description Creates a table editor with specified dimensions
 */
SocialCalc.CreateTableEditor = function(editor, width, height) {
   const scc = SocialCalc.Constants;
   const AssignID = SocialCalc.AssignID;

   editor.toplevel = document.createElement("div");
   editor.width = width;
   editor.height = height;

   editor.griddiv = document.createElement("div");
   editor.tablewidth = Math.max(0, width - scc.defaultTableControlThickness);
   editor.tableheight = Math.max(0, height - scc.defaultTableControlThickness);
   editor.griddiv.style.width = `${editor.tablewidth}px`;
   editor.griddiv.style.height = `${editor.tableheight}px`;
   editor.griddiv.style.overflow = "hidden";
   editor.griddiv.style.cursor = "default";
   if (scc.cteGriddivClass) editor.griddiv.className = scc.cteGriddivClass;
   AssignID(editor, editor.griddiv, "griddiv");

   editor.FitToEditTable();

   editor.EditorRenderSheet();

   editor.griddiv.appendChild(editor.fullgrid);

   editor.verticaltablecontrol = new SocialCalc.TableControl(editor, true, editor.tableheight);
   editor.verticaltablecontrol.CreateTableControl();
   AssignID(editor, editor.verticaltablecontrol.main, "tablecontrolv");

   editor.horizontaltablecontrol = new SocialCalc.TableControl(editor, false, editor.tablewidth);
   editor.horizontaltablecontrol.CreateTableControl();
   AssignID(editor, editor.horizontaltablecontrol.main, "tablecontrolh");

   let table, tbody, tr, td, img, anchor, ta, ha;

   table = document.createElement("table");
   editor.layouttable = table;
   table.cellSpacing = 0;
   table.cellPadding = 0;
   AssignID(editor, table, "layouttable");

   tbody = document.createElement("tbody");
   table.appendChild(tbody);

   tr = document.createElement("tr");
   tbody.appendChild(tr);
   td = document.createElement("td");
   td.appendChild(editor.griddiv);
   tr.appendChild(td);
   td = document.createElement("td");
   td.appendChild(editor.verticaltablecontrol.main);
   tr.appendChild(td);

   tr = document.createElement("tr");
   tbody.appendChild(tr);
   td = document.createElement("td");
   td.appendChild(editor.horizontaltablecontrol.main);
   tr.appendChild(td);

   td = document.createElement("td"); // logo display: Required by CPAL License for this code!
   td.style.background = `url(${editor.imageprefix}logo.gif) no-repeat center center`;
   td.innerHTML = `<div style='cursor:pointer;font-size:1px;'><img src='${editor.imageprefix}1x1.gif' border='0' width='18' height='18'></div>`;
   tr.appendChild(td);
   editor.logo = td;
   AssignID(editor, editor.logo, "logo");
   SocialCalc.TooltipRegister(td.firstChild.firstChild, "SocialCalc", null);

   editor.toplevel.appendChild(editor.layouttable);

   if (!editor.noEdit) {
      editor.inputEcho = new SocialCalc.InputEcho(editor);
      AssignID(editor, editor.inputEcho.main, "inputecho");
   }
   editor.cellhandles = new SocialCalc.CellHandles(editor);

   ta = document.createElement("textarea"); // used for ctrl-c/ctrl-v where an invisible text area is needed
   SocialCalc.setStyles(ta, "display:none;position:absolute;height:1px;width:1px;opacity:0;filter:alpha(opacity=0);");
   ta.value = "";
   editor.pasteTextarea = ta;
   AssignID(editor, editor.pasteTextarea, "pastetextarea");

   if (navigator.userAgent.match(/Safari\//) && !navigator.userAgent.match(/Chrome\//)) { // special code for Safari 5 change
      window.removeEventListener('beforepaste', SocialCalc.SafariPasteFunction, false);
      window.addEventListener('beforepaste', SocialCalc.SafariPasteFunction, false);
      window.removeEventListener('beforecopy', SocialCalc.SafariPasteFunction, false);
      window.addEventListener('beforecopy', SocialCalc.SafariPasteFunction, false);
      window.removeEventListener('beforecut', SocialCalc.SafariPasteFunction, false);
      window.addEventListener('beforecut', SocialCalc.SafariPasteFunction, false);
   }

   editor.toplevel.appendChild(editor.pasteTextarea);

   const div = document.createElement("div");
   div.innerHTML = '    <br/>';
   if (div.firstChild.nodeType === 1) {
      /* We are running in IE -- Using HTML-based area for Ctrl-V */
      ha = document.createElement("div"); // used for ctrl-v where an invisible html area is needed
      editor.pasteHTMLarea = ha;
      editor.toplevel.appendChild(editor.pasteHTMLarea);
      ha.contentEditable = true;
      AssignID(editor, editor.pasteHTMLarea, "pastehtmlarea");
      SocialCalc.setStyles(ha, "display:block;visibility:hidden;position:absolute;height:1px;width:1px;opacity:0;filter:alpha(opacity=0);overflow:hidden");
   }

   SocialCalc.MouseWheelRegister(editor.toplevel, {WheelMove: SocialCalc.EditorProcessMouseWheel, editor: editor});

   if (SocialCalc.HasTouch) {
      SocialCalc.TouchRegister(editor.toplevel, {
         Swipe: SocialCalc.EditorProcessSwipe, 
         DoubleTap: SocialCalc.EditorProcessDoubleTap, 
         SingleTap: SocialCalc.EditorProcessSingleTap, 
         editor: editor
      });
   }

   if (editor.inputBox) { // this seems to fix an obscure bug with Firefox 2 Mac where Ctrl-V doesn't get fired right
      if (editor.inputBox.element) {
         editor.inputBox.element.focus();
         editor.inputBox.element.blur();
      }
   }
   SocialCalc.KeyboardSetFocus(editor);

   // do status reporting things
   SocialCalc.EditorSheetStatusCallback(null, "startup", null, editor);

   // done
   return editor.toplevel;
};

/**
 * @function SafariPasteFunction
 * @memberof SocialCalc
 * @description Special code needed for change that occurred with Safari 5 that made paste not work for some reason
 * @param {Event} e - The paste event
 */
SocialCalc.SafariPasteFunction = function(e) {
   e.preventDefault();
};

/**
 * @function ResizeTableEditor
 * @memberof SocialCalc
 * @description Move things around as appropriate and resize
 * @param {Object} editor - The table editor instance
 * @param {number} width - New width
 * @param {number} height - New height
 */
SocialCalc.ResizeTableEditor = function(editor, width, height) {
   const scc = SocialCalc.Constants;

   editor.width = width;
   editor.height = height;

   editor.toplevel.style.width = `${width}px`;
   editor.toplevel.style.height = `${height}px`;

   editor.tablewidth = Math.max(0, width - scc.defaultTableControlThickness);
   editor.tableheight = Math.max(0, height - scc.defaultTableControlThickness);
   editor.griddiv.style.width = `${editor.tablewidth}px`;
   editor.griddiv.style.height = `${editor.tableheight}px`;

   editor.verticaltablecontrol.main.style.height = `${editor.tableheight}px`;
   editor.horizontaltablecontrol.main.style.width = `${editor.tablewidth}px`;

   editor.FitToEditTable();

   editor.ScheduleRender();

   return;
};

/**
 * @function SaveEditorSettings
 * @memberof SocialCalc
 * @description Returns a string representation of the pane settings, etc.
 * 
 * The format is:
 * 
 *    version:1.0
 *    rowpane:panenumber:firstnum:lastnum
 *    colpane:panenumber:firstnum:lastnum
 *    ecell:coord -- if set
 *    range:anchorcoord:top:bottom:left:right -- if set
 * 
 * You can add additional values to be saved by using editor.SettingsCallbacks:
 * 
 *   editor.SettingsCallbacks["item-name"] = {save: savefunction, load: loadfunction}
 * 
 * where savefunction(editor, "item-name") returns a string with the new lines to be added to the saved settings
 * which include the trailing newlines, and loadfunction(editor, "item-name", line, flags) is given the line to process
 * without the trailing newlines.
 * 
 * @param {Object} editor - The table editor instance
 * @returns {string} String representation of the editor settings
 */
SocialCalc.SaveEditorSettings = function(editor) {
   let i, setting;
   const context = editor.context;
   const range = editor.range;
   let result = "";

   result += "version:1.0\n";

   for (i = 0; i < context.rowpanes.length; i++) {
      result += `rowpane:${i}:${context.rowpanes[i].first}:${context.rowpanes[i].last}\n`;
   }
   for (i = 0; i < context.colpanes.length; i++) {
      result += `colpane:${i}:${context.colpanes[i].first}:${context.colpanes[i].last}\n`;
   }

   if (editor.ecell) {
      result += `ecell:${editor.ecell.coord}\n`;
   }

   if (range.hasrange) {
      result += `range:${range.anchorcoord}:${range.top}:${range.bottom}:${range.left}:${range.right}\n`;
   }

   for (setting in editor.SettingsCallbacks) {
      result += editor.SettingsCallbacks[setting].save(editor, setting);
   }

   return result;
};

/**
 * @function LoadEditorSettings
 * @memberof SocialCalc
 * @description Sets the editor settings based on str. See SocialCalc.SaveEditorSettings for more details.
 * Unrecognized lines are ignored.
 * @param {Object} editor - The table editor instance
 * @param {string} str - Settings string to load
 * @param {*} flags - Loading flags
 */
SocialCalc.LoadEditorSettings = function(editor, str, flags) {
   const lines = str.split(/\r\n|\n/);
   let parts = [];
   let line, i, cr, row, col, coord, setting;
   const context = editor.context;
   let highlights, range;

   context.rowpanes = [{first: 1, last: 1}]; // reset to start
   context.colpanes = [{first: 1, last: 1}];
   editor.ecell = null;
   editor.range = {hasrange: false};
   editor.range2 = {hasrange: false};
   range = editor.range;
   context.highlights = {};
   highlights = context.highlights;

   for (i = 0; i < lines.length; i++) {
      line = lines[i];
      parts = line.split(":");
      setting = parts[0];
      switch (setting) {
         case "version":
            break;

         case "rowpane":
            context.rowpanes[parts[1] - 0] = {first: parts[2] - 0, last: parts[3] - 0};
            break;

         case "colpane":
            context.colpanes[parts[1] - 0] = {first: parts[2] - 0, last: parts[3] - 0};
            break;

         case "ecell":
            editor.ecell = SocialCalc.coordToCr(parts[1]);
            editor.ecell.coord = parts[1];
            highlights[parts[1]] = "cursor";
            break;

         case "range":
            range.hasrange = true;
            range.anchorcoord = parts[1];
            cr = SocialCalc.coordToCr(range.anchorcoord);
            range.anchorrow = cr.row;
            range.anchorcol = cr.col;
            range.top = parts[2] - 0;
            range.bottom = parts[3] - 0;
            range.left = parts[4] - 0;
            range.right = parts[5] - 0;
            for (row = range.top; row <= range.bottom; row++) {
               for (col = range.left; col <= range.right; col++) {
                  coord = SocialCalc.crToCoord(col, row);
                  if (highlights[coord] !== "cursor") {
                     highlights[coord] = "range";
                  }
               }
            }
            break;

         default:
            if (editor.SettingsCallbacks[setting]) {
               editor.SettingsCallbacks[setting].load(editor, setting, line, flags);
            }
            break;
      }
   }

   return;
};
/**
 * @function EditorRenderSheet
 * @memberof SocialCalc
 * @description Renders the sheet and updates editor.fullgrid. Sets event handlers.
 * @param {Object} editor - The table editor instance
 */
SocialCalc.EditorRenderSheet = function(editor) {
   editor.EditorMouseUnregister();

   editor.fullgrid = editor.context.RenderSheet(editor.fullgrid);

   if (editor.ecell) editor.SetECellHeaders("selected");

   SocialCalc.AssignID(editor, editor.fullgrid, "fullgrid"); // give it an id

   editor.EditorMouseRegister();
};

/**
 * @function EditorScheduleSheetCommands
 * @memberof SocialCalc
 * @description Schedules sheet commands for execution
 * @param {Object} editor - The table editor instance
 * @param {string} cmdstr - Command string to execute
 * @param {boolean} saveundo - Whether to save undo information
 * @param {boolean} ignorebusy - Whether to ignore busy state
 */
SocialCalc.EditorScheduleSheetCommands = function(editor, cmdstr, saveundo, ignorebusy) {
   if (editor.state !== "start" && !ignorebusy) { // ignore commands if editing a cell
      return;
   }

   if (editor.busy && !ignorebusy) { // hold off on commands if doing one
      editor.deferredCommands.push({cmdstr: cmdstr, saveundo: saveundo});
      return;
   }

   switch (cmdstr) {
      case "recalc":
      case "redisplay":
         editor.context.sheetobj.ScheduleSheetCommands(cmdstr, false);
         break;

      case "undo":
         editor.SheetUndo();
         break;

      case "redo":
         editor.SheetRedo();
         break;

      default:
         editor.context.sheetobj.ScheduleSheetCommands(cmdstr, saveundo);
         break;
   }
};

/**
 * @function EditorSheetStatusCallback
 * @memberof SocialCalc
 * @description Called during recalc, executing commands, etc.
 * @param {*} recalcdata - Recalculation data
 * @param {string} status - Current status
 * @param {*} arg - Status argument
 * @param {Object} editor - The table editor instance
 */
SocialCalc.EditorSheetStatusCallback = function(recalcdata, status, arg, editor) {
   let f, cell, dcmd;
   const sheetobj = editor.context.sheetobj;

   /**
    * @description Signal status to all registered callbacks
    * @param {string} s - Status to signal
    */
   const signalstatus = (s) => {
      for (f in editor.StatusCallback) {
         if (editor.StatusCallback[f].func) {
            editor.StatusCallback[f].func(editor, s, arg, editor.StatusCallback[f].params);
         }
      }
   };

   switch (status) {
      case "startup":
         break;

      case "cmdstart":
         editor.busy = true;
         sheetobj.celldisplayneeded = "";
         break;

      case "cmdextension":
         break;

      case "cmdend":
         signalstatus(status);

         if (sheetobj.changedrendervalues) {
            editor.context.PrecomputeSheetFontsAndLayouts();
            editor.context.CalculateCellSkipData();
            sheetobj.changedrendervalues = false;
         }

         if (sheetobj.celldisplayneeded && !sheetobj.renderneeded) {
            const cr = SocialCalc.coordToCr(sheetobj.celldisplayneeded);
            cell = SocialCalc.GetEditorCellElement(editor, cr.row, cr.col);
            editor.ReplaceCell(cell, cr.row, cr.col);
         }

         if (editor.deferredCommands.length) {
            dcmd = editor.deferredCommands.shift();
            editor.EditorScheduleSheetCommands(dcmd.cmdstr, dcmd.saveundo, true);
            return;
         }

         if (sheetobj.attribs.needsrecalc &&
               (sheetobj.attribs.recalc !== "off" || sheetobj.recalconce)
               && editor.recalcFunction) {
            editor.FitToEditTable();
            sheetobj.renderneeded = false; // recalc will force a render
            if (sheetobj.recalconce) delete sheetobj.recalconce; // only do once
            editor.recalcFunction(editor);
         } else {
            if (sheetobj.renderneeded) {
               editor.FitToEditTable();
               sheetobj.renderneeded = false;
               editor.ScheduleRender();
            } else {
               editor.SchedulePositionCalculations(); // just in case command changed positions
//               editor.busy = false;
//               signalstatus("cmdendnorender");
            }
         }
         return;

      case "calcstart":
         editor.busy = true;
         break;

      case "calccheckdone":
      case "calcorder":
      case "calcstep":
      case "calcloading":
      case "calcserverfunc":
         break;

      case "calcfinished":
         signalstatus(status);
         editor.ScheduleRender();
         return;

      case "schedrender":
         editor.busy = true; // in case got here without cmd or recalc
         break;

      case "renderdone":
         break;

      case "schedposcalc":
         editor.busy = true; // in case got here without cmd or recalc
         break;

      case "doneposcalc":
         if (editor.deferredCommands.length) {
            signalstatus(status);
            dcmd = editor.deferredCommands.shift();
            editor.EditorScheduleSheetCommands(dcmd.cmdstr, dcmd.saveundo, true);
         } else {
            editor.busy = false;
            signalstatus(status);
            if (editor.state === "start") editor.DisplayCellContents(); // make sure up to date
         }
         return;

      default:
         addmsg(`Unknown status: ${status}`);
         break;
   }

   signalstatus(status);

   return;
};

/**
 * @description Timer-driven steps for use with SocialCalc.EditorSheetStatusCallback
 * @type {Object}
 * @property {Object|null} editor - For callback
 */
SocialCalc.EditorStepInfo = {
//   status: "", // saved value to pass to callback
   editor: null // for callback
//   arg: null, // for callback
//   timerobj: null
};

/*
SocialCalc.EditorStepSet = function(editor, status, arg) {
   const esi = SocialCalc.EditorStepInfo;
addmsg("step: "+status);
   if (esi.timerobj) {
alert("Already waiting. Old/new: "+esi.status+"/"+status);
   }
   esi.editor = editor;
   esi.status = status;
   esi.timerobj = window.setTimeout(SocialCalc.EditorStepDone, 1);
};

SocialCalc.EditorStepDone = function() {
   const esi = SocialCalc.EditorStepInfo;
   esi.timerobj = null;
   SocialCalc.EditorSheetStatusCallback(null, esi.status, null, esi.editor);
};
*/

/**
 * @function EditorGetStatuslineString
 * @memberof SocialCalc
 * @description Generates status line string based on current editor state
 * @param {Object} editor - The table editor instance
 * @param {string} status - Current status
 * @param {*} arg - Status argument
 * @param {Object} params - Parameters object where it can use "calculating" and "command" to keep track of state
 * @returns {string} String for status line
 */
SocialCalc.EditorGetStatuslineString = function(editor, status, arg, params) {
   const scc = SocialCalc.Constants;

   let sstr, progress, coord, circ, r, c, cell, sum, ele;

   progress = "";

   switch (status) {
      case "moveecell":
      case "rangechange":
      case "startup":
         break;
      case "cmdstart":
         params.command = true;
         document.body.style.cursor = "progress";
         editor.griddiv.style.cursor = "progress";
         progress = scc.s_statusline_executing;
         break;
      case "cmdextension":
         progress = `Command Extension: ${arg}`;
         break;
      case "cmdend":
         params.command = false;
         break;
      case "schedrender":
         progress = scc.s_statusline_displaying;
         break;
      case "renderdone":
         progress = " ";
         break;
      case "schedposcalc":
         progress = scc.s_statusline_displaying;
         break;
      case "cmdendnorender":
      case "doneposcalc":
         document.body.style.cursor = "default";
         editor.griddiv.style.cursor = "default";
         break;
      case "calcorder":
         progress = scc.s_statusline_ordering + Math.floor(100 * arg.count / (arg.total || 1)) + "%";
         break;
      case "calcstep":
         progress = scc.s_statusline_calculating + Math.floor(100 * arg.count / (arg.total || 1)) + "%";
         break;
      case "calcloading":
         progress = `${scc.s_statusline_calculatingls}: ${arg.sheetname}`;
         break;
      case "calcserverfunc":
         progress = scc.s_statusline_calculating + Math.floor(100 * arg.count / (arg.total || 1)) + "%, " + 
                   scc.s_statusline_doingserverfunc + arg.funcname + scc.s_statusline_incell + arg.coord;
         break;
      case "calcstart":
         params.calculating = true;
         document.body.style.cursor = "progress";
         editor.griddiv.style.cursor = "progress"; // griddiv has an explicit cursor style
         progress = scc.s_statusline_calcstart;
         break;
      case "calccheckdone":
         break;
      case "calcfinished":
         params.calculating = false;
         break;
      default:
         progress = status;
         break;
   }

   if (!progress && params.calculating) {
      progress = scc.s_statusline_calculating;
   }

   // if there is a range, calculate sum (not during busy times)
   if (!params.calculating && !params.command && !progress && editor.range.hasrange 
       && (editor.range.left !== editor.range.right || editor.range.top !== editor.range.bottom)) {
      sum = 0;
      for (r = editor.range.top; r <= editor.range.bottom; r++) {
         for (c = editor.range.left; c <= editor.range.right; c++) {
            cell = editor.context.sheetobj.cells[SocialCalc.crToCoord(c, r)];
            if (!cell) continue;
            if (cell.valuetype && cell.valuetype.charAt(0) === "n") {
               sum += cell.datavalue - 0;
            }
         }
      }

      sum = SocialCalc.FormatNumber.formatNumberWithFormat(sum, "[,]General", "");

      coord = `${SocialCalc.crToCoord(editor.range.left, editor.range.top)}:${SocialCalc.crToCoord(editor.range.right, editor.range.bottom)}`;
      progress = `${coord} (${editor.range.right - editor.range.left + 1}x${editor.range.bottom - editor.range.top + 1}) ${scc.s_statusline_sum}=${sum} ${progress}`;
   }
   sstr = `${editor.ecell.coord} &nbsp; ${progress}`;

   if (!params.calculating && editor.context.sheetobj.attribs.needsrecalc === "yes") {
      sstr += ` &nbsp; ${scc.s_statusline_recalcneeded}`;
   }

   circ = editor.context.sheetobj.attribs.circularreferencecell;
   if (circ) {
      circ = circ.replace(/\|/, " referenced by ");
      sstr += ` &nbsp; ${scc.s_statusline_circref}${circ}</span>`;
   }

   return sstr;
};
/**
 * @namespace EditorMouseInfo
 * @memberof SocialCalc
 * @description Mouse handling information for the editor
 */
SocialCalc.EditorMouseInfo = {
   /**
    * @description The registeredElements array is used to identify editor grid in which the mouse is doing things.
    * One item for each active editor, each an object with:
    *    .element, .editor
    * @type {Array<Object>}
    */
   registeredElements: [],

   /**
    * @member {Object|null} editor - Editor being processed (between mousedown and mouseup)
    */
   editor: null,
   
   /**
    * @member {HTMLElement|null} element - Element being processed
    */
   element: null,

   /**
    * @member {boolean} ignore - If true, mousedowns are ignored
    */
   ignore: false,

   /**
    * @member {string} mousedowncoord - Coord where mouse went down for drag range
    */
   mousedowncoord: "",
   
   /**
    * @member {string} mouselastcoord - Coord where mouse last was during drag
    */
   mouselastcoord: "",
   
   /**
    * @member {string} mouseresizecol - Col being resized
    */
   mouseresizecol: "",
   
   /**
    * @member {number|null} mouseresizeclientx - Where resize started
    */
   mouseresizeclientx: null,
   
   /**
    * @member {HTMLElement|null} mouseresizedisplay - Element tracking new size
    */
   mouseresizedisplay: null
};

/**
 * @function EditorMouseRegister
 * @memberof SocialCalc
 * @description Registers mouse event handlers for the editor
 * @param {Object} editor - The table editor instance
 */
SocialCalc.EditorMouseRegister = function(editor) {
   const mouseinfo = SocialCalc.EditorMouseInfo;
   const element = editor.fullgrid;
   let i;

   for (i = 0; i < mouseinfo.registeredElements.length; i++) {
      if (mouseinfo.registeredElements[i].editor === editor) {
         if (mouseinfo.registeredElements[i].element === element) {
            return; // already set - don't do it again
         }
         break;
      }
   }

   if (i < mouseinfo.registeredElements.length) {
      mouseinfo.registeredElements[i].element = element;
   } else {
      mouseinfo.registeredElements.push({element: element, editor: editor});
   }

   if (element.addEventListener) { // DOM Level 2 -- Firefox, et al
      element.addEventListener("mousedown", SocialCalc.ProcessEditorMouseDown, false);
      element.addEventListener("dblclick", SocialCalc.ProcessEditorDblClick, false);
   } else if (element.attachEvent) { // IE 5+
      element.attachEvent("onmousedown", SocialCalc.ProcessEditorMouseDown);
      element.attachEvent("ondblclick", SocialCalc.ProcessEditorDblClick);
   } else { // don't handle this
      throw new Error("Browser not supported");
   }

   mouseinfo.ignore = false; // just in case

   return;
};

/**
 * @function EditorMouseUnregister
 * @memberof SocialCalc
 * @description Unregisters mouse event handlers for the editor
 * @param {Object} editor - The table editor instance
 */
SocialCalc.EditorMouseUnregister = function(editor) {
   const mouseinfo = SocialCalc.EditorMouseInfo;
   const element = editor.fullgrid;
   let i, oldelement;

   for (i = 0; i < mouseinfo.registeredElements.length; i++) {
      if (mouseinfo.registeredElements[i].editor === editor) {
         break;
      }
   }

   if (i < mouseinfo.registeredElements.length) {
      oldelement = mouseinfo.registeredElements[i].element; // remove old handlers
      if (oldelement.removeEventListener) { // DOM Level 2
         oldelement.removeEventListener("mousedown", SocialCalc.ProcessEditorMouseDown, false);
         oldelement.removeEventListener("dblclick", SocialCalc.ProcessEditorDblClick, false);
      } else if (oldelement.detachEvent) { // IE
         oldelement.detachEvent("onmousedown", SocialCalc.ProcessEditorMouseDown);
         oldelement.detachEvent("ondblclick", SocialCalc.ProcessEditorDblClick);
      }
      mouseinfo.registeredElements.splice(i, 1);
   }

   return;
};

/**
 * @function ProcessEditorMouseDown
 * @memberof SocialCalc
 * @description Processes mouse down events on the editor
 * @param {Event} e - The mouse event
 */
SocialCalc.ProcessEditorMouseDown = function(e) {
   let editor, result, coord, textarea, wval, range;

   const event = e || window.event;

   const viewport = SocialCalc.GetViewportInfo();
   const clientX = event.clientX + viewport.horizontalScroll;
   const clientY = event.clientY + viewport.verticalScroll;

   const mouseinfo = SocialCalc.EditorMouseInfo;
   let ele = event.target || event.srcElement; // source object is often within what we want
   let mobj;

   if (mouseinfo.ignore) return; // ignore this

   for (mobj = null; !mobj && ele; ele = ele.parentNode) { // go up tree looking for one of our elements
      mobj = SocialCalc.LookupElement(ele, mouseinfo.registeredElements);
   }
   if (!mobj) {
      mouseinfo.editor = null;
      return; // not one of our elements
   }

   editor = mobj.editor;
   mouseinfo.element = ele;
   range = editor.range;
   result = SocialCalc.GridMousePosition(editor, clientX, clientY);

   if (!result || result.rowheader) return; // not on a cell or col header
   mouseinfo.editor = editor; // remember for later

   if (result.colheader && result.coltoresize) { // col header - do drag resize
      SocialCalc.ProcessEditorColsizeMouseDown(e, ele, result);
      return;
   }

   if (!result.coord) return; // not us

   if (!range.hasrange) {
      if (e.shiftKey)
         editor.RangeAnchor();
   }

   coord = editor.MoveECell(result.coord);

   if (range.hasrange) {
      if (e.shiftKey)
         editor.RangeExtend();
      else
         editor.RangeRemove();
   }

   mouseinfo.mousedowncoord = coord; // remember if starting drag range select
   mouseinfo.mouselastcoord = coord;

   editor.EditorMouseRange(coord);

   SocialCalc.KeyboardSetFocus(editor);
   if (editor.state !== "start" && editor.inputBox) editor.inputBox.element.focus();

   // Event code from JavaScript, Flanagan, 5th Edition, pg. 422
   if (document.addEventListener) { // DOM Level 2 -- Firefox, et al
      document.addEventListener("mousemove", SocialCalc.ProcessEditorMouseMove, true); // capture everywhere
      document.addEventListener("mouseup", SocialCalc.ProcessEditorMouseUp, true); // capture everywhere
   } else if (ele.attachEvent) { // IE 5+
      ele.setCapture();
      ele.attachEvent("onmousemove", SocialCalc.ProcessEditorMouseMove);
      ele.attachEvent("onmouseup", SocialCalc.ProcessEditorMouseUp);
      ele.attachEvent("onlosecapture", SocialCalc.ProcessEditorMouseUp);
   }
   if (event.stopPropagation) event.stopPropagation(); // DOM Level 2
   else event.cancelBubble = true; // IE 5+
   if (event.preventDefault) event.preventDefault(); // DOM Level 2
   else event.returnValue = false; // IE 5+

   return;
};

/**
 * @function EditorMouseRange
 * @memberof SocialCalc
 * @description Handles mouse range selection for the editor
 * @param {Object} editor - The table editor instance
 * @param {string} coord - Cell coordinate
 */
SocialCalc.EditorMouseRange = function(editor, coord) {
   let inputtext, wval;
   const range = editor.range;

   switch (editor.state) { // editing a cell - shouldn't get here if no inputBox
      case "input":
         inputtext = editor.inputBox.GetText();
         wval = editor.workingvalues;
         if (("(+-*/,:!&<>=^".indexOf(inputtext.slice(-1)) >= 0 && inputtext.slice(0, 1) === "=") ||
             (inputtext === "=")) {
            wval.partialexpr = inputtext;
         }
         if (wval.partialexpr) { // if in pointing operation
            if (coord) {
               if (range.hasrange) {
                  const sheetpref = (wval.currentsheet === wval.startsheet) ? "" : `${wval.currentsheet}!`;
                  editor.inputBox.SetText(`${wval.partialexpr}${sheetpref}${SocialCalc.crToCoord(range.left, range.top)}:${sheetpref}${SocialCalc.crToCoord(range.right, range.bottom)}`);
               } else {
                  const sheetpref = (wval.currentsheet === wval.startsheet) ? "" : `${wval.currentsheet}!`;
                  editor.inputBox.SetText(`${wval.partialexpr}${sheetpref}${coord}`);
               }
            }
         } else { // not in point -- done editing
            editor.inputBox.Blur();
            editor.inputBox.ShowInputBox(false);
            editor.state = "start";
            editor.cellhandles.ShowCellHandles(true);
            editor.EditorSaveEdit();
            editor.inputBox.DisplayCellContents(null);
         }
         break;

      case "inputboxdirect":
         editor.inputBox.Blur();
         editor.inputBox.ShowInputBox(false);
         editor.state = "start";
         editor.cellhandles.ShowCellHandles(true);
         editor.EditorSaveEdit();
         editor.inputBox.DisplayCellContents(null);
         break;
   }
};
/**
 * @function ProcessEditorMouseMove
 * @memberof SocialCalc
 * @description Processes mouse move events on the editor
 * @param {Event} e - The mouse event
 */
SocialCalc.ProcessEditorMouseMove = function(e) {
   let editor, element, result, coord, now, textarea, sheetobj, cellobj, wval;

   const event = e || window.event;

   const viewport = SocialCalc.GetViewportInfo();
   const clientX = event.clientX + viewport.horizontalScroll;
   const clientY = event.clientY + viewport.verticalScroll;

   const mouseinfo = SocialCalc.EditorMouseInfo;
   editor = mouseinfo.editor;
   if (!editor) return; // not us, ignore
   if (mouseinfo.ignore) return; // ignore this
   element = mouseinfo.element;

   result = SocialCalc.GridMousePosition(editor, clientX, clientY); // get cell with move

   if (!result) return;

   if (result && !result.coord) {
      SocialCalc.SetDragAutoRepeat(editor, result);
      return;
   }

   SocialCalc.SetDragAutoRepeat(editor, null); // stop repeating if it was

   if (!result.coord) return;

   if (result.coord !== mouseinfo.mouselastcoord) {
      if (!e.shiftKey && !editor.range.hasrange) {
         editor.RangeAnchor(mouseinfo.mousedowncoord);
      }
      editor.MoveECell(result.coord);
      editor.RangeExtend();
   }
   mouseinfo.mouselastcoord = result.coord;

   editor.EditorMouseRange(result.coord);

   if (event.stopPropagation) event.stopPropagation(); // DOM Level 2
   else event.cancelBubble = true; // IE 5+
   if (event.preventDefault) event.preventDefault(); // DOM Level 2
   else event.returnValue = false; // IE 5+

   return;
};

/**
 * @function ProcessEditorMouseUp
 * @memberof SocialCalc
 * @description Processes mouse up events on the editor
 * @param {Event} e - The mouse event
 * @returns {boolean} Returns false
 */
SocialCalc.ProcessEditorMouseUp = function(e) {
   let editor, element, result, coord, now, textarea, sheetobj, cellobj, wval;

   const event = e || window.event;

   const viewport = SocialCalc.GetViewportInfo();
   const clientX = event.clientX + viewport.horizontalScroll;
   const clientY = event.clientY + viewport.verticalScroll;

   const mouseinfo = SocialCalc.EditorMouseInfo;
   editor = mouseinfo.editor;
   if (!editor) return; // not us, ignore
   if (mouseinfo.ignore) return; // ignore this
   element = mouseinfo.element;

   result = SocialCalc.GridMousePosition(editor, clientX, clientY); // get cell with up

   SocialCalc.SetDragAutoRepeat(editor, null); // stop repeating if it was

   if (!result) return;

   if (!result.coord) result.coord = editor.ecell.coord;

   if (editor.range.hasrange) {
      editor.MoveECell(result.coord);
      editor.RangeExtend();
   } else if (result.coord && result.coord !== mouseinfo.mousedowncoord) {
      editor.RangeAnchor(mouseinfo.mousedowncoord);
      editor.MoveECell(result.coord);
      editor.RangeExtend();
   }

   editor.EditorMouseRange(result.coord);

   if (event.stopPropagation) event.stopPropagation(); // DOM Level 2
   else event.cancelBubble = true; // IE 5+
   if (event.preventDefault) event.preventDefault(); // DOM Level 2
   else event.returnValue = false; // IE 5+

   if (document.removeEventListener) { // DOM Level 2
      document.removeEventListener("mousemove", SocialCalc.ProcessEditorMouseMove, true);
      document.removeEventListener("mouseup", SocialCalc.ProcessEditorMouseUp, true);
   } else if (element.detachEvent) { // IE
      element.detachEvent("onlosecapture", SocialCalc.ProcessEditorMouseUp);
      element.detachEvent("onmouseup", SocialCalc.ProcessEditorMouseUp);
      element.detachEvent("onmousemove", SocialCalc.ProcessEditorMouseMove);
      element.releaseCapture();
   }

   mouseinfo.editor = null;

   return false;
};

/**
 * @function ProcessEditorColsizeMouseDown
 * @memberof SocialCalc
 * @description Handles mouse down events for column resizing
 * @param {Event} e - The mouse event
 * @param {HTMLElement} ele - The element
 * @param {Object} result - The grid position result
 */
SocialCalc.ProcessEditorColsizeMouseDown = function(e, ele, result) {
   const event = e || window.event;
   const mouseinfo = SocialCalc.EditorMouseInfo;
   const editor = mouseinfo.editor;
   const viewport = SocialCalc.GetViewportInfo();
   const clientX = event.clientX + viewport.horizontalScroll;

   mouseinfo.mouseresizecolnum = result.coltoresize; // remember col being resized
   mouseinfo.mouseresizecol = SocialCalc.rcColname(result.coltoresize);
   mouseinfo.mousedownclientx = clientX;

   const sizedisplay = document.createElement("div");
   mouseinfo.mouseresizedisplay = sizedisplay;
   sizedisplay.style.width = "auto";
   sizedisplay.style.position = "absolute";
   sizedisplay.style.zIndex = 100;
   sizedisplay.style.top = `${editor.headposition.top + 0}px`;
   sizedisplay.style.left = `${editor.colpositions[result.coltoresize]}px`;
   sizedisplay.innerHTML = `<table cellpadding="0" cellspacing="0"><tr><td style="height:100px;border:1px dashed black;background-color:white;width:${editor.context.colwidth[mouseinfo.mouseresizecolnum] - 2}px;">&nbsp;</td><td><div style="font-size:small;color:white;background-color:gray;padding:4px;">${editor.context.colwidth[mouseinfo.mouseresizecolnum]}</div></td></tr></table>`;
   SocialCalc.setStyles(sizedisplay.firstChild.lastChild.firstChild.childNodes[0], "filter:alpha(opacity=85);opacity:.85;"); // so no warning msg with Firefox about filter

   editor.toplevel.appendChild(sizedisplay);

   // Event code from JavaScript, Flanagan, 5th Edition, pg. 422
   if (document.addEventListener) { // DOM Level 2 -- Firefox, et al
      document.addEventListener("mousemove", SocialCalc.ProcessEditorColsizeMouseMove, true); // capture everywhere
      document.addEventListener("mouseup", SocialCalc.ProcessEditorColsizeMouseUp, true); // capture everywhere
   } else if (editor.toplevel.attachEvent) { // IE 5+
      editor.toplevel.setCapture();
      editor.toplevel.attachEvent("onmousemove", SocialCalc.ProcessEditorColsizeMouseMove);
      editor.toplevel.attachEvent("onmouseup", SocialCalc.ProcessEditorColsizeMouseUp);
      editor.toplevel.attachEvent("onlosecapture", SocialCalc.ProcessEditorColsizeMouseUp);
   }
   if (event.stopPropagation) event.stopPropagation(); // DOM Level 2
   else event.cancelBubble = true; // IE 5+
   if (event.preventDefault) event.preventDefault(); // DOM Level 2
   else event.returnValue = false; // IE 5+

   return;
};

/**
 * @function ProcessEditorColsizeMouseMove
 * @memberof SocialCalc
 * @description Handles mouse move events during column resizing
 * @param {Event} e - The mouse event
 */
SocialCalc.ProcessEditorColsizeMouseMove = function(e) {
   const event = e || window.event;
   const mouseinfo = SocialCalc.EditorMouseInfo;
   const editor = mouseinfo.editor;
   if (!editor) return; // not us, ignore
   const viewport = SocialCalc.GetViewportInfo();
   const clientX = event.clientX + viewport.horizontalScroll;

   let newsize = (editor.context.colwidth[mouseinfo.mouseresizecolnum] - 0) + (clientX - mouseinfo.mousedownclientx);
   if (newsize < SocialCalc.Constants.defaultMinimumColWidth) newsize = SocialCalc.Constants.defaultMinimumColWidth;

   const sizedisplay = mouseinfo.mouseresizedisplay;
//   sizedisplay.firstChild.lastChild.firstChild.childNodes[1].firstChild.innerHTML = newsize+"";
//   sizedisplay.firstChild.lastChild.firstChild.childNodes[0].firstChild.style.width = (newsize-2)+"px";
   sizedisplay.innerHTML = `<table cellpadding="0" cellspacing="0"><tr><td style="height:100px;border:1px dashed black;background-color:white;width:${newsize - 2}px;">&nbsp;</td><td><div style="font-size:small;color:white;background-color:gray;padding:4px;">${newsize}</div></td></tr></table>`;
   SocialCalc.setStyles(sizedisplay.firstChild.lastChild.firstChild.childNodes[0], "filter:alpha(opacity=85);opacity:.85;"); // so no warning msg with Firefox about filter

   if (event.stopPropagation) event.stopPropagation(); // DOM Level 2
   else event.cancelBubble = true; // IE 5+
   if (event.preventDefault) event.preventDefault(); // DOM Level 2
   else event.returnValue = false; // IE 5+

   return;
};

/**
 * @function ProcessEditorColsizeMouseUp
 * @memberof SocialCalc
 * @description Handles mouse up events for column resizing
 * @param {Event} e - The mouse event
 * @returns {boolean} Returns false
 */
SocialCalc.ProcessEditorColsizeMouseUp = function(e) {
   const event = e || window.event;
   const mouseinfo = SocialCalc.EditorMouseInfo;
   const editor = mouseinfo.editor;
   if (!editor) return; // not us, ignore
   const element = mouseinfo.element;
   const viewport = SocialCalc.GetViewportInfo();
   const clientX = event.clientX + viewport.horizontalScroll;

   if (event.stopPropagation) event.stopPropagation(); // DOM Level 2
   else event.cancelBubble = true; // IE 5+
   if (event.preventDefault) event.preventDefault(); // DOM Level 2
   else event.returnValue = false; // IE 5+

   if (document.removeEventListener) { // DOM Level 2
      document.removeEventListener("mousemove", SocialCalc.ProcessEditorColsizeMouseMove, true);
      document.removeEventListener("mouseup", SocialCalc.ProcessEditorColsizeMouseUp, true);
   } else if (editor.toplevel.detachEvent) { // IE
      editor.toplevel.detachEvent("onlosecapture", SocialCalc.ProcessEditorColsizeMouseUp);
      editor.toplevel.detachEvent("onmouseup", SocialCalc.ProcessEditorColsizeMouseUp);
      editor.toplevel.detachEvent("onmousemove", SocialCalc.ProcessEditorColsizeMouseMove);
      editor.toplevel.releaseCapture();
   }

   let newsize = (editor.context.colwidth[mouseinfo.mouseresizecolnum] - 0) + (clientX - mouseinfo.mousedownclientx);
   if (newsize < SocialCalc.Constants.defaultMinimumColWidth) newsize = SocialCalc.Constants.defaultMinimumColWidth;

   editor.EditorScheduleSheetCommands(`set ${mouseinfo.mouseresizecol} width ${newsize}`, true, false);

   if (editor.timeout) window.clearTimeout(editor.timeout);
   editor.timeout = window.setTimeout(SocialCalc.FinishColsize, 1); // wait - Firefox 2 has a bug otherwise with next mousedown

   return false;
};
/**
 * @function FinishColsize
 * @memberof SocialCalc
 * @description Finishes column resizing operation and cleans up
 */
SocialCalc.FinishColsize = function() {
   const mouseinfo = SocialCalc.EditorMouseInfo;
   const editor = mouseinfo.editor;
   if (!editor) return;

   editor.toplevel.removeChild(mouseinfo.mouseresizedisplay);
   mouseinfo.mouseresizedisplay = null;

//   editor.FitToEditTable();
//   editor.EditorRenderSheet();
//   editor.SchedulePositionCalculations();

   mouseinfo.editor = null;

   return;
};

/**
 * @namespace AutoRepeatInfo
 * @memberof SocialCalc
 * @description Handle auto-repeat of dragging the cursor into the borders of the sheet
 */
SocialCalc.AutoRepeatInfo = {
   /**
    * @member {Object|null} timer - Timer object for repeating
    */
   timer: null,
   
   /**
    * @member {Object|null} mouseinfo - Result from SocialCalc.GridMousePosition
    */
   mouseinfo: null,
   
   /**
    * @member {number} repeatinterval - Milliseconds to wait between repeats
    */
   repeatinterval: 1000,
   
   /**
    * @member {Object|null} editor - Editor object to use when it repeats
    */
   editor: null,
   
   /**
    * @member {Function|null} repeatcallback - Used instead of default when repeating (e.g., for cellhandles)
    * @description Called as: repeatcallback(newcoord, direction)
    */
   repeatcallback: null
};

/**
 * @function SetDragAutoRepeat
 * @memberof SocialCalc
 * @description Control auto-repeat. If mouseinfo==null, cancel.
 * @param {Object} editor - The editor instance
 * @param {Object|null} mouseinfo - Mouse information object, null to cancel
 * @param {Function} callback - Callback function for repeat operations
 */
SocialCalc.SetDragAutoRepeat = function(editor, mouseinfo, callback) {
   const repeatinfo = SocialCalc.AutoRepeatInfo;
   let coord, direction;

   repeatinfo.repeatcallback = callback; // null in regular case

   if (!mouseinfo) { // cancel
      if (repeatinfo.timer) { // If was repeating, stop
         window.clearTimeout(repeatinfo.timer); // cancel timer
         repeatinfo.timer = null;
      }
      repeatinfo.mouseinfo = null;
      return; // done
   }

   repeatinfo.editor = editor;

   if (repeatinfo.mouseinfo) { // check for change while repeating
      if (mouseinfo.rowheader || mouseinfo.rowfooter) {
         if (mouseinfo.row !== repeatinfo.mouseinfo.row) { // changed row while dragging sidewards
            coord = SocialCalc.crToCoord(editor.ecell.col, mouseinfo.row); // change to it
            if (repeatinfo.repeatcallback) {
               if (mouseinfo.row < repeatinfo.mouseinfo.row) {
                  direction = "left";
               } else if (mouseinfo.row > repeatinfo.mouseinfo.row) {
                  direction = "right";
               } else {
                  direction = "";
               }
               repeatinfo.repeatcallback(coord, direction);
            } else {
               editor.MoveECell(coord);
               editor.MoveECell(coord);
               editor.RangeExtend();
               editor.EditorMouseRange(coord);
            }
         }            
      } else if (mouseinfo.colheader || mouseinfo.colfooter) {
         if (mouseinfo.col !== repeatinfo.mouseinfo.col) { // changed col while dragging vertically
            coord = SocialCalc.crToCoord(mouseinfo.col, editor.ecell.row); // change to it
            if (repeatinfo.repeatcallback) {
               if (mouseinfo.row < repeatinfo.mouseinfo.row) {
                  direction = "left";
               } else if (mouseinfo.row > repeatinfo.mouseinfo.row) {
                  direction = "right";
               } else {
                  direction = "";
               }
               repeatinfo.repeatcallback(coord, direction);
            } else {
               editor.MoveECell(coord);
               editor.RangeExtend();
               editor.EditorMouseRange(coord);
            }
         }            
      }
   }

   repeatinfo.mouseinfo = mouseinfo;

   if (mouseinfo.distance < 5) repeatinfo.repeatinterval = 333;
   else if (mouseinfo.distance < 10) repeatinfo.repeatinterval = 250;
   else if (mouseinfo.distance < 25) repeatinfo.repeatinterval = 100;
   else if (mouseinfo.distance < 35) repeatinfo.repeatinterval = 75;
   else { // too far - stop repeating
      if (repeatinfo.timer) { // if repeating, cancel it
         window.clearTimeout(repeatinfo.timer); // cancel timer
         repeatinfo.timer = null;
      }
      return;
   }

   if (!repeatinfo.timer) { // start if not already running
      repeatinfo.timer = window.setTimeout(SocialCalc.DragAutoRepeat, repeatinfo.repeatinterval);
   }

   return;
};

/**
 * @function DragAutoRepeat
 * @memberof SocialCalc
 * @description Handles the auto-repeat functionality for dragging operations
 */
SocialCalc.DragAutoRepeat = function() {
   const repeatinfo = SocialCalc.AutoRepeatInfo;
   const mouseinfo = repeatinfo.mouseinfo;

   let direction, coord, cr;

   if (mouseinfo.rowheader) direction = "left";
   else if (mouseinfo.rowfooter) direction = "right";
   else if (mouseinfo.colheader) direction = "up";
   else if (mouseinfo.colfooter) direction = "down";

   if (repeatinfo.repeatcallback) {
      cr = SocialCalc.coordToCr(repeatinfo.editor.ecell.coord);
      if (direction === "left" && cr.col > 1) cr.col--;
      else if (direction === "right") cr.col++;
      else if (direction === "up" && cr.row > 1) cr.row--;
      else if (direction === "down") cr.row++;
      coord = SocialCalc.crToCoord(cr.col, cr.row);
      repeatinfo.repeatcallback(coord, direction);
   } else {
      coord = repeatinfo.editor.MoveECellWithKey(`[a${direction}]shifted`);
      if (coord) repeatinfo.editor.EditorMouseRange(coord);
   }

   repeatinfo.timer = window.setTimeout(SocialCalc.DragAutoRepeat, repeatinfo.repeatinterval);
};

/**
 * @function ProcessEditorDblClick
 * @memberof SocialCalc
 * @description Handles double-click events on the editor
 * @param {Event} e - The double-click event
 */
SocialCalc.ProcessEditorDblClick = function(e) {
   let editor, result, coord, textarea, wval, range;

   const event = e || window.event;

   const viewport = SocialCalc.GetViewportInfo();
   const clientX = event.clientX + viewport.horizontalScroll;
   const clientY = event.clientY + viewport.verticalScroll;

   const mouseinfo = SocialCalc.EditorMouseInfo;
   let ele = event.target || event.srcElement; // source object is often within what we want
   let mobj;

   if (mouseinfo.ignore) return; // ignore this

   for (mobj = null; !mobj && ele; ele = ele.parentNode) { // go up tree looking for one of our elements
      mobj = SocialCalc.LookupElement(ele, mouseinfo.registeredElements);
   }
   if (!mobj) {
      mouseinfo.editor = null;
      return; // not one of our elements
   }

   editor = mobj.editor;

   result = SocialCalc.GridMousePosition(editor, clientX, clientY);
   if (!result || !result.coord) return; // not within cell area - ignore

   mouseinfo.editor = editor; // remember for later
   mouseinfo.element = ele;
   range = editor.range;

   const sheetobj = editor.context.sheetobj;

   switch (editor.state) {
      case "start":
         SocialCalc.EditorOpenCellEdit(editor);
         break;

      case "input":
         break;

      default:
         break;
   }

   if (event.stopPropagation) event.stopPropagation(); // DOM Level 2
   else event.cancelBubble = true; // IE 5+
   if (event.preventDefault) event.preventDefault(); // DOM Level 2
   else event.returnValue = false; // IE 5+

   return;
};
/**
 * @function EditorOpenCellEdit
 * @memberof SocialCalc
 * @description Opens a cell for editing
 * @param {Object} editor - The table editor instance
 * @returns {boolean} Returns true if editing couldn't be started
 */
SocialCalc.EditorOpenCellEdit = function(editor) {
   let wval;

   if (!editor.ecell) return true; // no ecell
   if (!editor.inputBox) return true; // no input box, so no editing (happens on noEdit)
   if (SocialCalc.Callbacks && SocialCalc.Callbacks.IsCellEditable) {
      if (!SocialCalc.Callbacks.IsCellEditable(editor)) {
         return true;
      } 
   }
   if (editor.inputBox.element.disabled) return true; // multi-line: ignore
   if (editor.inputBox.element.style.display === 'none') {
      for (const f in editor.StatusCallback) {
         editor.StatusCallback[f].func(editor, "editecell", null, editor.StatusCallback[f].params);
      }
      return true; // no inputBox display, so no editing
   }
   editor.inputBox.ShowInputBox(true);
   editor.inputBox.Focus();
   if (SocialCalc.HasTouch) {
      editor.state = "input";
   } else {
      editor.state = "inputboxdirect";
   }
   editor.inputBox.SetText("");
   editor.inputBox.DisplayCellContents();
   editor.inputBox.Select("end");
   wval = editor.workingvalues;
   wval.partialexpr = "";
   wval.ecoord = editor.ecell.coord;
   wval.erow = editor.ecell.row;
   wval.ecol = editor.ecell.col;
   wval.startsheet = wval.currentsheet;
   wval.startsheetid = wval.currentsheetid;

   return;
};

/**
 * @function EditorProcessKey
 * @memberof SocialCalc
 * @description Processes keyboard input for the editor
 * @param {Object} editor - The table editor instance
 * @param {string} ch - The character or key name
 * @param {Event} e - The keyboard event
 * @returns {boolean} Returns false to prevent default behavior, true to allow it
 */
SocialCalc.EditorProcessKey = function(editor, ch, e) {
   let result, cell, cellobj, valueinfo, fch, coord, inputtext, f;

   const sheetobj = editor.context.sheetobj;
   const wval = editor.workingvalues;
   const range = editor.range;

   if (typeof ch !== "string") ch = "";

   switch (editor.state) {
      case "start":
         if (e.shiftKey && ch.substr(0, 2) === "[a") {
            ch = `${ch}shifted`;
         }
         if (ch === "[enter]") ch = "[adown]";
         if (ch === "[tab]") ch = e.shiftKey ? "[aleft]" : "[aright]";
         if (ch.substr(0, 2) === "[a" || ch.substr(0, 3) === "[pg" || ch === "[home]") {
            result = editor.MoveECellWithKey(ch);
            return !result;
         }
         if (ch === "[del]" || ch === "[backspace]") {
            if (!editor.noEdit) {
               editor.EditorApplySetCommandsToRange("empty", "");
            }
            if (SocialCalc.Callbacks && SocialCalc.Callbacks.IsCellEditable) {
               if (!SocialCalc.Callbacks.IsCellEditable(editor)) {
                  return true;
               } 
            }

            break;
         }
         if (ch === "[esc]") {
            if (range.hasrange) {
               editor.RangeRemove();
               editor.MoveECell(range.anchorcoord);
               for (f in editor.StatusCallback) {
                  editor.StatusCallback[f].func(editor, "specialkey", ch, editor.StatusCallback[f].params);
               }
            }
            return false;
         }

         if (ch === "[f2]") {
            if (editor.noEdit) return true;
            SocialCalc.EditorOpenCellEdit(editor);
            return false;
         }

         if ((ch.length > 1 && ch.substr(0, 1) === "[") || ch.length === 0) { // some control key
            if (editor.ctrlkeyFunction && ch.length > 0) {
               return editor.ctrlkeyFunction(editor, ch);
            } else {
               return true;
            }
         }
         if (!editor.ecell) return true; // no ecell
         if (!editor.inputBox) return true; // no inputBox so no editing
         if (editor.inputBox.element.style.display === 'none') {
            for (f in editor.StatusCallback) {
               editor.StatusCallback[f].func(editor, "editecell", ch, editor.StatusCallback[f].params);
            }
            return true; // no inputBox display, so no editing
         }
         if (SocialCalc.Callbacks && SocialCalc.Callbacks.IsCellEditable) {
            if (!SocialCalc.Callbacks.IsCellEditable(editor)) {
               return true;
            } 
         }

         editor.inputBox.element.disabled = false; // make sure editable
         editor.state = "input";
         editor.inputBox.ShowInputBox(true);
         editor.inputBox.Focus();
         editor.inputBox.SetText(ch);
         editor.inputBox.Select("end");
         wval.partialexpr = "";
         wval.ecoord = editor.ecell.coord;
         wval.erow = editor.ecell.row;
         wval.ecol = editor.ecell.col;
         wval.startsheet = wval.currentsheet;
         wval.startsheetid = wval.currentsheetid;
         editor.RangeRemove();
         break;

      case "input":
         inputtext = editor.inputBox.GetText(); // should not get here if no inputBox
         if (editor.inputBox.skipOne) return false; // ignore a key already handled
         if (ch === "[esc]" || ch === "[enter]" || ch === "[tab]" || (ch && ch.substr(0, 2) === "[a")) {
            if (("(+-*/,:!&<>=^".indexOf(inputtext.slice(-1)) >= 0 && inputtext.slice(0, 1) === "=") ||
                (inputtext === "=")) {
               wval.partialexpr = inputtext;
            }
            if (wval.partialexpr) { // if in pointing operation
               if (e.shiftKey && ch.substr(0, 2) === "[a") {
                  ch = `${ch}shifted`;
               }
               coord = editor.MoveECellWithKey(ch);
               if (coord) {
                  if (range.hasrange) {
                     editor.inputBox.SetText(`${wval.partialexpr}${SocialCalc.crToCoord(range.left, range.top)}:${SocialCalc.crToCoord(range.right, range.bottom)}`);
                  } else {
                     editor.inputBox.SetText(`${wval.partialexpr}${coord}`);
                  }
                  return false;
               }
            }
            editor.inputBox.Blur();
            editor.inputBox.ShowInputBox(false);
            editor.state = "start";
            editor.cellhandles.ShowCellHandles(true);
            if (ch !== "[esc]") {
               editor.EditorSaveEdit();
               if (editor.ecell.coord !== wval.ecoord) {
                  editor.MoveECell(wval.ecoord);
               }
               if (ch === "[enter]") ch = "[adown]";
               if (ch === "[tab]") ch = e.shiftKey ? "[aleft]" : "[aright]";
               if (ch.substr(0, 2) === "[a") {
                  editor.MoveECellWithKey(ch);
               }
            } else {
               editor.inputBox.DisplayCellContents();
               editor.RangeRemove();
               editor.MoveECell(wval.ecoord);
            }
            break;
         }
         if (wval.partialexpr && ch === "[backspace]") {
            editor.inputBox.SetText(wval.partialexpr);
            wval.partialexpr = "";
            editor.RangeRemove();
            editor.MoveECell(wval.ecoord);
            editor.inputBox.ShowInputBox(true); // make sure it's moved back if necessary
            return false;
         }
         if (ch === "[f2]") return false;
         if (range.hasrange) {
            editor.RangeRemove();
         }
         editor.MoveECell(wval.ecoord);
         if (wval.partialexpr) {
            editor.inputBox.ShowInputBox(true); // make sure it's moved back if necessary
            wval.partialexpr = ""; // not pointing
         }
         return true;

      case "inputboxdirect":
         inputtext = editor.inputBox.GetText(); // should not get here if no inputBox
         if (ch === "[esc]" || ch === "[enter]" || ch === "[tab]") {
            editor.inputBox.Blur();
            editor.inputBox.ShowInputBox(false);
            editor.state = "start";
            editor.cellhandles.ShowCellHandles(true);
            if (ch === "[esc]") {
               editor.inputBox.DisplayCellContents();
            } else {
               editor.EditorSaveEdit();
               if (editor.ecell.coord !== wval.ecoord) {
                  editor.MoveECell(wval.ecoord);
               }
               if (ch === "[enter]") ch = "[adown]";
               if (ch === "[tab]") ch = e.shiftKey ? "[aleft]" : "[aright]";
               if (ch.substr(0, 2) === "[a") {
                  editor.MoveECellWithKey(ch);
               }
            }
            break;
         }
         if (ch === "[f2]") return false;
         return true;

      case "skip-and-start":
         editor.state = "start";
         editor.cellhandles.ShowCellHandles(true);
         return false;

      default:
         return true;
   }

   return false;
};

/**
 * @function EditorAddToInput
 * @memberof SocialCalc
 * @description Adds text to the input box
 * @param {Object} editor - The table editor instance
 * @param {string} str - The string to add
 * @param {string} prefix - Optional prefix to add before the string
 * @returns {boolean} Returns true if operation couldn't be completed
 */
SocialCalc.EditorAddToInput = function(editor, str, prefix) {
   const wval = editor.workingvalues;

   if (editor.noEdit) return;

   if (SocialCalc.Callbacks && SocialCalc.Callbacks.IsCellEditable) {
      if (!SocialCalc.Callbacks.IsCellEditable(editor)) {
         return true;
      } 
   }

   switch (editor.state) {
      case "start":
         editor.state = "input";
         editor.inputBox.ShowInputBox(true);
         editor.inputBox.element.disabled = false; // make sure editable and overwrite old
         editor.inputBox.Focus();
         editor.inputBox.SetText(`${prefix || ""}${str}`);
         editor.inputBox.Select("end");
         wval.partialexpr = "";
         wval.ecoord = editor.ecell.coord;
         wval.erow = editor.ecell.row;
         wval.ecol = editor.ecell.col;
         wval.startsheet = wval.currentsheet;
         wval.startsheetid = wval.currentsheetid;
         editor.RangeRemove();
         break;

      case "input":
      case "inputboxdirect":
         editor.inputBox.element.focus();
         if (wval.partialexpr) {
            editor.inputBox.SetText(wval.partialexpr);
            wval.partialexpr = "";
            editor.RangeRemove();
            editor.MoveECell(wval.ecoord);
         }
         editor.inputBox.SetText(`${editor.inputBox.GetText()}${str}`);
         break;

      default:
         break;
   }
};
/**
 * @function EditorDisplayCellContents
 * @memberof SocialCalc
 * @description Displays the contents of the current cell
 * @param {Object} editor - The table editor instance
 */
SocialCalc.EditorDisplayCellContents = function(editor) {
   if (editor.inputBox) editor.inputBox.DisplayCellContents();
};

/**
 * @function EditorSaveEdit
 * @memberof SocialCalc
 * @description Saves the current edit operation
 * @param {Object} editor - The table editor instance
 * @param {string} text - Optional text to save (if not provided, gets from input box)
 * @returns {boolean} Returns true if operation couldn't be completed
 */
SocialCalc.EditorSaveEdit = function(editor, text) {
   let result, cell, valueinfo, fch, type, value, oldvalue, cmdline;

   const sheetobj = editor.context.sheetobj;
   const wval = editor.workingvalues;

   if (SocialCalc.Callbacks && SocialCalc.Callbacks.IsCellEditable) {
      if (!SocialCalc.Callbacks.IsCellEditable(editor)) {
         return true;
      } 
   }

   type = "text t";
   value = typeof text === "string" ? text : editor.inputBox.GetText(); // either explicit or from input box

   oldvalue = `${SocialCalc.GetCellContents(sheetobj, wval.ecoord)}`;
   if (value === oldvalue) { // no change
      return;
   }

   if (!SocialCalc.Callbacks.CheckConstraints(editor, value)) {
      return true;
   }

   fch = value.charAt(0);
   if (fch === "=" && value.indexOf("\n") === -1) {
      type = "formula";
      value = value.substring(1);
   } else if (fch === "'") {
      type = "text t";
      value = value.substring(1);
   } else if (value.length === 0) {
      type = "empty";
   } else {
      valueinfo = SocialCalc.DetermineValueType(value);
      if (valueinfo.type === "n" && value === `${valueinfo.value}`) { // see if don't need "constant"
         type = "value n";
      } else if (valueinfo.type.charAt(0) === "t") {
         type = `text ${valueinfo.type}`;
      } else if (valueinfo.type === "") {
         type = "text t";
      } else {
         type = `constant ${valueinfo.type} ${valueinfo.value}`;
      }
   }

   if (type.charAt(0) === "t") { // text
      value = SocialCalc.encodeForSave(value); // newlines, :, and \ are escaped
   }

   // if startsheet different from currentsheet, switch to start sheet
   // for the save to take effect
   if (SocialCalc.WorkBook && (wval.currentsheet !== wval.startsheet)) {
      const control = SocialCalc.GetCurrentWorkBookControl();
      const cmdstr = `activatesheet ${wval.startsheetid}`;
      control.ExecuteWorkBookControlCommand({ 
         cmdtype: "wcmd", 
         id: "0", 
         cmdstr: cmdstr
      }, false);
   }

   cmdline = `set ${wval.ecoord} ${type} ${value}`;
   editor.EditorScheduleSheetCommands(cmdline, true, false);

   return;
};

/**
 * @function EditorApplySetCommandsToRange
 * @memberof SocialCalc
 * @description Takes ecell or range and does a "set" command with cmd
 * @param {Object} editor - The table editor instance
 * @param {string} cmd - The command to apply
 */
SocialCalc.EditorApplySetCommandsToRange = function(editor, cmd) {
   let cell, row, col, line, errortext;

   const sheetobj = editor.context.sheetobj;
   const ecell = editor.ecell;
   const range = editor.range;

   if (range.hasrange) {
      const coord = `${SocialCalc.crToCoord(range.left, range.top)}:${SocialCalc.crToCoord(range.right, range.bottom)}`;
      line = `set ${coord} ${cmd}`;
      errortext = editor.EditorScheduleSheetCommands(line, true, false);
   } else {
      line = `set ${ecell.coord} ${cmd}`;
      errortext = editor.EditorScheduleSheetCommands(line, true, false);
   }

   editor.DisplayCellContents();
};

/**
 * @function EditorProcessMouseWheel
 * @memberof SocialCalc
 * @description Processes mouse wheel events for scrolling
 * @param {Event} event - The mouse wheel event
 * @param {number} delta - The wheel delta value
 * @param {Object} mousewheelinfo - Mouse wheel information
 * @param {Object} wobj - Wheel object containing editor reference
 */
SocialCalc.EditorProcessMouseWheel = function(event, delta, mousewheelinfo, wobj) {
   if (wobj.functionobj.editor.busy) return; // ignore if busy

   if (delta > 0) {
      wobj.functionobj.editor.ScrollRelative(true, -1);
   }
   if (delta < 0) {
      wobj.functionobj.editor.ScrollRelative(true, +1);
   }
};

/**
 * @function GridMousePosition
 * @memberof SocialCalc
 * @description Returns an object with row and col numbers and coord (spans handled for coords),
 * and rowheader/colheader true if in header (where coord will be undefined).
 * If in colheader, will return coltoresize if on appropriate place in col header.
 * Also, there is rowfooter (on right) and colfooter (on bottom).
 * In row/col header/footer, returns "distance" as pixels over the edge.
 * @param {Object} editor - The table editor instance
 * @param {number} clientX - The client X coordinate
 * @param {number} clientY - The client Y coordinate
 * @returns {Object|null} Position information object or null
 */
SocialCalc.GridMousePosition = function(editor, clientX, clientY) { 
   let row, col, colpane;
   const result = {};

   for (row = 1; row < editor.rowpositions.length; row++) {
      if (!editor.rowheight[row]) continue; // not rendered yet -- may be above or below us
      if (editor.rowpositions[row] + editor.rowheight[row] > clientY) {
         break;
      }
   }
   for (col = 1; col < editor.colpositions.length; col++) {
      if (!editor.colwidth[col]) continue;
      if (editor.colpositions[col] + editor.colwidth[col] > clientX) {
         break;
      }
   }

   result.row = row;
   result.col = col;

   if (editor.headposition) {
      if (clientX < editor.headposition.left && clientX >= editor.gridposition.left) {
         result.rowheader = true;
         result.distance = editor.headposition.left - clientX;
         return result;
      } else if (clientY < editor.headposition.top && clientY > editor.gridposition.top) { // > because of sizing row
         result.colheader = true;
         result.distance = editor.headposition.top - clientY;
         result.coltoresize = col - (editor.colpositions[col] + editor.colwidth[col] / 2 > clientX ? 1 : 0) || 1;
         for (colpane = 0; colpane < editor.context.colpanes.length; colpane++) {
            if (result.coltoresize >= editor.context.colpanes[colpane].first &&
                result.coltoresize <= editor.context.colpanes[colpane].last) { // visible column
               return result;
            }
         }
         delete result.coltoresize;
         return result;
      } else if (clientX >= editor.verticaltablecontrol.controlborder) {
         result.rowfooter = true;
         result.distance = clientX - editor.verticaltablecontrol.controlborder;
         return result;
      } else if (clientY >= editor.horizontaltablecontrol.controlborder) {
         result.colfooter = true;
         result.distance = clientY - editor.horizontaltablecontrol.controlborder;
         return result;
      } else if (clientX < editor.gridposition.left) {
         result.rowheader = true;
         result.distance = editor.headposition.left - clientX;
         return result;
      } else if (clientY <= editor.gridposition.top) {
         result.colheader = true;
         result.distance = editor.headposition.top - clientY;
         return result;
      } else {
         result.coord = SocialCalc.crToCoord(result.col, result.row);
         if (editor.context.cellskip[result.coord]) { // handle skipped cells
            result.coord = editor.context.cellskip[result.coord];
         }
         return result;
      }
   }

   return null;
};
/**
 * @function GetEditorCellElement
 * @memberof SocialCalc
 * @description Returns an object with element, the table cell element in the DOM that corresponds to row and column,
 * as well as rowpane and colpane, the panes with the cell.
 * If no such element, then returns null.
 * @param {Object} editor - The table editor instance
 * @param {number} row - The row number
 * @param {number} col - The column number
 * @returns {Object|null} Object containing element, rowpane, and colpane, or null if not found
 */
SocialCalc.GetEditorCellElement = function(editor, row, col) {
   let rowpane, colpane, c, coord;
   let rowindex = 0;
   let colindex = 0;

   for (rowpane = 0; rowpane < editor.context.rowpanes.length; rowpane++) {
      if (row >= editor.context.rowpanes[rowpane].first && row <= editor.context.rowpanes[rowpane].last) {
         for (colpane = 0; colpane < editor.context.colpanes.length; colpane++) {
            if (col >= editor.context.colpanes[colpane].first && col <= editor.context.colpanes[colpane].last) {
               rowindex += row - editor.context.rowpanes[rowpane].first + 2;
               for (c = editor.context.colpanes[colpane].first; c <= col; c++) {
                  coord = editor.context.cellskip[SocialCalc.crToCoord(c, row)];
                  if (!coord || !editor.context.CoordInPane(coord, rowpane, colpane)) // don't count col-spanned cells
                     colindex++;
               }
               return {
                  element: editor.griddiv.firstChild.lastChild.childNodes[rowindex].childNodes[colindex],
                  rowpane: rowpane, 
                  colpane: colpane
               };
            }
            for (c = editor.context.colpanes[colpane].first; c <= editor.context.colpanes[colpane].last; c++) {
               coord = editor.context.cellskip[SocialCalc.crToCoord(c, row)];
               if (!coord || !editor.context.CoordInPane(coord, rowpane, colpane)) // don't count col-spanned cells
                  colindex++;
            }
            colindex += 1;
         }
      }
      rowindex += editor.context.rowpanes[rowpane].last - editor.context.rowpanes[rowpane].first + 1 + 1;
   }

   return null;
};

/**
 * @function MoveECellWithKey
 * @memberof SocialCalc
 * @description Processes an arrow key, etc., moving the edit cell.
 * If not a movement key, returns null.
 * @param {Object} editor - The table editor instance
 * @param {string} ch - The key character/command
 * @returns {string|null} New cell coordinate or null if not a movement key
 */
SocialCalc.MoveECellWithKey = function(editor, ch) {
   let coord, row, col, cell;
   let shifted = false;

   if (!editor.ecell) {
      return null;
   }

   if (ch.slice(-7) === "shifted") {
      ch = ch.slice(0, -7);
      shifted = true;
   }

   row = editor.ecell.row;
   col = editor.ecell.col;
   cell = editor.context.sheetobj.cells[editor.ecell.coord];

   switch (ch) {
      case "[adown]":
         row += (cell && cell.rowspan) || 1;
         break;
      case "[aup]":
         row--;
         break;
      case "[pgdn]":
         row += editor.pageUpDnAmount - 1 + ((cell && cell.rowspan) || 1);
         break;
      case "[pgup]":
         row -= editor.pageUpDnAmount;
         break;
      case "[aright]":
         col += (cell && cell.colspan) || 1;
         break;
      case "[aleft]":
         col--;
         break;
      case "[home]":
         row = 1;
         col = 1;
         break;
      default:
         return null;
   }

   if (!editor.range.hasrange) {
      if (shifted)
         editor.RangeAnchor();
   }

   coord = editor.MoveECell(SocialCalc.crToCoord(col, row));

   if (editor.range.hasrange) {
      if (shifted)
         editor.RangeExtend();
      else
         editor.RangeRemove();
   }

   return coord;
};

/**
 * @function MoveECell
 * @memberof SocialCalc
 * @description Takes a coordinate and returns the new edit cell coordinate (which may be
 * different if newecell is covered by a span).
 * @param {Object} editor - The table editor instance
 * @param {string} newcell - The new cell coordinate
 * @returns {string} The actual new cell coordinate
 */
SocialCalc.MoveECell = function(editor, newcell) {
   let cell, f;

   const highlights = editor.context.highlights;

   if (editor.ecell) {
      if (editor.ecell.coord === newcell) return newcell; // already there - don't do anything and don't tell anybody

      if (SocialCalc.Callbacks.broadcast) {
         SocialCalc.Callbacks.broadcast('ecell', { 
            original: editor.ecell.coord, 
            ecell: newcell 
         });
      }

      cell = SocialCalc.GetEditorCellElement(editor, editor.ecell.row, editor.ecell.col);
      delete highlights[editor.ecell.coord];
      if (editor.range2.hasrange &&
         editor.ecell.row >= editor.range2.top && editor.ecell.row <= editor.range2.bottom &&
         editor.ecell.col >= editor.range2.left && editor.ecell.col <= editor.range2.right) {
         highlights[editor.ecell.coord] = "range2";
      }
      editor.UpdateCellCSS(cell, editor.ecell.row, editor.ecell.col);
      editor.SetECellHeaders(""); // set to regular col/rowname styles
      editor.cellhandles.ShowCellHandles(false);
   } else if (SocialCalc.Callbacks.broadcast) {
      SocialCalc.Callbacks.broadcast('ecell', { ecell: newcell });
   }

   newcell = editor.context.cellskip[newcell] || newcell;
   editor.ecell = SocialCalc.coordToCr(newcell);
   editor.ecell.coord = newcell;
   cell = SocialCalc.GetEditorCellElement(editor, editor.ecell.row, editor.ecell.col);
   highlights[newcell] = "cursor";

   for (f in editor.MoveECellCallback) { // let others know
      editor.MoveECellCallback[f](editor);
   }

   editor.UpdateCellCSS(cell, editor.ecell.row, editor.ecell.col);
   editor.SetECellHeaders("selected");

   for (f in editor.StatusCallback) { // let status line, etc., know
      editor.StatusCallback[f].func(editor, "moveecell", newcell, editor.StatusCallback[f].params);
   }

   if (editor.busy) {
      editor.ensureecell = true; // wait for when not busy
   } else {
      editor.ensureecell = false;
      editor.EnsureECellVisible();
   }

   return newcell;
};

/**
 * @function EnsureECellVisible
 * @memberof SocialCalc
 * @description Ensures the current edit cell is visible by scrolling if necessary
 * @param {Object} editor - The table editor instance
 */
SocialCalc.EnsureECellVisible = function(editor) {
   let vamount = 0;
   let hamount = 0;

   if (editor.ecell.row > editor.lastnonscrollingrow) {
      if (editor.ecell.row < editor.firstscrollingrow) {
         vamount = editor.ecell.row - editor.firstscrollingrow;
      } else if (editor.ecell.row > editor.lastvisiblerow) {
         vamount = editor.ecell.row - editor.lastvisiblerow;
      }
   }
   
   if (editor.ecell.col > editor.lastnonscrollingcol) {
      if (editor.ecell.col < editor.firstscrollingcol) {
         hamount = editor.ecell.col - editor.firstscrollingcol;
      } else if (editor.ecell.col > editor.lastvisiblecol) {
         hamount = editor.ecell.col - editor.lastvisiblecol;
      }
   }

   if (vamount !== 0 || hamount !== 0) {
      editor.ScrollRelativeBoth(vamount, hamount);
   } else {
      editor.cellhandles.ShowCellHandles(true);
   }
};
/**
 * @function ReplaceCell
 * @memberof SocialCalc
 * @description Replaces a cell element with newly rendered content
 * @param {Object} editor - The table editor instance
 * @param {Object} cell - The cell object containing the element to replace
 * @param {number} row - The row number
 * @param {number} col - The column number
 */
SocialCalc.ReplaceCell = function(editor, cell, row, col) {
   let newelement, a;
   if (!cell) return;
   newelement = editor.context.RenderCell(row, col, cell.rowpane, cell.colpane, true, null);
   if (newelement) {
      // Don't use a real element and replaceChild, which seems to have focus issues with IE, Firefox, and speed issues
      cell.element.innerHTML = newelement.innerHTML;
      cell.element.style.cssText = "";
      cell.element.className = newelement.className;
      for (a in newelement.style) {
         if (newelement.style[a] !== "cssText")
            cell.element.style[a] = newelement.style[a];
      }
   }
};

/**
 * @function UpdateCellCSS
 * @memberof SocialCalc
 * @description Updates the CSS styling of a cell element
 * @param {Object} editor - The table editor instance
 * @param {Object} cell - The cell object containing the element to update
 * @param {number} row - The row number
 * @param {number} col - The column number
 */
SocialCalc.UpdateCellCSS = function(editor, cell, row, col) {
   let newelement, a;
   if (!cell) return;
   newelement = editor.context.RenderCell(row, col, cell.rowpane, cell.colpane, true, null);
   if (newelement) {
      cell.element.style.cssText = "";
      cell.element.className = newelement.className;
      for (a in newelement.style) {
         if (newelement.style[a] !== "cssText")
            cell.element.style[a] = newelement.style[a];
      }
   }
};

/**
 * @function SetECellHeaders
 * @memberof SocialCalc
 * @description Sets the styling for row and column headers of the current edit cell
 * @param {Object} editor - The table editor instance
 * @param {string} selected - The selection state ("selected" or empty string)
 */
SocialCalc.SetECellHeaders = function(editor, selected) {
   const ecell = editor.ecell;
   const context = editor.context;

   let rowpane, colpane, first, last;
   let rowindex = 0;
   let colindex = 0;
   let headercell;

   if (!ecell) return;

   for (rowpane = 0; rowpane < context.rowpanes.length; rowpane++) {
      first = context.rowpanes[rowpane].first;
      last = context.rowpanes[rowpane].last;
      if (ecell.row >= first && ecell.row <= last) {
         headercell = editor.fullgrid.childNodes[1].childNodes[2 + rowindex + ecell.row - first].childNodes[0];
         if (headercell) {
            if (context.classnames) headercell.className = context.classnames[`${selected}rowname`];
            if (context.explicitStyles) headercell.style.cssText = context.explicitStyles[`${selected}rowname`];
            headercell.style.verticalAlign = "top"; // to get around Safari making top of centered row number be
                                                    // considered top of row (and can't get <row> position in Safari)
         }
      }
      rowindex += last - first + 1 + 1;
   }

   for (colpane = 0; colpane < context.colpanes.length; colpane++) {
      first = context.colpanes[colpane].first;
      last = context.colpanes[colpane].last;
      if (ecell.col >= first && ecell.col <= last) {
         headercell = editor.fullgrid.childNodes[1].childNodes[1].childNodes[1 + colindex + ecell.col - first];
         if (headercell) {
            if (context.classnames) headercell.className = context.classnames[`${selected}colname`];
            if (context.explicitStyles) headercell.style.cssText = context.explicitStyles[`${selected}colname`];
         }
      }
      colindex += last - first + 1 + 1;
   }
};

/**
 * @function RangeAnchor
 * @memberof SocialCalc
 * @description Sets the anchor of a range to ecoord (or ecell if missing)
 * @param {Object} editor - The table editor instance
 * @param {string} ecoord - The coordinate to anchor the range to (optional)
 */
SocialCalc.RangeAnchor = function(editor, ecoord) {
   if (editor.range.hasrange) {
      editor.RangeRemove();
   }

   editor.RangeExtend(ecoord);
};

/**
 * @function RangeExtend
 * @memberof SocialCalc
 * @description Sets the other corner of the range to ecoord or, if missing, ecell
 * @param {Object} editor - The table editor instance
 * @param {string} ecoord - The coordinate to extend the range to (optional)
 */
SocialCalc.RangeExtend = function(editor, ecoord) {
   let a, cell, cr, coord, row, col, f;

   const highlights = editor.context.highlights;
   const range = editor.range;
   const range2 = editor.range2;

   let ecell;
   if (ecoord) {
      ecell = SocialCalc.coordToCr(ecoord);
      ecell.coord = ecoord;
   } else {
      ecell = editor.ecell;
   }

   if (!ecell) return; // just in case

   if (!range.hasrange) { // called without RangeAnchor...
      range.anchorcoord = ecell.coord;
      range.anchorrow = ecell.row;
      range.top = ecell.row;
      range.bottom = ecell.row;
      range.anchorcol = ecell.col;
      range.left = ecell.col;
      range.right = ecell.col;
      range.hasrange = true;
   }

   if (range.anchorrow < ecell.row) {
      range.top = range.anchorrow;
      range.bottom = ecell.row;
   } else {
      range.top = ecell.row;
      range.bottom = range.anchorrow;
   }
   if (range.anchorcol < ecell.col) {
      range.left = range.anchorcol;
      range.right = ecell.col;
   } else {
      range.left = ecell.col;
      range.right = range.anchorcol;
   }

   for (coord in highlights) {
      switch (highlights[coord]) {
         case "range":
            highlights[coord] = "unrange";
            break;
         case "range2":
            highlights[coord] = "unrange2";
            break;
      }
   }

   for (row = range.top; row <= range.bottom; row++) {
      for (col = range.left; col <= range.right; col++) {
         coord = SocialCalc.crToCoord(col, row);
         switch (highlights[coord]) {
            case "unrange":
               highlights[coord] = "range";
               break;
            case "cursor":
               break;
            case "unrange2":
            default:
               highlights[coord] = "newrange";
               break;
         }
      }
   }

   for (row = range2.top; range2.hasrange && row <= range2.bottom; row++) {
      for (col = range2.left; col <= range2.right; col++) {
         coord = SocialCalc.crToCoord(col, row);
         switch (highlights[coord]) {
            case "unrange2":
               highlights[coord] = "range2";
               break;
            case "range":
            case "newrange":
            case "cursor":
               break;
            default:
               highlights[coord] = "newrange2";
               break;
         }
      }
   }

   for (coord in highlights) {
      switch (highlights[coord]) {
         case "unrange":
            delete highlights[coord];
            break;
         case "newrange":
            highlights[coord] = "range";
            break;
         case "newrange2":
            highlights[coord] = "range2";
            break;
         case "range":
         case "range2":
         case "cursor":
            continue;
      }

      cr = SocialCalc.coordToCr(coord);
      cell = SocialCalc.GetEditorCellElement(editor, cr.row, cr.col);
      editor.UpdateCellCSS(cell, cr.row, cr.col);
   }

   for (f in editor.RangeChangeCallback) { // let others know
      editor.RangeChangeCallback[f](editor);
   }

   // create range/coord string and do status callback
   coord = SocialCalc.crToCoord(editor.range.left, editor.range.top);
   if (editor.range.left !== editor.range.right || editor.range.top !== editor.range.bottom) { // more than one cell
      coord += `:${SocialCalc.crToCoord(editor.range.right, editor.range.bottom)}`;
   }
   for (f in editor.StatusCallback) {
      editor.StatusCallback[f].func(editor, "rangechange", coord, editor.StatusCallback[f].params);
   }

   return;
};
/**
 * @function RangeRemove
 * @memberof SocialCalc
 * @description Turns off the range
 * @param {Object} editor - The table editor instance
 */
SocialCalc.RangeRemove = function(editor) {
   let cell, cr, coord, row, col, f;

   const highlights = editor.context.highlights;
   const range = editor.range;
   const range2 = editor.range2;

   if (!range.hasrange && !range2.hasrange) return;

   for (row = range2.top; range2.hasrange && row <= range2.bottom; row++) {
      for (col = range2.left; col <= range2.right; col++) {
         coord = SocialCalc.crToCoord(col, row);
         switch (highlights[coord]) {
            case "range":
               highlights[coord] = "newrange2";
               break;
            case "range2":
            case "cursor":
               break;
            default:
               highlights[coord] = "newrange2";
               break;
         }
      }
   }

   for (coord in highlights) {
      switch (highlights[coord]) {
         case "range":
            delete highlights[coord];
            break;
         case "newrange2":
            highlights[coord] = "range2";
            break;
         case "cursor":
            continue;
      }
      cr = SocialCalc.coordToCr(coord);
      cell = SocialCalc.GetEditorCellElement(editor, cr.row, cr.col);
      editor.UpdateCellCSS(cell, cr.row, cr.col);
   }

   range.hasrange = false;

   for (f in editor.RangeChangeCallback) { // let others know
      editor.RangeChangeCallback[f](editor);
   }

   for (f in editor.StatusCallback) {
      editor.StatusCallback[f].func(editor, "rangechange", "", editor.StatusCallback[f].params);
   }

   return;
};

/**
 * @function Range2Remove
 * @memberof SocialCalc
 * @description Turns off the range2
 * @param {Object} editor - The table editor instance
 */
SocialCalc.Range2Remove = function(editor) {
   let cell, cr, coord, row, col, f;

   const highlights = editor.context.highlights;
   const range2 = editor.range2;

   if (!range2.hasrange) return;

   for (coord in highlights) {
      switch (highlights[coord]) {
         case "range2":
            delete highlights[coord];
            break;
         case "range":
         case "cursor":
            continue;
      }
      cr = SocialCalc.coordToCr(coord);
      cell = SocialCalc.GetEditorCellElement(editor, cr.row, cr.col);
      editor.UpdateCellCSS(cell, cr.row, cr.col);
   }

   range2.hasrange = false;

   return;
};

/**
 * @function FitToEditTable
 * @memberof SocialCalc
 * @description Figure out (through column width declarations and approximation of pixels per row)
 * how many rendered rows and columns you need to be at least a little larger than
 * the editor's editing area.
 * @param {Object} editor - The table editor instance
 */
SocialCalc.FitToEditTable = function(editor) {
   let colnum, colname, colwidth, totalwidth, totalrows, rowpane, needed;

   const context = editor.context;
   const sheetobj = context.sheetobj;
   const sheetcolattribs = sheetobj.colattribs;

   // Calculate column width data
   totalwidth = context.showRCHeaders ? context.rownamewidth - 0 : 0;
   let colpane;
   for (colpane = 0; colpane < context.colpanes.length - 1; colpane++) { // Get width of all but last pane
      for (colnum = context.colpanes[colpane].first; colnum <= context.colpanes[colpane].last; colnum++) {
         colname = SocialCalc.rcColname(colnum);
         colwidth = sheetobj.colattribs.width[colname] || sheetobj.attribs.defaultcolwidth || SocialCalc.Constants.defaultColWidth;
         if (colwidth === "blank" || colwidth === "auto") colwidth = "";
         totalwidth += (colwidth && ((colwidth - 0) > 0)) ? (colwidth - 0) : 10;
      }
   }

   for (colnum = context.colpanes[colpane].first; colnum <= 10000; colnum++) { //!!! max for safety, but makes that col max!!!
      colname = SocialCalc.rcColname(colnum);
      colwidth = sheetobj.colattribs.width[colname] || sheetobj.attribs.defaultcolwidth || SocialCalc.Constants.defaultColWidth;
      if (colwidth === "blank" || colwidth === "auto") colwidth = "";
      totalwidth += (colwidth && ((colwidth - 0) > 0)) ? (colwidth - 0) : 10;
      if (totalwidth > editor.tablewidth) break;
   }

   context.colpanes[colpane].last = colnum;

   // Calculate row height data
   totalrows = context.showRCHeaders ? 1 : 0;
   for (rowpane = 0; rowpane < context.rowpanes.length - 1; rowpane++) { // count all panes but last one
      totalrows += context.rowpanes[rowpane].last - context.rowpanes[rowpane].first + 1;
   }

   needed = editor.tableheight - totalrows * context.pixelsPerRow; // estimate amount needed

   context.rowpanes[rowpane].last = context.rowpanes[rowpane].first + Math.floor(needed / context.pixelsPerRow) + 1;
};

/**
 * @function CalculateEditorPositions
 * @memberof SocialCalc
 * @description Calculate the screen positions and other values of various editing elements
 * These values change and need to be recomputed when the pane first/last or cell contents change,
 * as well as new column widths, etc.
 * 
 * Note: Only call this after the grid has been rendered! You may have to wait for a timeout...
 * @param {Object} editor - The table editor instance
 */
SocialCalc.CalculateEditorPositions = function(editor) {
   let rowpane, colpane, i;

   editor.gridposition = SocialCalc.GetElementPosition(editor.griddiv);
   editor.headposition =
      SocialCalc.GetElementPosition(editor.griddiv.firstChild.lastChild.childNodes[2].childNodes[1]); // 3rd tr 2nd td

   editor.rowpositions = [];
   for (rowpane = 0; rowpane < editor.context.rowpanes.length; rowpane++) {
      editor.CalculateRowPositions(rowpane, editor.rowpositions, editor.rowheight);
   }
   for (i = 0; i < editor.rowpositions.length; i++) {
      if (editor.rowpositions[i] > editor.gridposition.top + editor.tableheight) break;
   }
   editor.lastvisiblerow = i - 1;

   editor.colpositions = [];
   for (colpane = 0; colpane < editor.context.colpanes.length; colpane++) {
      editor.CalculateColPositions(colpane, editor.colpositions, editor.colwidth);
   }
   for (i = 0; i < editor.colpositions.length; i++) {
      if (editor.colpositions[i] > editor.gridposition.left + editor.tablewidth) break;
   }
   editor.lastvisiblecol = i - 1;

   editor.firstscrollingrow = editor.context.rowpanes[editor.context.rowpanes.length - 1].first;
   editor.firstscrollingrowtop = editor.rowpositions[editor.firstscrollingrow] || editor.headposition.top;
   editor.lastnonscrollingrow = editor.context.rowpanes.length - 1 > 0 ?
         editor.context.rowpanes[editor.context.rowpanes.length - 2].last : 0;
   editor.firstscrollingcol = editor.context.colpanes[editor.context.colpanes.length - 1].first;
   editor.firstscrollingcolleft = editor.colpositions[editor.firstscrollingcol] || editor.headposition.left;
   editor.lastnonscrollingcol = editor.context.colpanes.length - 1 > 0 ?
         editor.context.colpanes[editor.context.colpanes.length - 2].last : 0;
   // Now do the table controls
   editor.verticaltablecontrol.ComputeTableControlPositions();
   editor.horizontaltablecontrol.ComputeTableControlPositions();
};

/**
 * @function ScheduleRender
 * @memberof SocialCalc
 * @description Do a series of timeouts to render the sheet, wait for background layout and
 * rendering by the browser, and then update editor visuals, sliders, etc.
 * @param {Object} editor - The table editor instance
 */
SocialCalc.ScheduleRender = function(editor) {
   if (editor.timeout) window.clearTimeout(editor.timeout); // in case called more than once, just use latest

   SocialCalc.EditorSheetStatusCallback(null, "schedrender", null, editor);
   SocialCalc.EditorStepInfo.editor = editor;
   editor.timeout = window.setTimeout(SocialCalc.DoRenderStep, 1);
};

/**
 * @function DoRenderStep
 * @memberof SocialCalc
 * @description Executes the rendering step in the scheduled render process
 */
SocialCalc.DoRenderStep = function() {
   const editor = SocialCalc.EditorStepInfo.editor;

   editor.timeout = null;

   editor.EditorRenderSheet();

   SocialCalc.EditorSheetStatusCallback(null, "renderdone", null, editor);

   SocialCalc.EditorSheetStatusCallback(null, "schedposcalc", null, editor);

   editor.timeout = window.setTimeout(SocialCalc.DoPositionCalculations, 1);
};

/**
 * @function SchedulePositionCalculations
 * @memberof SocialCalc
 * @description Schedules position calculations for the editor
 * @param {Object} editor - The table editor instance
 */
SocialCalc.SchedulePositionCalculations = function(editor) {
   SocialCalc.EditorStepInfo.editor = editor;

   SocialCalc.EditorSheetStatusCallback(null, "schedposcalc", null, editor);

   editor.timeout = window.setTimeout(SocialCalc.DoPositionCalculations, 1);
};

/**
 * @function DoPositionCalculations
 * @memberof SocialCalc
 * @description Update editor visuals, sliders, etc.
 * Note: Only call this after the DOM objects have been modified and rendered!
 */
SocialCalc.DoPositionCalculations = function() {
   const editor = SocialCalc.EditorStepInfo.editor;

   editor.timeout = null;

   let ok = false;
   try {
      editor.CalculateEditorPositions();
      ok = true;
   } catch (e) {}

   if (!ok) {
      if (typeof $ !== 'undefined') {
         $(window).trigger('resize');
         setTimeout(SocialCalc.DoPositionCalculations, 400);
      }
      return; /* Workaround IE6 partial-initialized-DOM bug */
   }

   editor.verticaltablecontrol.PositionTableControlElements();
   editor.horizontaltablecontrol.PositionTableControlElements();

   SocialCalc.EditorSheetStatusCallback(null, "doneposcalc", null, editor);

   if (editor.ensureecell && editor.ecell && !editor.deferredCommands.length) { // don't do if deferred cmd to execute
      editor.ensureecell = false;
      editor.EnsureECellVisible(); // this could cause another redisplay
   }

   editor.cellhandles.ShowCellHandles(true);

//!!! Need to now check to see if this positioned controls out of the editing area
//!!! (such as when there is a large wrapped cell and it pushes the pane boundary too far down).

   if (SocialCalc.Callbacks.broadcast) SocialCalc.Callbacks.broadcast('ask.ecell');
};

/**
 * @function CalculateRowPositions
 * @memberof SocialCalc
 * @description Calculates the screen positions of rows in a specific pane
 * @param {Object} editor - The table editor instance
 * @param {number} panenum - The pane number to calculate
 * @param {Array} positions - Array to store row positions
 * @param {Array} sizes - Array to store row sizes
 */
SocialCalc.CalculateRowPositions = function(editor, panenum, positions, sizes) {
   let toprow, rowpane, rownum, offset, trowobj, cellposition;

   const context = editor.context;
   const sheetobj = context.sheetobj;

   let tbodyobj;

   if (!context.showRCHeaders) throw new Error("Needs showRCHeaders=true");

   tbodyobj = editor.fullgrid.lastChild;

   // Calculate start of this pane as row in this table:
   toprow = 2;
   for (rowpane = 0; rowpane < panenum; rowpane++) {
      toprow += context.rowpanes[rowpane].last - context.rowpanes[rowpane].first + 2; // skip pane and spacing row
   }

   offset = 0;
   for (rownum = context.rowpanes[rowpane].first; rownum <= context.rowpanes[rowpane].last; rownum++) {
      trowobj = tbodyobj.childNodes[toprow + offset];
      offset++;
      cellposition = SocialCalc.GetElementPosition(trowobj.firstChild);

// Safari has problem: If a cell in the row is high, cell 1 is centered and it returns top of centered part 
// but if you get position of row element, it always returns the same value (not the row's)
// So we require row number to be vertical aligned to top

      if (!positions[rownum]) {
         positions[rownum] = cellposition.top; // first one takes precedence
         sizes[rownum] = trowobj.firstChild.offsetHeight;
      }
   }

   return;
};

/**
 * @function CalculateColPositions
 * @memberof SocialCalc
 * @description Calculates the screen positions of columns in a specific pane
 * @param {Object} editor - The table editor instance
 * @param {number} panenum - The pane number to calculate
 * @param {Array} positions - Array to store column positions
 * @param {Array} sizes - Array to store column sizes
 */
SocialCalc.CalculateColPositions = function(editor, panenum, positions, sizes) {
   let leftcol, colpane, colnum, offset, trowobj, cellposition;

   const context = editor.context;
   const sheetobj = context.sheetobj;

   let tbodyobj;

   if (!context.showRCHeaders) throw new Error("Needs showRCHeaders=true");

   tbodyobj = editor.fullgrid.lastChild;

   // Calculate start of this pane as column in this table:
   leftcol = 1;
   for (colpane = 0; colpane < panenum; colpane++) {
      leftcol += context.colpanes[colpane].last - context.colpanes[colpane].first + 2; // skip pane and spacing col
   }

   trowobj = tbodyobj.childNodes[1]; // get heading row, which has all columns
   offset = 0;
   for (colnum = context.colpanes[colpane].first; colnum <= context.colpanes[colpane].last; colnum++) {
      cellposition = SocialCalc.GetElementPosition(trowobj.childNodes[leftcol + offset]);
      if (!positions[colnum]) {
         positions[colnum] = cellposition.left; // first one takes precedence
         if (trowobj.childNodes[leftcol + offset]) {
            sizes[colnum] = trowobj.childNodes[leftcol + offset].offsetWidth;
         }
      }
      offset++;
   }

   return;
};

/**
 * @function ScrollRelative
 * @memberof SocialCalc
 * @description If vertical true, scrolls up(-)/down(+), else left(-)/right(+)
 * @param {Object} editor - The table editor instance
 * @param {boolean} vertical - Whether to scroll vertically (true) or horizontally (false)
 * @param {number} amount - The amount to scroll (positive or negative)
 */
SocialCalc.ScrollRelative = function(editor, vertical, amount) {
   if (vertical) {
      editor.ScrollRelativeBoth(amount, 0);
   } else {
      editor.ScrollRelativeBoth(0, amount);
   }
   return;
};

/**
 * @function ScrollRelativeBoth
 * @memberof SocialCalc
 * @description Does both vertical and horizontal scrolling with one render
 * @param {Object} editor - The table editor instance
 * @param {number} vamount - Vertical scroll amount
 * @param {number} hamount - Horizontal scroll amount
 */
SocialCalc.ScrollRelativeBoth = function(editor, vamount, hamount) {
   const context = editor.context;

   const vplen = context.rowpanes.length;
   const vlimit = vplen > 1 ? context.rowpanes[vplen - 2].last + 1 : 1; // don't scroll past here
   if (context.rowpanes[vplen - 1].first + vamount < vlimit) { // limit amount
      vamount = (-context.rowpanes[vplen - 1].first) + vlimit;
   }

   const hplen = context.colpanes.length;
   const hlimit = hplen > 1 ? context.colpanes[hplen - 2].last + 1 : 1; // don't scroll past here
   if (context.colpanes[hplen - 1].first + hamount < hlimit) { // limit amount
      hamount = (-context.colpanes[hplen - 1].first) + hlimit;
   }

   if ((vamount === 1 || vamount === -1) && hamount === 0) { // special case quick scrolls
      if (vamount === 1) {
         editor.ScrollTableUpOneRow();
      } else {
         editor.ScrollTableDownOneRow();
      }
      if (editor.ecell) editor.SetECellHeaders("selected");
      editor.SchedulePositionCalculations();
      return;
   }

   // Do a gross move and render
   if (vamount !== 0 || hamount !== 0) {
      context.rowpanes[vplen - 1].first += vamount;
      context.rowpanes[vplen - 1].last += vamount;
      context.colpanes[hplen - 1].first += hamount;
      context.colpanes[hplen - 1].last += hamount;
      editor.FitToEditTable();
      editor.ScheduleRender();
   }
};
/**
 * @function PageRelative
 * @memberof SocialCalc
 * @description If vertical true, pages up(direction is -)/down(+), else left(-)/right(+)
 * @param {Object} editor - The table editor instance
 * @param {boolean} vertical - Whether to page vertically (true) or horizontally (false)
 * @param {number} direction - Direction to page: negative for up/left, positive for down/right
 */
SocialCalc.PageRelative = function(editor, vertical, direction) {
   const context = editor.context;
   const panes = vertical ? "rowpanes" : "colpanes";
   const lastpane = context[panes][context[panes].length - 1];
   const lastvisible = vertical ? "lastvisiblerow" : "lastvisiblecol";
   const sizearray = vertical ? editor.rowheight : editor.colwidth;
   const defaultsize = vertical ? SocialCalc.Constants.defaultAssumedRowHeight : SocialCalc.Constants.defaultColWidth;
   let size, newfirst, totalsize, current;

   if (direction > 0) { // down/right
      newfirst = editor[lastvisible];
      if (newfirst === lastpane.first) newfirst += 1; // move at least one
   } else {
      if (vertical) { // calculate amount to scroll
         totalsize = editor.tableheight - (editor.firstscrollingrowtop - editor.gridposition.top);
      } else {
         totalsize = editor.tablewidth - (editor.firstscrollingcolleft - editor.gridposition.left);
      }
      totalsize -= sizearray[editor[lastvisible]] > 0 ? sizearray[editor[lastvisible]] : defaultsize;

      for (newfirst = lastpane.first - 1; newfirst > 0; newfirst--) {
         size = sizearray[newfirst] > 0 ? sizearray[newfirst] : defaultsize;
         if (totalsize < size) break;
         totalsize -= size;
      }

      current = lastpane.first;
      if (newfirst >= current) newfirst = current - 1; // move at least 1
      if (newfirst < 1) newfirst = 1;
   }

   lastpane.first = newfirst;
   lastpane.last = newfirst + 1;
   editor.LimitLastPanes();
   editor.FitToEditTable();
   editor.ScheduleRender();
};

/**
 * @function LimitLastPanes
 * @memberof SocialCalc
 * @description Makes sure that the "first" of the last panes isn't before the last of the previous pane
 * @param {Object} editor - The table editor instance
 */
SocialCalc.LimitLastPanes = function(editor) {
   const context = editor.context;
   let plen;

   plen = context.rowpanes.length;
   if (plen > 1 && context.rowpanes[plen - 1].first <= context.rowpanes[plen - 2].last)
      context.rowpanes[plen - 1].first = context.rowpanes[plen - 2].last + 1;

   plen = context.colpanes.length;
   if (plen > 1 && context.colpanes[plen - 1].first <= context.colpanes[plen - 2].last)
      context.colpanes[plen - 1].first = context.colpanes[plen - 2].last + 1;
};

/**
 * @function ScrollTableUpOneRow
 * @memberof SocialCalc
 * @description Scrolls the table up by one row using DOM manipulation for performance
 * @param {Object} editor - The table editor instance
 * @returns {HTMLElement} The table object
 */
SocialCalc.ScrollTableUpOneRow = function(editor) {
   let toprow, rowpane, rownum, colnum, colpane, cell, oldrownum, maxspan, newbottomrow, newrow, oldchild, bottomrownum;
   const rowneedsrefresh = {};

   const context = editor.context;
   const sheetobj = context.sheetobj;
   const tableobj = editor.fullgrid;

   let tbodyobj;

   tbodyobj = tableobj.lastChild;

   toprow = context.showRCHeaders ? 2 : 1;
   for (rowpane = 0; rowpane < context.rowpanes.length - 1; rowpane++) {
      toprow += context.rowpanes[rowpane].last - context.rowpanes[rowpane].first + 2; // skip pane and spacing row
   }

   tbodyobj.removeChild(tbodyobj.childNodes[toprow]);

   context.rowpanes[rowpane].first++;
   context.rowpanes[rowpane].last++;
   editor.FitToEditTable();
   context.CalculateColWidthData(); // Just in case, since normally done in RenderSheet

   newbottomrow = context.RenderRow(context.rowpanes[rowpane].last, rowpane);
   tbodyobj.appendChild(newbottomrow);

   // if scrolled off a row with starting rowspans, replace rows for the largest rowspan
   let maxrowspan = 1;
   oldrownum = context.rowpanes[rowpane].first - 1;

   for (colpane = 0; colpane < context.colpanes.length; colpane++) {
      for (colnum = context.colpanes[colpane].first; colnum <= context.colpanes[colpane].last; colnum++) {
         const coord = SocialCalc.crToCoord(colnum, oldrownum);
         if (context.cellskip[coord]) continue;
         cell = sheetobj.cells[coord];
         if (cell && cell.rowspan > maxrowspan) maxrowspan = cell.rowspan;
      }
   }

   if (maxrowspan > 1) {
      for (rownum = 1; rownum < maxrowspan; rownum++) {
         if (rownum + oldrownum >= context.rowpanes[rowpane].last) break;
         newrow = context.RenderRow(rownum + oldrownum, rowpane);
         oldchild = tbodyobj.childNodes[toprow + rownum - 1];
         tbodyobj.replaceChild(newrow, oldchild);
      }
   }

   // if added a row that includes rowspans from above, update the size of those to include new row
   bottomrownum = context.rowpanes[rowpane].last;

   for (colpane = 0; colpane < context.colpanes.length; colpane++) {
      for (colnum = context.colpanes[colpane].first; colnum <= context.colpanes[colpane].last; colnum++) {
         const coord = context.cellskip[SocialCalc.crToCoord(colnum, bottomrownum)];
         if (!coord) continue; // only look at spanned cells
         rownum = context.coordToCR[coord].row - 0;
         if (rownum === context.rowpanes[rowpane].last ||
             rownum < context.rowpanes[rowpane].first) continue; // this row (colspan) or starts above pane
         cell = sheetobj.cells[coord];
         if (cell && cell.rowspan > 1) rowneedsrefresh[rownum] = true; // remember row num to update
      }
   }

   for (rownum in rowneedsrefresh) {
      newrow = context.RenderRow(rownum, rowpane);
      oldchild = tbodyobj.childNodes[(toprow + (rownum - context.rowpanes[rowpane].first))];
      tbodyobj.replaceChild(newrow, oldchild);
   }

   return tableobj;
};

/**
 * @function ScrollTableDownOneRow
 * @memberof SocialCalc
 * @description Scrolls the table down by one row using DOM manipulation for performance
 * @param {Object} editor - The table editor instance
 * @returns {HTMLElement} The table object
 */
SocialCalc.ScrollTableDownOneRow = function(editor) {
   let toprow, rowpane, rownum, colnum, colpane, cell, newrownum, maxspan, newbottomrow, newrow, oldchild, bottomrownum;
   const rowneedsrefresh = {};

   const context = editor.context;
   const sheetobj = context.sheetobj;
   const tableobj = editor.fullgrid;

   let tbodyobj;

   tbodyobj = tableobj.lastChild;

   toprow = context.showRCHeaders ? 2 : 1;
   for (rowpane = 0; rowpane < context.rowpanes.length - 1; rowpane++) {
      toprow += context.rowpanes[rowpane].last - context.rowpanes[rowpane].first + 2; // skip pane and spacing row
   }

   tbodyobj.removeChild(tbodyobj.childNodes[toprow + (context.rowpanes[rowpane].last - context.rowpanes[rowpane].first)]);

   context.rowpanes[rowpane].first--;
   context.rowpanes[rowpane].last--;
   editor.FitToEditTable();
   context.CalculateColWidthData(); // Just in case, since normally done in RenderSheet

   newrow = context.RenderRow(context.rowpanes[rowpane].first, rowpane);
   tbodyobj.insertBefore(newrow, tbodyobj.childNodes[toprow]);

   // if inserted a row with starting rowspans, replace rows for the largest rowspan
   let maxrowspan = 1;
   newrownum = context.rowpanes[rowpane].first;

   for (colpane = 0; colpane < context.colpanes.length; colpane++) {
      for (colnum = context.colpanes[colpane].first; colnum <= context.colpanes[colpane].last; colnum++) {
         const coord = SocialCalc.crToCoord(colnum, newrownum);
         if (context.cellskip[coord]) continue;
         cell = sheetobj.cells[coord];
         if (cell && cell.rowspan > maxrowspan) maxrowspan = cell.rowspan;
      }
   }

   if (maxrowspan > 1) {
      for (rownum = 1; rownum < maxrowspan; rownum++) {
         if (rownum + newrownum > context.rowpanes[rowpane].last) break;
         newrow = context.RenderRow(rownum + newrownum, rowpane);
         oldchild = tbodyobj.childNodes[toprow + rownum];
         tbodyobj.replaceChild(newrow, oldchild);
      }
   }

   // if last row now includes rowspans or rowspans from above, update the size of those to remove deleted row
   bottomrownum = context.rowpanes[rowpane].last;

   for (colpane = 0; colpane < context.colpanes.length; colpane++) {
      for (colnum = context.colpanes[colpane].first; colnum <= context.colpanes[colpane].last; colnum++) {
         let coord = SocialCalc.crToCoord(colnum, bottomrownum);
         cell = sheetobj.cells[coord];
         if (cell && cell.rowspan > 1) {
            rowneedsrefresh[bottomrownum] = true; // need to update this row
            continue;
         }
         coord = context.cellskip[SocialCalc.crToCoord(colnum, bottomrownum)];
         if (!coord) continue; // only look at spanned cells
         rownum = context.coordToCR[coord].row - 0;
         if (rownum === bottomrownum ||
             rownum < context.rowpanes[rowpane].first) continue; // this row (colspan) or starts above pane
         cell = sheetobj.cells[coord];
         if (cell && cell.rowspan > 1) rowneedsrefresh[rownum] = true; // remember row num to update
      }
   }

   for (rownum in rowneedsrefresh) {
      newrow = context.RenderRow(rownum, rowpane);
      oldchild = tbodyobj.childNodes[(toprow + (rownum - context.rowpanes[rowpane].first))];
      tbodyobj.replaceChild(newrow, oldchild);
   }

   return tableobj;
};

/**
 * @class InputBox
 * @memberof SocialCalc
 * @description This class deals with the text box for editing cell contents.
 * It mainly controls a user input box for typed content and is used to interact with
 * the keyboard code, etc.
 * 
 * You can use this inside a formula bar control of some sort.
 * You create this after you have created a table editor object (but not necessarily 
 * done the CreateTableEditor method).
 * 
 * When the user starts typing text, or double-clicks on a cell, this object
 * comes into play.
 * 
 * The element given when this is first constructed should be an input HTMLElement or
 * something that acts like one. Check the code here to see what is done to it.
 * @param {HTMLElement} element - The input element associated with this InputBox
 * @param {Object} editor - The TableEditor this belongs to
 */
SocialCalc.InputBox = function(element, editor) {
   if (!element) return; // invoked without enough data to work

   /**
    * @member {HTMLElement} element - The input element associated with this InputBox
    */
   this.element = element;
   
   /**
    * @member {Object} editor - The TableEditor this belongs to
    */
   this.editor = editor;
   
   /**
    * @member {Object|null} inputEcho - Input echo object
    */
   this.inputEcho = null;

   editor.inputBox = this;

   element.onmousedown = SocialCalc.InputBoxOnMouseDown;

   editor.MoveECellCallback.formulabar = function(e) {
      if (e.state !== "start") return; // if not in normal keyboard mode don't replace formula bar
      editor.inputBox.DisplayCellContents(e.ecell.coord);
   };
};

/**
 * @description Methods for SocialCalc.InputBox prototype
 */

/**
 * @method DisplayCellContents
 * @memberof SocialCalc.InputBox
 * @description Displays the contents of a specified cell in the input box
 * @param {string} coord - Cell coordinate to display (optional)
 */
SocialCalc.InputBox.prototype.DisplayCellContents = function(coord) {
   SocialCalc.InputBoxDisplayCellContents(this, coord);
};

/**
 * @method ShowInputBox
 * @memberof SocialCalc.InputBox
 * @description Shows or hides the input box
 * @param {boolean} show - Whether to show the input box
 */
SocialCalc.InputBox.prototype.ShowInputBox = function(show) {
   this.editor.inputEcho.ShowInputEcho(show);
};

/**
 * @method GetText
 * @memberof SocialCalc.InputBox
 * @description Gets the current text value from the input element
 * @returns {string} The current input text
 */
SocialCalc.InputBox.prototype.GetText = function() {
   return this.element.value;
};

/**
 * @method SetText
 * @memberof SocialCalc.InputBox
 * @description Sets the text value of the input element
 * @param {string} newtext - The new text to set
 */
SocialCalc.InputBox.prototype.SetText = function(newtext) {
   if (!this.element) return;
   this.element.value = newtext;
   if (!SocialCalc.HasTouch) {
      this.editor.inputEcho.SetText(`${newtext}_`);
   } else {
      this.editor.inputEcho.SetText(newtext);
   }
};

/**
 * @method Focus
 * @memberof SocialCalc.InputBox
 * @description Sets focus to the input box
 */
SocialCalc.InputBox.prototype.Focus = function() {
   SocialCalc.InputBoxFocus(this);
};

/**
 * @method Blur
 * @memberof SocialCalc.InputBox
 * @description Removes focus from the input element
 * @returns {*} Result of element.blur()
 */
SocialCalc.InputBox.prototype.Blur = function() {
   return this.element.blur();
};

/**
 * @method Select
 * @memberof SocialCalc.InputBox
 * @description Selects text in the input element
 * @param {string} t - Selection type ("end" to move cursor to end)
 */
SocialCalc.InputBox.prototype.Select = function(t) {
   if (!this.element) return;
   switch (t) {
      case "end":
         if (document.selection && document.selection.createRange) {
            /* IE 4+ - Safer than setting .selectionEnd as it also works for Textareas. */
            const range = document.selection.createRange().duplicate();
            range.moveToElementText(this.element);
            range.collapse(false);
            range.select();
         } else if (this.element.selectionStart !== undefined) {
            this.element.selectionStart = this.element.value.length;
            this.element.selectionEnd = this.element.value.length;
         }
         break;
   }
};

/**
 * @namespace Functions
 * @description InputBox related functions
 */

/**
 * @function InputBoxDisplayCellContents
 * @memberof SocialCalc
 * @description Sets input box to the contents of the specified cell (or ecell if null)
 * @param {Object} inputbox - The InputBox instance
 * @param {string} coord - Cell coordinate (optional, uses ecell if null)
 */
SocialCalc.InputBoxDisplayCellContents = function(inputbox, coord) {
   const scc = SocialCalc.Constants;

   if (!inputbox) return;
   if (!coord) coord = inputbox.editor.ecell.coord;
   let text = SocialCalc.GetCellContents(inputbox.editor.context.sheetobj, coord);
   if (text.indexOf("\n") !== -1) {
      text = scc.s_inputboxdisplaymultilinetext;
      inputbox.element.disabled = true;
   } else {
      inputbox.element.disabled = false;
   }
   inputbox.SetText(text);
};

/**
 * @function InputBoxFocus
 * @memberof SocialCalc
 * @description Call this to have the input box get the focus and respond to keystrokes
 * but still pass them off to SocialCalc.ProcessKey
 * @param {Object} inputbox - The InputBox instance
 */
SocialCalc.InputBoxFocus = function(inputbox) {
   if (!inputbox) return;
   inputbox.element.focus();
   const editor = inputbox.editor;
   editor.state = "input";
   const wval = editor.workingvalues;
   wval.partialexpr = "";
   wval.ecoord = editor.ecell.coord;
   wval.erow = editor.ecell.row;
   wval.ecol = editor.ecell.col;
};

/**
 * @function InputBoxOnMouseDown
 * @memberof SocialCalc
 * @description This is called when the input box gets the focus. It then responds to keystrokes
 * and pass them off to SocialCalc.ProcessKey, but in a different editing state.
 * @param {Event} e - The mouse down event
 * @returns {boolean} True to let browser handle default behavior
 */
SocialCalc.InputBoxOnMouseDown = function(e) {
   const editor = SocialCalc.Keyboard.focusTable; // get TableEditor doing keyboard stuff
   if (!editor) return true; // we're not handling it -- let browser do default
   const wval = editor.workingvalues;

   switch (editor.state) {
      case "start":
         editor.state = "inputboxdirect";
         wval.partialexpr = "";
         wval.ecoord = editor.ecell.coord;
         wval.erow = editor.ecell.row;
         wval.ecol = editor.ecell.col;
         editor.inputEcho.ShowInputEcho(true);
         break;

      case "input":
         wval.partialexpr = ""; // make sure not pointing
         editor.MoveECell(wval.ecoord);
         editor.state = "inputboxdirect";
         SocialCalc.KeyboardFocus(); // may have come here from outside of grid
         break;

      case "inputboxdirect":
         break;
   }
};

/**
 * @class InputEcho
 * @memberof SocialCalc
 * @description This object creates and controls an element that echos what's in the InputBox during editing
 * It is draggable.
 * @param {Object} editor - The TableEditor this belongs to
 */
SocialCalc.InputEcho = function(editor) {
   const scc = SocialCalc.Constants;

   /**
    * @member {Object} editor - The TableEditor this belongs to
    */
   this.editor = editor;
   
   /**
    * @member {string} text - Current value of what is displayed
    */
   this.text = "";
   
   /**
    * @member {number|null} interval - Timer handle
    */
   this.interval = null;

   /**
    * @member {HTMLElement|null} container - Element containing main echo as well as prompt line
    */
   this.container = null;
   
   /**
    * @member {HTMLElement|null} main - Main echo area
    */
   this.main = null;
   
   /**
    * @member {HTMLElement|null} prompt - Prompt element
    */
   this.prompt = null;

   /**
    * @member {HTMLElement|null} functionbox - Function chooser dialog
    */
   this.functionbox = null;

   this.container = document.createElement("div");
   SocialCalc.setStyles(this.container, "display:none;position:absolute;zIndex:10;");

   this.topprompt = document.createElement("div");
   if (scc.defaultInputEchoPromptClass) this.topprompt.className = scc.defaultInputEchoPromptClass;
   if (scc.defaultInputEchoPromptStyle) SocialCalc.setStyles(this.topprompt, scc.defaultInputEchoPromptStyle);
   this.topprompt.innerHTML = "";

   this.container.appendChild(this.topprompt);

   this.main = document.createElement("div");
   if (SocialCalc.HasTouch) {
      this.maininput = document.createElement("input");
   }
   if (scc.defaultInputEchoClass) this.main.className = scc.defaultInputEchoClass;
   if (scc.defaultInputEchoStyle) SocialCalc.setStyles(this.main, scc.defaultInputEchoStyle);
   if (SocialCalc.HasTouch) {
      this.main.appendChild(this.maininput);
   } else {
      this.main.innerHTML = "&nbsp;";
   }

   this.container.appendChild(this.main);

   this.prompt = document.createElement("div");
   if (scc.defaultInputEchoPromptClass) this.prompt.className = scc.defaultInputEchoPromptClass;
   if (scc.defaultInputEchoPromptStyle) SocialCalc.setStyles(this.prompt, scc.defaultInputEchoPromptStyle);
   this.prompt.innerHTML = "";

   this.container.appendChild(this.prompt);

   SocialCalc.DragRegister(this.main, true, true, {
      MouseDown: SocialCalc.DragFunctionStart, 
      MouseMove: SocialCalc.DragFunctionPosition,
      MouseUp: SocialCalc.DragFunctionPosition,
      Disabled: null, 
      positionobj: this.container
   });

   editor.toplevel.appendChild(this.container);
};
/**
 * @description Methods for SocialCalc.InputEcho prototype
 */

/**
 * @method ShowInputEcho
 * @memberof SocialCalc.InputEcho
 * @description Shows or hides the input echo
 * @param {boolean} show - Whether to show the input echo
 * @returns {*} Result of SocialCalc.ShowInputEcho
 */
SocialCalc.InputEcho.prototype.ShowInputEcho = function(show) {
   return SocialCalc.ShowInputEcho(this, show);
};

/**
 * @method SetText
 * @memberof SocialCalc.InputEcho
 * @description Sets the text content of the input echo
 * @param {string} str - The text to set
 * @returns {*} Result of SocialCalc.SetInputEchoText
 */
SocialCalc.InputEcho.prototype.SetText = function(str) {
   return SocialCalc.SetInputEchoText(this, str);
};

/**
 * @namespace Functions
 * @description InputEcho related functions
 */

/**
 * @function ShowInputEcho
 * @memberof SocialCalc
 * @description Controls the display and positioning of the input echo
 * @param {Object} inputecho - The InputEcho instance
 * @param {boolean} show - Whether to show or hide the input echo
 */
SocialCalc.ShowInputEcho = function(inputecho, show) {
   let cell, position;
   const editor = inputecho.editor;

   if (!editor) return;

   if (show) {
      editor.cellhandles.ShowCellHandles(false);
      cell = SocialCalc.GetEditorCellElement(editor, editor.ecell.row, editor.ecell.col);
      if (cell) {
         position = SocialCalc.GetElementPosition(cell.element);
         inputecho.container.style.left = `${position.left - 1}px`;
         inputecho.container.style.top = `${position.top - 1}px`;
      }
      inputecho.container.style.display = "block";
      if (inputecho.interval) window.clearInterval(inputecho.interval); // just in case
      inputecho.interval = window.setInterval(SocialCalc.InputEchoHeartbeat, 50);
      if (SocialCalc.HasTouch) {
         inputecho.maininput.focus();
      }
   } else {
      if (inputecho.interval) window.clearInterval(inputecho.interval);
      inputecho.container.style.display = "none";
      inputecho.topprompt.innerHTML = "";
      if (SocialCalc.HasTouch) {      
         inputecho.maininput.blur();
         inputecho.maininput.value = "";
      }
   }
};

/**
 * @function SetInputEchoText
 * @memberof SocialCalc
 * @description Sets the text content and handles function prompts for the input echo
 * @param {Object} inputecho - The InputEcho instance
 * @param {string} str - The text string to set
 */
SocialCalc.SetInputEchoText = function(inputecho, str) {
   const scc = SocialCalc.Constants;
   let fname, fstr;
   let newstr = SocialCalc.special_chars(str);
   newstr = newstr.replace(/\n/g, "<br>");

   if (!SocialCalc.HasTouch) {
      if (inputecho.text !== newstr) {
         inputecho.main.innerHTML = newstr;
         inputecho.text = newstr;
      }
   } else {
      if (inputecho.maininput.value !== newstr) {
         inputecho.maininput.value = newstr;
      }
   }

   const parts = str.match(/.*[\+\-\*\/\&\^\<\>\=\,\(]([A-Za-z][A-ZA-z]\w*?)\([^\)]*$/);
   if (str.charAt(0) === "=" && parts) {
      fname = parts[1].toUpperCase();
      if (SocialCalc.Formula.FunctionList[fname]) {
         SocialCalc.Formula.FillFunctionInfo(); //  make sure filled
         fstr = SocialCalc.special_chars(`${fname}(${SocialCalc.Formula.FunctionArgString(fname)})`);
      } else {
         fstr = scc.ietUnknownFunction + fname;
      }
      if (inputecho.prompt.innerHTML !== fstr) {
         inputecho.prompt.innerHTML = fstr;
         inputecho.prompt.style.display = "block";
      }
   } else if (inputecho.prompt.style.display !== "none") {
      inputecho.prompt.innerHTML = "";
      inputecho.prompt.style.display = "none";
   }

   const editor = inputecho.editor;   

   if (editor.workingvalues.currentsheet !== editor.workingvalues.startsheet) {
      const promptstr = `Editing:${editor.workingvalues.startsheet}!${editor.workingvalues.ecoord}`;
      if (promptstr !== inputecho.topprompt.innerHTML) {
         inputecho.topprompt.innerHTML = `Editing:${editor.workingvalues.startsheet}!${editor.workingvalues.ecoord}`;
         inputecho.topprompt.style.display = "block";     
      }
   } else {
      if (inputecho.topprompt.style.display !== "none") {
         inputecho.topprompt.innerHTML = "";
         inputecho.topprompt.style.display = "none";
      }
   }
};

/**
 * @function InputEchoHeartbeat
 * @memberof SocialCalc
 * @description Periodic function to sync input echo with input box content
 * @returns {boolean} Returns true if no editor is handling keyboard input
 */
SocialCalc.InputEchoHeartbeat = function() {
   const editor = SocialCalc.Keyboard.focusTable; // get TableEditor doing keyboard stuff
   if (!editor) return true; // we're not handling it -- let browser do default

   if (!SocialCalc.HasTouch || editor.state === "inputboxdirect") {
      editor.inputEcho.SetText(`${editor.inputBox.GetText()}_`);
   } else {
      editor.inputBox.element.value = editor.inputEcho.maininput.value;
   }
};

/**
 * @function InputEchoMouseDown
 * @memberof SocialCalc
 * @description Handles mouse down events on the input echo
 * @param {Event} e - The mouse event
 * @returns {boolean} Returns true to let browser handle default behavior
 */
SocialCalc.InputEchoMouseDown = function(e) {
   const event = e || window.event;

   const editor = SocialCalc.Keyboard.focusTable; // get TableEditor doing keyboard stuff
   if (!editor) return true; // we're not handling it -- let browser do default

//      if (event.stopPropagation) event.stopPropagation(); // DOM Level 2
//      else event.cancelBubble = true; // IE 5+
//      if (event.preventDefault) event.preventDefault(); // DOM Level 2
//      else event.returnValue = false; // IE 5+

   editor.inputBox.element.focus();

//      return false;
};

/**
 * @class CellHandles
 * @memberof SocialCalc
 * @description This object creates and controls the elements around the cursor cell for dragging, etc.
 * @param {Object} editor - The TableEditor this belongs to
 */
SocialCalc.CellHandles = function(editor) {
   const scc = SocialCalc.Constants;
   let functions;

   if (editor.noEdit) return; // leave us with nothing

   /**
    * @member {Object} editor - The TableEditor this belongs to
    */
   this.editor = editor;

   /**
    * @member {boolean} noCursorSuffix - Whether to show cursor suffix
    */
   this.noCursorSuffix = false;

   /**
    * @member {boolean} movedmouse - Used to detect no-op mouse movements
    */
   this.movedmouse = false;

   this.draghandle = document.createElement("div");
   SocialCalc.setStyles(this.draghandle, "display:none;position:absolute;zIndex:8;border:1px solid white;width:4px;height:4px;fontSize:1px;backgroundColor:#0E93D8;cursor:default;");
   this.draghandle.innerHTML = '&nbsp;';
   editor.toplevel.appendChild(this.draghandle);
   SocialCalc.AssignID(editor, this.draghandle, "draghandle");

   let imagetype = "png";
   if (navigator.userAgent.match(/MSIE 6\.0/)) {
      imagetype = "gif";
   }

   this.dragpalette = document.createElement("div");
   SocialCalc.setStyles(this.dragpalette, `display:none;position:absolute;zIndex:8;width:90px;height:90px;fontSize:1px;textAlign:center;cursor:default;backgroundImage:url(${SocialCalc.Constants.defaultImagePrefix}drag-handles.${imagetype});`);
   this.dragpalette.innerHTML = '&nbsp;';
   editor.toplevel.appendChild(this.dragpalette);
   SocialCalc.AssignID(editor, this.dragpalette, "dragpalette");

   this.dragtooltip = document.createElement("div");
   SocialCalc.setStyles(this.dragtooltip, "display:none;position:absolute;zIndex:9;border:1px solid black;width:100px;height:auto;fontSize:10px;backgroundColor:#FFFFFF;");
   this.dragtooltip.innerHTML = '&nbsp;';
   editor.toplevel.appendChild(this.dragtooltip);
   SocialCalc.AssignID(editor, this.dragtooltip, "dragtooltip");

   this.fillinghandle = document.createElement("div");
   SocialCalc.setStyles(this.fillinghandle, "display:none;position:absolute;zIndex:9;border:1px solid black;width:auto;height:14px;fontSize:10px;backgroundColor:#FFFFFF;");
   this.fillinghandle.innerHTML = '&nbsp;';
   editor.toplevel.appendChild(this.fillinghandle);
   SocialCalc.AssignID(editor, this.fillinghandle, "fillinghandle");

   if (this.draghandle.addEventListener) { // DOM Level 2 -- Firefox, et al
      this.draghandle.addEventListener("mousemove", SocialCalc.CellHandlesMouseMoveOnHandle, false);
      this.dragpalette.addEventListener("mousedown", SocialCalc.CellHandlesMouseDown, false);
      this.dragpalette.addEventListener("mousemove", SocialCalc.CellHandlesMouseMoveOnHandle, false);
   } else if (this.draghandle.attachEvent) { // IE 5+
      this.draghandle.attachEvent("onmousemove", SocialCalc.CellHandlesMouseMoveOnHandle);
      this.dragpalette.attachEvent("onmousedown", SocialCalc.CellHandlesMouseDown);
      this.dragpalette.attachEvent("onmousemove", SocialCalc.CellHandlesMouseMoveOnHandle);
   } else { // don't handle this
      throw new Error("Browser not supported");
   }
};

/**
 * @description Methods for SocialCalc.CellHandles prototype
 */

/**
 * @method ShowCellHandles
 * @memberof SocialCalc.CellHandles
 * @description Shows or hides the cell handles
 * @param {boolean} show - Whether to show the handles
 * @param {boolean} moveshow - Whether to show move handles
 * @returns {*} Result of SocialCalc.ShowCellHandles
 */
SocialCalc.CellHandles.prototype.ShowCellHandles = function(show, moveshow) {
   return SocialCalc.ShowCellHandles(this, show, moveshow);
};

/**
 * @namespace Functions
 * @description CellHandles related functions
 */

/**
 * @function ShowCellHandles
 * @memberof SocialCalc
 * @description Controls the display and positioning of cell handles around the current cell
 * @param {Object} cellhandles - The CellHandles instance
 * @param {boolean} show - Whether to show the handles
 * @param {boolean} moveshow - Whether to show move-specific handles
 */
SocialCalc.ShowCellHandles = function(cellhandles, show, moveshow) {
   let cell, cell2, position, position2;
   const editor = cellhandles.editor;
   let doshow = false;
   let row, col, viewport;

   if (!editor) return;

   do { // a block that you can "break" out of easily
      if (!show) break;

      row = editor.ecell.row;
      col = editor.ecell.col;

      if (editor.state !== "start") break;
      if (row >= editor.lastvisiblerow) break;
      if (col >= editor.lastvisiblecol) break;
      if (row < editor.firstscrollingrow) break;
      if (col < editor.firstscrollingcol) break;

      if (editor.rowpositions[row + 1] + 20 > editor.horizontaltablecontrol.controlborder) {
         break;
      }
      if (editor.rowpositions[row + 1] - 10 < editor.headposition.top) {
         break;
      }
      if (editor.colpositions[col + 1] + 20 > editor.verticaltablecontrol.controlborder) {
         break;
      }
      if (editor.colpositions[col + 1] - 30 < editor.headposition.left) {
         break;
      }

      cellhandles.draghandle.style.left = `${editor.colpositions[col + 1] - 1}px`;
      cellhandles.draghandle.style.top = `${editor.rowpositions[row + 1] - 1}px`;
      cellhandles.draghandle.style.display = "block";

      if (moveshow) {
         cellhandles.draghandle.style.display = "none";
         cellhandles.dragpalette.style.left = `${editor.colpositions[col + 1] - 45}px`;
         cellhandles.dragpalette.style.top = `${editor.rowpositions[row + 1] - 45}px`;
         cellhandles.dragpalette.style.display = "block";
         viewport = SocialCalc.GetViewportInfo();
         cellhandles.dragtooltip.style.right = `${viewport.width - (editor.colpositions[col + 1] - 1)}px`;
         cellhandles.dragtooltip.style.bottom = `${viewport.height - (editor.rowpositions[row + 1] - 1)}px`;
         cellhandles.dragtooltip.style.display = "none";
      }

      doshow = true;
   } while (false); // only do once

   if (!doshow) {
      cellhandles.draghandle.style.display = "none";
   }
   if (!moveshow) {
      cellhandles.dragpalette.style.display = "none";
      cellhandles.dragtooltip.style.display = "none";
   }
};
/**
 * @function CellHandlesMouseMoveOnHandle
 * @memberof SocialCalc
 * @description Handles mouse move events on cell handles, showing tooltips and managing hover states
 * @param {Event} e - The mouse move event
 */
SocialCalc.CellHandlesMouseMoveOnHandle = function(e) {
   const scc = SocialCalc.Constants;

   const event = e || window.event;
   const target = event.target || event.srcElement;
   const viewport = SocialCalc.GetViewportInfo();
   const clientX = event.clientX + viewport.horizontalScroll;
   const clientY = event.clientY + viewport.verticalScroll;

   const editor = SocialCalc.Keyboard.focusTable; // get TableEditor doing keyboard stuff
   if (!editor) return true; // we're not handling it -- let browser do default
   const cellhandles = editor.cellhandles;
   if (!cellhandles.editor) return true; // no handles

   if (!editor.cellhandles.mouseDown) {
      editor.cellhandles.ShowCellHandles(true, true); // show move handles, too

      if (target === cellhandles.dragpalette) {
         const whichhandle = SocialCalc.SegmentDivHit([scc.CH_radius1, scc.CH_radius2], editor.cellhandles.dragpalette, clientX, clientY);
         if (whichhandle === 0) { // off of active part of palette
            SocialCalc.CellHandlesHoverTimeout();
            return;
         }
         if (cellhandles.tooltipstimer) {
            window.clearTimeout(cellhandles.tooltipstimer);
            cellhandles.tooltipstimer = null;
         }
         cellhandles.tooltipswhichhandle = whichhandle;
         cellhandles.tooltipstimer = window.setTimeout(SocialCalc.CellHandlesTooltipsTimeout, 700);
      }

      if (cellhandles.timer) {
         window.clearTimeout(cellhandles.timer);
         cellhandles.timer = null;
      }
      cellhandles.timer = window.setTimeout(SocialCalc.CellHandlesHoverTimeout, 3000);
   }

   return;
};

/**
 * @function SegmentDivHit
 * @memberof SocialCalc
 * @description Takes segtable = [upperleft quadrant, upperright, bottomright, bottomleft]
 * where each quadrant is either:
 *     0 = ignore hits here
 *     number = return this value
 *     array = a new segtable for this subquadrant
 * 
 * Alternatively, segtable can be:
 * [radius 1, radius 2] and it returns 0 if no hit,
 * -1, -2, -3, -4 for inner quadrants, and +1...+4 for outer quadrants
 * @param {Array} segtable - Segment table defining hit areas
 * @param {HTMLElement} divWithMouseHit - The element that received the mouse hit
 * @param {number} x - X coordinate of the mouse
 * @param {number} y - Y coordinate of the mouse
 * @returns {number} Which segment was hit
 */
SocialCalc.SegmentDivHit = function(segtable, divWithMouseHit, x, y) {
   const width = divWithMouseHit.offsetWidth;
   const height = divWithMouseHit.offsetHeight;
   let left = divWithMouseHit.offsetLeft;
   let top = divWithMouseHit.offsetTop;
   let v = 0;
   let table = segtable;
   const len = Math.sqrt(Math.pow(x - left - (width / 2.0 - 0.5), 2) + Math.pow(y - top - (height / 2.0 - 0.5), 2));

   if (table.length === 2) { // type 2 segtable
      if (x >= left && x < left + width / 2 && y >= top && y < top + height / 2) { // upper left
         if (len <= segtable[0]) v = -1;
         else if (len <= segtable[1]) v = 1;
      }
      if (x >= left + width / 2 && x < left + width && y >= top && y < top + height / 2) { // upper right
         if (len <= segtable[0]) v = -2;
         else if (len <= segtable[1]) v = 2;
      }
      if (x >= left + width / 2 && x < left + width && y >= top + height / 2 && y < top + height) { // bottom right
         if (len <= segtable[0]) v = -3;
         else if (len <= segtable[1]) v = 3;
      }
      if (x >= left && x < left + width / 2 && y >= top + height / 2 && y < top + height) { // bottom left
         if (len <= segtable[0]) v = -4;
         else if (len <= segtable[1]) v = 4;
      }
      return v;
   }

   let quadrant = "";
   while (true) {
      if (x >= left && x < left + width / 2 && y >= top && y < top + height / 2) { // upper left
         quadrant += "1";
         v = table[0];
         if (typeof v === "number") {
            break;
         }
         table = v;
         width = width / 2;
         height = height / 2;
         continue;
      }
      if (x >= left + width / 2 && x < left + width && y >= top && y < top + height / 2) { // upper right
         quadrant += "2";
         v = table[1];
         if (typeof v === "number") {
            break;
         }
         table = v;
         width = width / 2;
         left = left + width;
         height = height / 2;
         continue;
      }
      if (x >= left + width / 2 && x < left + width && y >= top + height / 2 && y < top + height) { // bottom right
         quadrant += "3";
         v = table[2];
         if (typeof v === "number") {
            break;
         }
         table = v;
         width = width / 2;
         left = left + width;
         height = height / 2;
         top = top + height;
         continue;
      }
      if (x >= left && x < left + width / 2 && y >= top + height / 2 && y < top + height) { // bottom left
         quadrant += "4";
         v = table[3];
         if (typeof v === "number") {
            break;
         }
         table = v;
         width = width / 2;
         height = height / 2;
         top = top + height;
         continue;
      }
      return 0; // didn't match
   }

//addmsg((x-divWithMouseHit.offsetLeft)+","+(y-divWithMouseHit.offsetTop)+"="+quadrant+" "+v);
   return v;
};

/**
 * @function CellHandlesHoverTimeout
 * @memberof SocialCalc
 * @description Handles timeout for cell handle hover state, hiding move handles
 * @returns {boolean} Returns true if no editor is handling keyboard input
 */
SocialCalc.CellHandlesHoverTimeout = function() {
   const editor = SocialCalc.Keyboard.focusTable; // get TableEditor doing keyboard stuff
   if (!editor) return true; // we're not handling it -- let browser do default
   const cellhandles = editor.cellhandles;
   if (cellhandles.timer) {
      window.clearTimeout(cellhandles.timer);
      cellhandles.timer = null;
   }
   if (cellhandles.tooltipstimer) {
      window.clearTimeout(cellhandles.tooltipstimer);
      cellhandles.tooltipstimer = null;
   }
   editor.cellhandles.ShowCellHandles(true, false); // hide move handles
};

/**
 * @function CellHandlesTooltipsTimeout
 * @memberof SocialCalc
 * @description Handles timeout for displaying tooltips on cell handles based on which handle is hovered
 * @returns {boolean} Returns true if no editor is handling keyboard input
 */
SocialCalc.CellHandlesTooltipsTimeout = function() {
   const editor = SocialCalc.Keyboard.focusTable; // get TableEditor doing keyboard stuff
   if (!editor) return true; // we're not handling it -- let browser do default
   const cellhandles = editor.cellhandles;
   if (cellhandles.tooltipstimer) {
      window.clearTimeout(cellhandles.tooltipstimer);
      cellhandles.tooltipstimer = null;
   }

   const whichhandle = cellhandles.tooltipswhichhandle;
   if (whichhandle === 0) { // off of active part of palette
      SocialCalc.CellHandlesHoverTimeout();
      return;
   }
   if (whichhandle === -3) {
      cellhandles.dragtooltip.innerHTML = scc.s_CHfillAllTooltip;
   } else if (whichhandle === 3) {
      cellhandles.dragtooltip.innerHTML = scc.s_CHfillContentsTooltip;
   } else if (whichhandle === -2) {
      cellhandles.dragtooltip.innerHTML = scc.s_CHmovePasteAllTooltip;
   } else if (whichhandle === -4) {
      cellhandles.dragtooltip.innerHTML = scc.s_CHmoveInsertAllTooltip;
   } else if (whichhandle === 2) {
      cellhandles.dragtooltip.innerHTML = scc.s_CHmovePasteContentsTooltip;
   } else if (whichhandle === 4) {
      cellhandles.dragtooltip.innerHTML = scc.s_CHmoveInsertContentsTooltip;
   } else {
      cellhandles.dragtooltip.innerHTML = "&nbsp;";
      cellhandles.dragtooltip.style.display = "none";
      return;
   }

   cellhandles.dragtooltip.style.display = "block";
};
/**
 * @function CellHandlesMouseDown
 * @memberof SocialCalc
 * @description Handles mouse down events on cell handles, initiating drag operations for fill or move
 * @param {Event} e - The mouse down event
 * @returns {boolean} Returns true to let browser handle default behavior if not handled
 */
SocialCalc.CellHandlesMouseDown = function(e) {
   const scc = SocialCalc.Constants;
   let editor, result, coord, textarea, wval, range;

   const event = e || window.event;

   const viewport = SocialCalc.GetViewportInfo();
   const clientX = event.clientX + viewport.horizontalScroll;
   const clientY = event.clientY + viewport.verticalScroll;

   const mouseinfo = SocialCalc.EditorMouseInfo;

   editor = SocialCalc.Keyboard.focusTable; // get TableEditor doing keyboard stuff
   if (!editor) return true; // we're not handling it -- let browser do default

   if (editor.busy) return; // don't do anything when busy (is this correct?)

   const cellhandles = editor.cellhandles;

   cellhandles.movedmouse = false; // detect no-op

   if (cellhandles.timer) { // cancel timer
      window.clearTimeout(cellhandles.timer);
      cellhandles.timer = null;
   }
   if (cellhandles.tooltipstimer) {
      window.clearTimeout(cellhandles.tooltipstimer);
      cellhandles.tooltipstimer = null;
   }
   cellhandles.dragtooltip.innerHTML = "&nbsp;";
   cellhandles.dragtooltip.style.display = "none";

   range = editor.range;
 
   const whichhandle = SocialCalc.SegmentDivHit([scc.CH_radius1, scc.CH_radius2], editor.cellhandles.dragpalette, clientX, clientY);
   if (whichhandle === 1 || whichhandle === -1 || whichhandle === 0) {
      cellhandles.ShowCellHandles(true, false); // hide move handles
      return;
   }

   mouseinfo.ignore = true; // stop other code from looking at the mouse

   if (whichhandle === -3) {
      cellhandles.dragtype = "Fill";
//      mouseinfo.element = editor.cellhandles.fillhandle;
      cellhandles.noCursorSuffix = false;
   } else if (whichhandle === 3) {
      cellhandles.dragtype = "FillC";
//      mouseinfo.element = editor.cellhandles.fillhandle;
      cellhandles.noCursorSuffix = false;
   } else if (whichhandle === -2) {
      cellhandles.dragtype = "Move";
//      mouseinfo.element = editor.cellhandles.movehandle1;
      cellhandles.noCursorSuffix = true;
   } else if (whichhandle === -4) {
      cellhandles.dragtype = "MoveI";
//      mouseinfo.element = editor.cellhandles.movehandle2;
      cellhandles.noCursorSuffix = false;
   } else if (whichhandle === 2) {
      cellhandles.dragtype = "MoveC";
//      mouseinfo.element = editor.cellhandles.movehandle1;
      cellhandles.noCursorSuffix = true;
   } else if (whichhandle === 4) {
      cellhandles.dragtype = "MoveIC";
//      mouseinfo.element = editor.cellhandles.movehandle2;
      cellhandles.noCursorSuffix = false;
   }

   cellhandles.filltype = null;

   switch (cellhandles.dragtype) {
      case "Fill":
      case "FillC":
         if (!range.hasrange) {
            editor.RangeAnchor();
         }
         break;

      case "Move":
      case "MoveI":
      case "MoveC":
      case "MoveIC":
         if (!range.hasrange) {
            editor.RangeAnchor();
         }
         editor.range2.top = editor.range.top;
         editor.range2.right = editor.range.right;
         editor.range2.bottom = editor.range.bottom;
         editor.range2.left = editor.range.left;
         editor.range2.hasrange = true;
         editor.RangeRemove();
         break;

      default:
         return; // not for us
   }

   cellhandles.fillinghandle.style.left = `${clientX}px`;
   cellhandles.fillinghandle.style.top = `${clientY - 17}px`;
   cellhandles.fillinghandle.innerHTML = scc.s_CHindicatorOperationLookup[cellhandles.dragtype] +
                                         (scc.s_CHindicatorDirectionLookup[editor.cellhandles.filltype] || "");
   cellhandles.fillinghandle.style.display = "block";

   cellhandles.ShowCellHandles(true, false); // hide move handles
   cellhandles.mouseDown = true;

   mouseinfo.editor = editor; // remember for later

   coord = editor.ecell.coord; // start with cell with handles

   cellhandles.startingcoord = coord;
   cellhandles.startingX = clientX;
   cellhandles.startingY = clientY;

   mouseinfo.mouselastcoord = coord;

   SocialCalc.KeyboardSetFocus(editor);

   if (document.addEventListener) { // DOM Level 2 -- Firefox, et al
      document.addEventListener("mousemove", SocialCalc.CellHandlesMouseMove, true); // capture everywhere
      document.addEventListener("mouseup", SocialCalc.CellHandlesMouseUp, true); // capture everywhere
   } else if (cellhandles.draghandle.attachEvent) { // IE 5+
      cellhandles.draghandle.setCapture();
      cellhandles.draghandle.attachEvent("onmousemove", SocialCalc.CellHandlesMouseMove);
      cellhandles.draghandle.attachEvent("onmouseup", SocialCalc.CellHandlesMouseUp);
      cellhandles.draghandle.attachEvent("onlosecapture", SocialCalc.CellHandlesMouseUp);
   }
   if (event.stopPropagation) event.stopPropagation(); // DOM Level 2
   else event.cancelBubble = true; // IE 5+
   if (event.preventDefault) event.preventDefault(); // DOM Level 2
   else event.returnValue = false; // IE 5+

   return;
};

/**
 * @function CellHandlesMouseMove
 * @memberof SocialCalc
 * @description Handles mouse move events during cell handle dragging operations
 * @param {Event} e - The mouse move event
 */
SocialCalc.CellHandlesMouseMove = function(e) {
   const scc = SocialCalc.Constants;
   let editor, element, result, coord, now, textarea, sheetobj, cellobj, wval;
   let crstart, crend, cr, c, r;

   const event = e || window.event;

   const viewport = SocialCalc.GetViewportInfo();
   const clientX = event.clientX + viewport.horizontalScroll;
   const clientY = event.clientY + viewport.verticalScroll;

   const mouseinfo = SocialCalc.EditorMouseInfo;
   editor = mouseinfo.editor;
   if (!editor) return; // not us, ignore
   const cellhandles = editor.cellhandles;

   element = mouseinfo.element;

   result = SocialCalc.GridMousePosition(editor, clientX, clientY); // get cell with move

   if (!result) return;

   if (result && !result.coord) {
      SocialCalc.SetDragAutoRepeat(editor, result, SocialCalc.CellHandlesDragAutoRepeat);
      return;
   }

   SocialCalc.SetDragAutoRepeat(editor, null); // stop repeating if it was

   if (!result.coord) return;

   crstart = SocialCalc.coordToCr(editor.cellhandles.startingcoord);
   crend = SocialCalc.coordToCr(result.coord);

   cellhandles.movedmouse = true; // did move, so not no-op

   switch (cellhandles.dragtype) {
      case "Fill":
      case "FillC":
         if (result.coord === cellhandles.startingcoord) { // reset when come back
            cellhandles.filltype = null;
            cellhandles.startingX = clientX;
            cellhandles.startingY = clientY;
         } else {
            if (cellhandles.filltype) { // moving and have already determined filltype
               if (cellhandles.filltype === "Down") { // coerce to that
                  crend.col = crstart.col;
                  if (crend.row < crstart.row) crend.row = crstart.row;
               } else {
                  crend.row = crstart.row;
                  if (crend.col < crstart.col) crend.col = crstart.col;
               }
            } else {
               if (Math.abs(clientY - cellhandles.startingY) > 10) {
                  cellhandles.filltype = "Down";
               } else if (Math.abs(clientX - cellhandles.startingX) > 10) {
                  cellhandles.filltype = "Right";
               }
               crend.col = crstart.col; // until decide, leave it at start
               crend.row = crstart.row;
            }
         }
         result.coord = SocialCalc.crToCoord(crend.col, crend.row);
         if (result.coord !== mouseinfo.mouselastcoord) {
            editor.MoveECell(result.coord);
            editor.RangeExtend();
         }
         break;

      case "Move":
      case "MoveC":
         if (result.coord !== mouseinfo.mouselastcoord) {
            editor.MoveECell(result.coord);
            c = editor.range2.right - editor.range2.left + result.col;
            r = editor.range2.bottom - editor.range2.top + result.row;
            editor.RangeAnchor(SocialCalc.crToCoord(c, r));
            editor.RangeExtend();
         }
         break;

      case "MoveI":
      case "MoveIC":
         if (result.coord === cellhandles.startingcoord) { // reset when come back
            cellhandles.filltype = null;
            cellhandles.startingX = clientX;
            cellhandles.startingY = clientY;
         } else {
            if (cellhandles.filltype) { // moving and have already determined filltype
               if (cellhandles.filltype === "Vertical") { // coerce to that
                  crend.col = editor.range2.left;
                  if (crend.row >= editor.range2.top && crend.row <= editor.range2.bottom + 1) crend.row = editor.range2.bottom + 2;
               } else {
                  crend.row = editor.range2.top;
                  if (crend.col >= editor.range2.left && crend.col <= editor.range2.right + 1) crend.col = editor.range2.right + 2;
               }
            } else {
               if (Math.abs(clientY - cellhandles.startingY) > 10) {
                  cellhandles.filltype = "Vertical";
               } else if (Math.abs(clientX - cellhandles.startingX) > 10) {
                  cellhandles.filltype = "Horizontal";
               }
               crend.col = crstart.col; // until decide, leave it at start
               crend.row = crstart.row;
            }
         }
         result.coord = SocialCalc.crToCoord(crend.col, crend.row);
         if (result.coord !== mouseinfo.mouselastcoord) {
            editor.MoveECell(result.coord);
            if (!cellhandles.filltype) { // no fill type
               editor.RangeRemove();
            } else {
               c = editor.range2.right - editor.range2.left + crend.col;
               r = editor.range2.bottom - editor.range2.top + crend.row;
               editor.RangeAnchor(SocialCalc.crToCoord(c, r));
               editor.RangeExtend();
            }
         }
         break;
   }

   cellhandles.fillinghandle.style.left = `${clientX}px`;
   cellhandles.fillinghandle.style.top = `${clientY - 17}px`;
   cellhandles.fillinghandle.innerHTML = scc.s_CHindicatorOperationLookup[cellhandles.dragtype] +
                                         (scc.s_CHindicatorDirectionLookup[editor.cellhandles.filltype] || "");
   cellhandles.fillinghandle.style.display = "block";

   mouseinfo.mouselastcoord = result.coord;

   if (event.stopPropagation) event.stopPropagation(); // DOM Level 2
   else event.cancelBubble = true; // IE 5+
   if (event.preventDefault) event.preventDefault(); // DOM Level 2
   else event.returnValue = false; // IE 5+

   return;
};
/**
 * @function CellHandlesDragAutoRepeat
 * @memberof SocialCalc
 * @description Handles auto-repeat functionality when dragging cell handles beyond the visible area
 * @param {string} coord - The coordinate where the drag is happening
 * @param {string} direction - The direction of the drag ("left", "right", "up", "down")
 */
SocialCalc.CellHandlesDragAutoRepeat = function(coord, direction) {
   const mouseinfo = SocialCalc.EditorMouseInfo;
   const editor = mouseinfo.editor;
   if (!editor) return; // not us, ignore
   const cellhandles = editor.cellhandles;

   const crstart = SocialCalc.coordToCr(editor.cellhandles.startingcoord);
   const crend = SocialCalc.coordToCr(coord);

   let newcoord, c, r;

   let vscroll = 0;
   let hscroll = 0;

   if (direction === "left") hscroll = -1;
   else if (direction === "right") hscroll = 1;
   else if (direction === "up") vscroll = -1;
   else if (direction === "down") vscroll = 1;
   editor.ScrollRelativeBoth(vscroll, hscroll);

   switch (cellhandles.dragtype) {
      case "Fill":
      case "FillC":
         if (cellhandles.filltype) { // moving and have already determined filltype
            if (cellhandles.filltype === "Down") { // coerce to that
               crend.col = crstart.col;
               if (crend.row < crstart.row) crend.row = crstart.row;
            } else {
               crend.row = crstart.row;
               if (crend.col < crstart.col) crend.col = crstart.col;
            }
         } else {
            crend.col = crstart.col; // until decide, leave it at start
            crend.row = crstart.row;
         }

         newcoord = SocialCalc.crToCoord(crend.col, crend.row);
         if (newcoord !== mouseinfo.mouselastcoord) {
            editor.MoveECell(coord);
            editor.RangeExtend();
         }
         break;

      case "Move":
      case "MoveC":
         if (coord !== mouseinfo.mouselastcoord) {
            editor.MoveECell(coord);
            c = editor.range2.right - editor.range2.left + editor.ecell.col;
            r = editor.range2.bottom - editor.range2.top + editor.ecell.row;
            editor.RangeAnchor(SocialCalc.crToCoord(c, r));
            editor.RangeExtend();
         }
         break;

      case "MoveI":
      case "MoveIC":
         if (cellhandles.filltype) { // moving and have already determined filltype
            if (cellhandles.filltype === "Vertical") { // coerce to that
               crend.col = editor.range2.left;
               if (crend.row >= editor.range2.top && crend.row <= editor.range2.bottom + 1) crend.row = editor.range2.bottom + 2;
            } else {
               crend.row = editor.range2.top;
               if (crend.col >= editor.range2.left && crend.col <= editor.range2.right + 1) crend.col = editor.range2.right + 2;
            }
         } else {
            crend.col = crstart.col; // until decide, leave it at start
            crend.row = crstart.row;
         }

         newcoord = SocialCalc.crToCoord(crend.col, crend.row);
         if (newcoord !== mouseinfo.mouselastcoord) {
            editor.MoveECell(newcoord);
            c = editor.range2.right - editor.range2.left + crend.col;
            r = editor.range2.bottom - editor.range2.top + crend.row;
            editor.RangeAnchor(SocialCalc.crToCoord(c, r));
            editor.RangeExtend();
         }
         break;
   }

   mouseinfo.mouselastcoord = newcoord;
};

/**
 * @function CellHandlesMouseUp
 * @memberof SocialCalc
 * @description Handles mouse up events for cell handle dragging, executing the appropriate commands
 * @param {Event} e - The mouse up event
 * @returns {boolean} Returns false to prevent default behavior
 */
SocialCalc.CellHandlesMouseUp = function(e) {
   let editor, element, result, coord, now, textarea, sheetobj, cellobj, wval, cstr, cmdtype, cmdtype2;
   let crstart, crend;
   let sizec, sizer, deltac, deltar;

   const event = e || window.event;

   const viewport = SocialCalc.GetViewportInfo();
   const clientX = event.clientX + viewport.horizontalScroll;
   const clientY = event.clientY + viewport.verticalScroll;

   const mouseinfo = SocialCalc.EditorMouseInfo;
   editor = mouseinfo.editor;
   if (!editor) return; // not us, ignore
   const cellhandles = editor.cellhandles;

   element = mouseinfo.element;

   mouseinfo.ignore = false;

   result = SocialCalc.GridMousePosition(editor, clientX, clientY); // get cell with up

   SocialCalc.SetDragAutoRepeat(editor, null); // stop repeating if it was

   cellhandles.mouseDown = false;
   cellhandles.noCursorSuffix = false;

   cellhandles.fillinghandle.style.display = "none";

   if (!result) result = {};
   if (!result.coord) result.coord = editor.ecell.coord;

   switch (cellhandles.dragtype) {
      case "Fill":
      case "Move":
      case "MoveI":
         cmdtype2 = " all";
         break;
      case "FillC":
      case "MoveC":
      case "MoveIC":
         cmdtype2 = " formulas";
         break;
   }

   if (!cellhandles.movedmouse) { // didn't move: just leave one cell selected
      cellhandles.dragtype = "Nothing";
   }

   switch (cellhandles.dragtype) {
      case "Nothing":
         editor.Range2Remove();
         editor.RangeRemove();
         break;

      case "Fill":
      case "FillC":
         crstart = SocialCalc.coordToCr(cellhandles.startingcoord);
         crend = SocialCalc.coordToCr(result.coord);
         if (cellhandles.filltype) {
            if (cellhandles.filltype === "Down") {
               crend.col = crstart.col;
            } else {
               crend.row = crstart.row;
            }
         }
         result.coord = SocialCalc.crToCoord(crend.col, crend.row);

         editor.MoveECell(result.coord);
         editor.RangeExtend();

         if (editor.cellhandles.filltype === "Right") {
            cmdtype = "right";
         } else {
            cmdtype = "down";
         }
         cstr = `fill${cmdtype} ${SocialCalc.crToCoord(editor.range.left, editor.range.top)}:${SocialCalc.crToCoord(editor.range.right, editor.range.bottom)}${cmdtype2}`;
         editor.EditorScheduleSheetCommands(cstr, true, false);
         break;

      case "Move":
      case "MoveC":
         editor.context.cursorsuffix = "";
         cstr = `movepaste ${SocialCalc.crToCoord(editor.range2.left, editor.range2.top)}:${SocialCalc.crToCoord(editor.range2.right, editor.range2.bottom)} ${editor.ecell.coord}${cmdtype2}`;
         editor.EditorScheduleSheetCommands(cstr, true, false);
         editor.Range2Remove();
         break;

      case "MoveI":
      case "MoveIC":
         editor.context.cursorsuffix = "";
         sizec = editor.range2.right - editor.range2.left;
         sizer = editor.range2.bottom - editor.range2.top;
         deltac = editor.ecell.col - editor.range2.left;
         deltar = editor.ecell.row - editor.range2.top;
         cstr = `moveinsert ${SocialCalc.crToCoord(editor.range2.left, editor.range2.top)}:${SocialCalc.crToCoord(editor.range2.right, editor.range2.bottom)} ${editor.ecell.coord}${cmdtype2}`;
         editor.EditorScheduleSheetCommands(cstr, true, false);
         editor.Range2Remove();
         editor.RangeRemove();
         if (editor.cellhandles.filltype === " Horizontal" && deltac > 0) {
            editor.MoveECell(SocialCalc.crToCoord(editor.ecell.col - sizec - 1, editor.ecell.row));
         } else if (editor.cellhandles.filltype === " Vertical" && deltar > 0) {
            editor.MoveECell(SocialCalc.crToCoord(editor.ecell.col, editor.ecell.row - sizer - 1));
         }
         editor.RangeAnchor(SocialCalc.crToCoord(editor.ecell.col + sizec, editor.ecell.row + sizer));
         editor.RangeExtend();
         break;
   }

   if (event.stopPropagation) event.stopPropagation(); // DOM Level 2
   else event.cancelBubble = true; // IE 5+
   if (event.preventDefault) event.preventDefault(); // DOM Level 2
   else event.returnValue = false; // IE 5+

   if (document.removeEventListener) { // DOM Level 2
      document.removeEventListener("mousemove", SocialCalc.CellHandlesMouseMove, true);
      document.removeEventListener("mouseup", SocialCalc.CellHandlesMouseUp, true);
   } else if (cellhandles.draghandle.detachEvent) { // IE
      cellhandles.draghandle.detachEvent("onlosecapture", SocialCalc.CellHandlesMouseUp);
      cellhandles.draghandle.detachEvent("onmouseup", SocialCalc.CellHandlesMouseUp);
      cellhandles.draghandle.detachEvent("onmousemove", SocialCalc.CellHandlesMouseMove);
      cellhandles.draghandle.releaseCapture();
   }

   mouseinfo.editor = null;

   return false;
};
/**
 * @class TableControl
 * @memberof SocialCalc
 * @description This class deals with the horizontal and vertical scrollbars and pane sliders.
 * 
 * Layout structure:
 * +--------------+
 * | Endcap       |
 * +- - - - - - - +
 * |              |
 * +--------------+
 * | Pane Slider  |
 * +--------------+
 * |              |
 * | Less Button  |
 * |              |
 * +--------------+
 * | Scroll Area  |
 * |              |
 * |              |
 * +--------------+
 * | Thumb        |
 * +--------------+
 * |              |
 * +--------------+
 * |              |
 * | More Button  |
 * |              |
 * +--------------+
 * 
 * @param {Object} editor - The TableEditor this belongs to
 * @param {boolean} vertical - True if vertical control, false if horizontal
 * @param {number} size - Length in pixels
 */
SocialCalc.TableControl = function(editor, vertical, size) {
   const scc = SocialCalc.Constants;

   /**
    * @member {Object} editor - The TableEditor this belongs to
    */
   this.editor = editor;

   /**
    * @member {boolean} vertical - True if vertical control, false if horizontal
    */
   this.vertical = vertical;
   
   /**
    * @member {number} size - Length in pixels
    */
   this.size = size;

   /**
    * @member {HTMLElement|null} main - Main element containing all the others
    */
   this.main = null;
   
   /**
    * @member {HTMLElement|null} endcap - The area at the top/left between the end and the pane slider
    */
   this.endcap = null;
   
   /**
    * @member {HTMLElement|null} paneslider - The slider to adjust the pane split
    */
   this.paneslider = null;
   
   /**
    * @member {HTMLElement|null} lessbutton - The top/left scroll button
    */
   this.lessbutton = null;
   
   /**
    * @member {HTMLElement|null} morebutton - The bottom/right scroll button
    */
   this.morebutton = null;
   
   /**
    * @member {HTMLElement|null} scrollarea - The area between the scroll buttons
    */
   this.scrollarea = null;
   
   /**
    * @member {HTMLElement|null} thumb - The sliding thing in the scrollarea
    */
   this.thumb = null;

   // computed position values:

   /**
    * @member {number|null} controlborder - Left or top screen position for vertical or horizontal control
    */
   this.controlborder = null;
   
   /**
    * @member {number|null} endcapstart - Top or left screen position for vertical or horizontal control
    */
   this.endcapstart = null;
   
   /**
    * @member {number|null} panesliderstart - Pane slider start position
    */
   this.panesliderstart = null;
   
   /**
    * @member {number|null} lessbuttonstart - Less button start position
    */
   this.lessbuttonstart = null;
   
   /**
    * @member {number|null} morebuttonstart - More button start position
    */
   this.morebuttonstart = null;
   
   /**
    * @member {number|null} scrollareastart - Scroll area start position
    */
   this.scrollareastart = null;
   
   /**
    * @member {number|null} scrollareaend - Scroll area end position
    */
   this.scrollareaend = null;
   
   /**
    * @member {number|null} scrollareasize - Scroll area size
    */
   this.scrollareasize = null;
   
   /**
    * @member {number|null} thumbpos - Thumb position
    */
   this.thumbpos = null;

   // constants:

   /**
    * @member {number} controlthickness - Other dimension of complete control in pixels
    */
   this.controlthickness = scc.defaultTableControlThickness;
   
   /**
    * @member {number} sliderthickness - Slider thickness
    */
   this.sliderthickness = scc.defaultTCSliderThickness;
   
   /**
    * @member {number} buttonthickness - Button thickness
    */
   this.buttonthickness = scc.defaultTCButtonThickness;
   
   /**
    * @member {number} thumbthickness - Thumb thickness
    */
   this.thumbthickness = scc.defaultTCThumbThickness;
   
   /**
    * @member {number} minscrollingpanesize - Minimum scrolling pane size (the 20 is to leave a little space)
    */
   this.minscrollingpanesize = this.buttonthickness + this.buttonthickness + this.thumbthickness + 20;
};

/**
 * @description Methods for SocialCalc.TableControl prototype
 */

/**
 * @method CreateTableControl
 * @memberof SocialCalc.TableControl
 * @description Creates the table control elements
 * @returns {*} Result of SocialCalc.CreateTableControl
 */
SocialCalc.TableControl.prototype.CreateTableControl = function() {
   return SocialCalc.CreateTableControl(this);
};

/**
 * @method PositionTableControlElements
 * @memberof SocialCalc.TableControl
 * @description Positions the table control elements
 */
SocialCalc.TableControl.prototype.PositionTableControlElements = function() {
   SocialCalc.PositionTableControlElements(this);
};

/**
 * @method ComputeTableControlPositions
 * @memberof SocialCalc.TableControl
 * @description Computes positions for table control elements
 */
SocialCalc.TableControl.prototype.ComputeTableControlPositions = function() {
   SocialCalc.ComputeTableControlPositions(this);
};

/**
 * @namespace Functions
 * @description TableControl related functions
 */

/**
 * @function CreateTableControl
 * @memberof SocialCalc
 * @description Creates and initializes all elements of a table control
 * @param {Object} control - The TableControl instance
 * @returns {HTMLElement} The main control element
 */
SocialCalc.CreateTableControl = function(control) {
   let s, functions, params;
   const AssignID = SocialCalc.AssignID;
   const setStyles = SocialCalc.setStyles;
   const scc = SocialCalc.Constants;
   const TooltipRegister = function(element, etype, vh) {
      if (scc[`s_${etype}Tooltip${vh}`]) {
         SocialCalc.TooltipRegister(element, scc[`s_${etype}Tooltip${vh}`], null);
      }
   };

   const imageprefix = control.editor.imageprefix;
   const vh = control.vertical ? "v" : "h";

   control.main = document.createElement("div");
   s = control.main.style;
   s.height = `${control.vertical ? control.size : control.controlthickness}px`;
   s.width = `${control.vertical ? control.controlthickness : control.size}px`;
   s.zIndex = 0;
   setStyles(control.main, scc.TCmainStyle);
   s.backgroundImage = `url(${imageprefix}main-${vh}.gif)`;
   if (scc.TCmainClass) control.main.className = scc.TCmainClass;

   control.main.style.display = "none"; // wait for layout

   control.endcap = document.createElement("div");
   s = control.endcap.style;
   s.height = `${control.controlthickness}px`;
   s.width = `${control.controlthickness}px`;
   s.zIndex = 1;
   s.overflow = "hidden"; // IE will make the DIV at least font-size height...so use this
   s.position = "absolute";
   setStyles(control.endcap, scc.TCendcapStyle);
   s.backgroundImage = `url(${imageprefix}endcap-${vh}.gif)`;
   if (scc.TCendcapClass) control.endcap.className = scc.TCendcapClass;
   AssignID(control.editor, control.endcap, `endcap${vh}`);

   control.main.appendChild(control.endcap);

   control.paneslider = document.createElement("div");
   s = control.paneslider.style;
   s.height = `${control.vertical ? control.sliderthickness : control.controlthickness}px`;
   s.overflow = "hidden"; // IE will make the DIV at least font-size height...so use this
   s.width = `${control.vertical ? control.controlthickness : control.sliderthickness}px`;
   s.position = "absolute";
   s[control.vertical ? "top" : "left"] = "4px";
   s.zIndex = 3;
   setStyles(control.paneslider, scc.TCpanesliderStyle);
   s.backgroundImage = `url(${imageprefix}paneslider-${vh}.gif)`;
   if (scc.TCpanesliderClass) control.paneslider.className = scc.TCpanesliderClass;
   AssignID(control.editor, control.paneslider, `paneslider${vh}`);
   TooltipRegister(control.paneslider, "paneslider", vh);

   functions = {
      MouseDown: SocialCalc.TCPSDragFunctionStart,
      MouseMove: SocialCalc.TCPSDragFunctionMove,
      MouseUp: SocialCalc.TCPSDragFunctionStop,
      Disabled: function() { return control.editor.busy; }
   };

   functions.control = control; // make sure this is there

   SocialCalc.DragRegister(control.paneslider, control.vertical, !control.vertical, functions);

   control.main.appendChild(control.paneslider);

   control.lessbutton = document.createElement("div");
   s = control.lessbutton.style;
   s.height = `${control.vertical ? control.buttonthickness : control.controlthickness}px`;
   s.width = `${control.vertical ? control.controlthickness : control.buttonthickness}px`;
   s.zIndex = 2;
   s.overflow = "hidden"; // IE will make the DIV at least font-size height...so use this
   s.position = "absolute";
   setStyles(control.lessbutton, scc.TClessbuttonStyle);
   s.backgroundImage = `url(${imageprefix}less-${vh}n.gif)`;
   if (scc.TClessbuttonClass) control.lessbutton.className = scc.TClessbuttonClass;
   AssignID(control.editor, control.lessbutton, `lessbutton${vh}`);

   params = {
      repeatwait: scc.TClessbuttonRepeatWait, 
      repeatinterval: scc.TClessbuttonRepeatInterval,
      normalstyle: `backgroundImage:url(${imageprefix}less-${vh}n.gif);`,
      downstyle: `backgroundImage:url(${imageprefix}less-${vh}d.gif);`,
      hoverstyle: `backgroundImage:url(${imageprefix}less-${vh}h.gif);`
   };
   functions = {
      MouseDown: function() { if (!control.editor.busy) control.editor.ScrollRelative(control.vertical, -1); },
      Repeat: function() { if (!control.editor.busy) control.editor.ScrollRelative(control.vertical, -1); },
      Disabled: function() { return control.editor.busy; }
   };

   SocialCalc.ButtonRegister(control.lessbutton, params, functions);

   control.main.appendChild(control.lessbutton);

   control.morebutton = document.createElement("div");
   s = control.morebutton.style;
   s.height = `${control.vertical ? control.buttonthickness : control.controlthickness}px`;
   s.width = `${control.vertical ? control.controlthickness : control.buttonthickness}px`;
   s.zIndex = 2;
   s.overflow = "hidden"; // IE will make the DIV at least font-size height...so use this
   s.position = "absolute";
   setStyles(control.morebutton, scc.TCmorebuttonStyle);
   s.backgroundImage = `url(${imageprefix}more-${vh}n.gif)`;
   if (scc.TCmorebuttonClass) control.morebutton.className = scc.TCmorebuttonClass;
   AssignID(control.editor, control.morebutton, `morebutton${vh}`);

   params = {
      repeatwait: scc.TCmorebuttonRepeatWait, 
      repeatinterval: scc.TCmorebuttonRepeatInterval,
      normalstyle: `backgroundImage:url(${imageprefix}more-${vh}n.gif);`,
      downstyle: `backgroundImage:url(${imageprefix}more-${vh}d.gif);`,
      hoverstyle: `backgroundImage:url(${imageprefix}more-${vh}h.gif);`
   };
   functions = {
      MouseDown: function() { if (!control.editor.busy) control.editor.ScrollRelative(control.vertical, +1); },
      Repeat: function() { if (!control.editor.busy) control.editor.ScrollRelative(control.vertical, +1); },
      Disabled: function() { return control.editor.busy; }
   };

   SocialCalc.ButtonRegister(control.morebutton, params, functions);

   control.main.appendChild(control.morebutton);

   control.scrollarea = document.createElement("div");
   s = control.scrollarea.style;
   s.height = `${control.controlthickness}px`;
   s.width = `${control.controlthickness}px`;
   s.zIndex = 1;
   s.overflow = "hidden"; // IE will make the DIV at least font-size height...so use this
   s.position = "absolute";
   setStyles(control.scrollarea, scc.TCscrollareaStyle);
   s.backgroundImage = `url(${imageprefix}scrollarea-${vh}.gif)`;
   if (scc.TCscrollareaClass) control.scrollarea.className = scc.TCscrollareaClass;
   AssignID(control.editor, control.scrollarea, `scrollarea${vh}`);

   params = {
      repeatwait: scc.TCscrollareaRepeatWait, 
      repeatinterval: scc.TCscrollareaRepeatWait
   };
   functions = {
      MouseDown: SocialCalc.ScrollAreaClick, 
      Repeat: SocialCalc.ScrollAreaClick,
      Disabled: function() { return control.editor.busy; }
   };
   functions.control = control;

   SocialCalc.ButtonRegister(control.scrollarea, params, functions);

   control.main.appendChild(control.scrollarea);

   control.thumb = document.createElement("div");
   s = control.thumb.style;
   s.height = `${control.vertical ? control.thumbthickness : control.controlthickness}px`;
   s.width = `${control.vertical ? control.controlthickness : control.thumbthickness}px`;
   s.zIndex = 2;
   s.overflow = "hidden"; // IE will make the DIV at least font-size height...so use this
   s.position = "absolute";
   setStyles(control.thumb, scc.TCthumbStyle);
   control.thumb.style.backgroundImage = `url(${imageprefix}thumb-${vh}n.gif)`;
   if (scc.TCthumbClass) control.thumb.className = scc.TCthumbClass;
   AssignID(control.editor, control.thumb, `thumb${vh}`);

   functions = {
      MouseDown: SocialCalc.TCTDragFunctionStart,
      MouseMove: SocialCalc.TCTDragFunctionMove,
      MouseUp: SocialCalc.TCTDragFunctionStop,
      Disabled: function() { return control.editor.busy; }
   };
   functions.control = control; // make sure this is there
   SocialCalc.DragRegister(control.thumb, control.vertical, !control.vertical, functions);

   params = {
      normalstyle: `backgroundImage:url(${imageprefix}thumb-${vh}n.gif)`, 
      name: "Thumb",
      downstyle: `backgroundImage:url(${imageprefix}thumb-${vh}d.gif)`,
      hoverstyle: `backgroundImage:url(${imageprefix}thumb-${vh}h.gif)`
   };
   SocialCalc.ButtonRegister(control.thumb, params, null); // give it button-like visual behavior

   control.main.appendChild(control.thumb);

   return control.main;
};

/**
 * @function ScrollAreaClick
 * @memberof SocialCalc
 * @description Button function to process pageup/down clicks on the scroll area
 * @param {Event} e - The event object
 * @param {Object} buttoninfo - Button information
 * @param {Object} bobj - Button object containing function and control references
 */
SocialCalc.ScrollAreaClick = function(e, buttoninfo, bobj) {
   const control = bobj.functionobj.control;
   const bposition = SocialCalc.GetElementPosition(bobj.element);
   const clickpos = control.vertical ? buttoninfo.clientY : buttoninfo.clientX;
   if (control.editor.busy) { // ignore if busy - wait for next repeat
      return;
   }
   control.editor.PageRelative(control.vertical, clickpos > control.thumbpos ? 1 : -1);

   return;
};
/**
 * @function PositionTableControlElements
 * @memberof SocialCalc
 * @description Positions all the table control elements based on computed positions
 * @param {Object} control - The TableControl instance
 */
SocialCalc.PositionTableControlElements = function(control) {
   let border, realend, thumbpos;

   const editor = control.editor;

   if (control.vertical) {
      border = `${control.controlborder}px`;
      control.endcap.style.top = `${control.endcapstart}px`;
      control.endcap.style.left = border;
      control.paneslider.style.top = `${control.panesliderstart}px`;
      control.paneslider.style.left = border;
      control.lessbutton.style.top = `${control.lessbuttonstart}px`;
      control.lessbutton.style.left = border;
      control.morebutton.style.top = `${control.morebuttonstart}px`;
      control.morebutton.style.left = border;
      control.scrollarea.style.top = `${control.scrollareastart}px`;
      control.scrollarea.style.left = border;
      control.scrollarea.style.height = `${control.scrollareasize}px`;
      realend = Math.max(editor.context.sheetobj.attribs.lastrow, editor.firstscrollingrow + 1);
      thumbpos = ((editor.firstscrollingrow - (editor.lastnonscrollingrow + 1)) * (control.scrollareasize - 3 * control.thumbthickness)) /
         (realend - (editor.lastnonscrollingrow + 1)) + control.scrollareastart - 1;
      thumbpos = Math.floor(thumbpos);
      control.thumb.style.top = `${thumbpos}px`;
      control.thumb.style.left = border;
   } else {
      border = `${control.controlborder}px`;
      control.endcap.style.left = `${control.endcapstart}px`;
      control.endcap.style.top = border;
      control.paneslider.style.left = `${control.panesliderstart}px`;
      control.paneslider.style.top = border;
      control.lessbutton.style.left = `${control.lessbuttonstart}px`;
      control.lessbutton.style.top = border;
      control.morebutton.style.left = `${control.morebuttonstart}px`;
      control.morebutton.style.top = border;
      control.scrollarea.style.left = `${control.scrollareastart}px`;
      control.scrollarea.style.top = border;
      control.scrollarea.style.width = `${control.scrollareasize}px`;
      realend = Math.max(editor.context.sheetobj.attribs.lastcol, editor.firstscrollingcol + 1);
      thumbpos = ((editor.firstscrollingcol - (editor.lastnonscrollingcol + 1)) * (control.scrollareasize - control.thumbthickness)) /
         (realend - editor.lastnonscrollingcol) + control.scrollareastart - 1;
      thumbpos = Math.floor(thumbpos);
      control.thumb.style.left = `${thumbpos}px`;
      control.thumb.style.top = border;
   }
   control.thumbpos = thumbpos;
   control.main.style.display = "block";
};

/**
 * @function ComputeTableControlPositions
 * @memberof SocialCalc
 * @description This routine computes the screen positions and other values needed for laying out
 * the table control elements.
 * @param {Object} control - The TableControl instance
 * @throws {Error} When editor positions haven't been computed yet
 */
SocialCalc.ComputeTableControlPositions = function(control) {
   const editor = control.editor;

   if (!editor.gridposition || !editor.headposition) throw new Error("Can't compute table control positions before editor positions");

   if (control.vertical) {
      control.controlborder = editor.gridposition.left + editor.tablewidth; // border=left position
      control.endcapstart = editor.gridposition.top; // start=top position
      control.panesliderstart = editor.firstscrollingrowtop - control.sliderthickness;
      control.lessbuttonstart = editor.firstscrollingrowtop - 1;
      control.morebuttonstart = editor.gridposition.top + editor.tableheight - control.buttonthickness;
      control.scrollareastart = editor.firstscrollingrowtop - 1 + control.buttonthickness;
      control.scrollareaend = control.morebuttonstart - 1;
      control.scrollareasize = control.scrollareaend - control.scrollareastart + 1;
   } else {
      control.controlborder = editor.gridposition.top + editor.tableheight; // border=top position
      control.endcapstart = editor.gridposition.left; // start=left position
      control.panesliderstart = editor.firstscrollingcolleft - control.sliderthickness;
      control.lessbuttonstart = editor.firstscrollingcolleft - 1;
      control.morebuttonstart = editor.gridposition.left + editor.tablewidth - control.buttonthickness;
      control.scrollareastart = editor.firstscrollingcolleft - 1 + control.buttonthickness;
      control.scrollareaend = control.morebuttonstart - 1;
      control.scrollareasize = control.scrollareaend - control.scrollareastart + 1;
   }
};

/**
 * @namespace TCPS
 * @description TableControl Pan Slider methods
 */

/**
 * @function TCPSDragFunctionStart
 * @memberof SocialCalc
 * @description TableControlPaneSlider function for starting drag
 * @param {Event} event - The drag start event
 * @param {Object} draginfo - Drag information object
 * @param {Object} dobj - Drag object containing control reference
 */
SocialCalc.TCPSDragFunctionStart = function(event, draginfo, dobj) {
   const editor = dobj.functionobj.control.editor;
   const scc = SocialCalc.Constants;

   SocialCalc.DragFunctionStart(event, draginfo, dobj);

   draginfo.trackingline = document.createElement("div");
   draginfo.trackingline.style.height = dobj.vertical ? scc.TCPStrackinglineThickness :
      `${editor.tableheight - (editor.headposition.top - editor.gridposition.top)}px`;
   draginfo.trackingline.style.width = dobj.vertical ? 
      `${editor.tablewidth - (editor.headposition.left - editor.gridposition.left)}px` : scc.TCPStrackinglineThickness;
   draginfo.trackingline.style.backgroundImage = `url(${editor.imageprefix}trackingline-${dobj.vertical ? "v" : "h"}.gif)`;
   if (scc.TCPStrackinglineClass) draginfo.trackingline.className = scc.TCPStrackinglineClass;
   SocialCalc.setStyles(draginfo.trackingline, scc.TCPStrackinglineStyle);

   if (dobj.vertical) {
      const row = SocialCalc.Lookup(draginfo.clientY + dobj.functionobj.control.sliderthickness, editor.rowpositions);
      draginfo.trackingline.style.top = `${editor.rowpositions[row] || editor.headposition.top}px`;
      draginfo.trackingline.style.left = `${editor.headposition.left}px`;
      if (editor.context.rowpanes.length - 1) { // has 2 already
         editor.context.SetRowPaneFirstLast(1, editor.context.rowpanes[0].last + 1, editor.context.rowpanes[0].last + 1);
         editor.FitToEditTable();
         editor.ScheduleRender();
      }
   } else {
      const col = SocialCalc.Lookup(draginfo.clientX + dobj.functionobj.control.sliderthickness, editor.colpositions);
      draginfo.trackingline.style.top = `${editor.headposition.top}px`;
      draginfo.trackingline.style.left = `${editor.colpositions[col] || editor.headposition.left}px`;
      if (editor.context.colpanes.length - 1) { // has 2 already
         editor.context.SetColPaneFirstLast(1, editor.context.colpanes[0].last + 1, editor.context.colpanes[0].last + 1);
         editor.FitToEditTable();
         editor.ScheduleRender();
      }
   }

   editor.griddiv.appendChild(draginfo.trackingline);
};

/**
 * @function TCPSDragFunctionMove
 * @memberof SocialCalc
 * @description Handles drag move events for the table control pane slider
 * @param {Event} event - The drag move event
 * @param {Object} draginfo - Drag information object
 * @param {Object} dobj - Drag object containing control reference
 */
SocialCalc.TCPSDragFunctionMove = function(event, draginfo, dobj) {
   let row, col, max, min;
   const control = dobj.functionobj.control;
   const sliderthickness = control.sliderthickness;
   const editor = control.editor;

   if (dobj.vertical) {
      max = control.morebuttonstart - control.minscrollingpanesize - draginfo.offsetY; // restrict movement
      if (draginfo.clientY > max) draginfo.clientY = max;
      min = editor.headposition.top - sliderthickness - draginfo.offsetY;
      if (draginfo.clientY < min) draginfo.clientY = min;

      row = SocialCalc.Lookup(draginfo.clientY + sliderthickness, editor.rowpositions);
      draginfo.trackingline.style.top = `${editor.rowpositions[row] || editor.headposition.top}px`;
   } else {
      max = control.morebuttonstart - control.minscrollingpanesize - draginfo.offsetX;
      if (draginfo.clientX > max) draginfo.clientX = max;
      min = editor.headposition.left - sliderthickness - draginfo.offsetX;
      if (draginfo.clientX < min) draginfo.clientX = min;

      col = SocialCalc.Lookup(draginfo.clientX + sliderthickness, editor.colpositions);
      draginfo.trackingline.style.left = `${editor.colpositions[col] || editor.headposition.left}px`;
   }

   SocialCalc.DragFunctionPosition(event, draginfo, dobj);
};

/**
 * @function TCPSDragFunctionStop
 * @memberof SocialCalc
 * @description Handles drag stop events for the table control pane slider
 * @param {Event} event - The drag stop event
 * @param {Object} draginfo - Drag information object
 * @param {Object} dobj - Drag object containing control reference
 */
SocialCalc.TCPSDragFunctionStop = function(event, draginfo, dobj) {
   let row, col, max, min;
   const control = dobj.functionobj.control;
   const sliderthickness = control.sliderthickness;
   const editor = control.editor;

   if (dobj.vertical) {
      max = control.morebuttonstart - control.minscrollingpanesize - draginfo.offsetY; // restrict movement
      if (draginfo.clientY > max) draginfo.clientY = max;
      min = editor.headposition.top - sliderthickness - draginfo.offsetY;
      if (draginfo.clientY < min) draginfo.clientY = min;

      row = SocialCalc.Lookup(draginfo.clientY + sliderthickness, editor.rowpositions);
      if (row > editor.context.sheetobj.attribs.lastrow) row = editor.context.sheetobj.attribs.lastrow; // can't extend sheet here
      if (!row || row <= editor.context.rowpanes[0].first) { // set to no panes, leaving first pane settings
         if (editor.context.rowpanes.length > 1) editor.context.rowpanes.length = 1;
      } else if (editor.context.rowpanes.length - 1) { // has 2 already
         if (!editor.timeout) { // not waiting for position calc (so positions could be wrong)
            editor.context.SetRowPaneFirstLast(0, editor.context.rowpanes[0].first, row - 1);
            editor.context.SetRowPaneFirstLast(1, row, row);
         }
      } else {
         editor.context.SetRowPaneFirstLast(0, editor.context.rowpanes[0].first, row - 1);
         editor.context.SetRowPaneFirstLast(1, row, row);
      }
   } else {
      max = control.morebuttonstart - control.minscrollingpanesize - draginfo.offsetX;
      if (draginfo.clientX > max) draginfo.clientX = max;
      min = editor.headposition.left - sliderthickness - draginfo.offsetX;
      if (draginfo.clientX < min) draginfo.clientX = min;

      col = SocialCalc.Lookup(draginfo.clientX + sliderthickness, editor.colpositions);
      if (col > editor.context.sheetobj.attribs.lastcol) col = editor.context.sheetobj.attribs.lastcol; // can't extend sheet here
      if (!col || col <= editor.context.colpanes[0].first) { // set to no panes, leaving first pane settings
         if (editor.context.colpanes.length > 1) editor.context.colpanes.length = 1;
      } else if (editor.context.colpanes.length - 1) { // has 2 already
         if (!editor.timeout) { // not waiting for position calc (so positions could be wrong)
            editor.context.SetColPaneFirstLast(0, editor.context.colpanes[0].first, col - 1);
            editor.context.SetColPaneFirstLast(1, col, col);
         }
      } else {
         editor.context.SetColPaneFirstLast(0, editor.context.colpanes[0].first, col - 1);
         editor.context.SetColPaneFirstLast(1, col, col);
      }
   }

   editor.FitToEditTable();

   editor.griddiv.removeChild(draginfo.trackingline);

   editor.ScheduleRender();
};
/**
 * @namespace TCT
 * @description TableControl Thumb methods
 * 
 * @todo Note: Need to make start use same code as move/stop for determining row/col, since stop will set that
 * @todo Note: Need to make start/move/stop use positioning code that corresponds closer to
 *       ComputeTableControlPositions calculations.
 */

/**
 * @function TCTDragFunctionStart
 * @memberof SocialCalc
 * @description TableControlThumb function for starting drag
 * @param {Event} event - The drag start event
 * @param {Object} draginfo - Drag information object
 * @param {Object} dobj - Drag object containing control reference
 */
SocialCalc.TCTDragFunctionStart = function(event, draginfo, dobj) {
   let rowpane, colpane, row, col;

   const control = dobj.functionobj.control;
   const editor = control.editor;
   const scc = SocialCalc.Constants;

   SocialCalc.DragFunctionStart(event, draginfo, dobj);

   if (draginfo.thumbstatus) { // get rid of old one if mouseup was out of window
      if (draginfo.thumbstatus.rowmsgele) draginfo.thumbstatus.rowmsgele = null;
      if (draginfo.thumbstatus.rowpreviewele) draginfo.thumbstatus.rowpreviewele = null;
      editor.toplevel.removeChild(draginfo.thumbstatus);
      draginfo.thumbstatus = null;
   }

   draginfo.thumbstatus = document.createElement("div");

   if (dobj.vertical) {
      if (scc.TCTDFSthumbstatusvClass) draginfo.thumbstatus.className = scc.TCTDFSthumbstatusvClass;
      SocialCalc.setStyles(draginfo.thumbstatus, scc.TCTDFSthumbstatusvStyle);
      draginfo.thumbstatus.style.top = `${draginfo.clientY + scc.TCTDFStopOffsetv}px`;
      draginfo.thumbstatus.style.left = `${control.controlborder - 10 - (editor.tablewidth / 2)}px`;
      draginfo.thumbstatus.style.width = `${editor.tablewidth / 2}px`;

      draginfo.thumbcontext = new SocialCalc.RenderContext(editor.context.sheetobj);
      draginfo.thumbcontext.showGrid = true;
      draginfo.thumbcontext.rowpanes = [{first: 1, last: 1}];
      const pane = editor.context.colpanes[editor.context.colpanes.length - 1];
      draginfo.thumbcontext.colpanes = [{first: pane.first, last: pane.last}];
      draginfo.thumbstatus.innerHTML = `<table cellspacing="0" cellpadding="0"><tr><td valign="top" style="${scc.TCTDFSthumbstatusrownumStyle}" class="${scc.TCTDFSthumbstatusrownumClass}"><div>msg</div></td><td valign="top"><div style="overflow:hidden;">preview</div></td></tr></table>`;
      draginfo.thumbstatus.rowmsgele = draginfo.thumbstatus.firstChild.firstChild.firstChild.firstChild.firstChild;
      draginfo.thumbstatus.rowpreviewele = draginfo.thumbstatus.firstChild.firstChild.firstChild.childNodes[1].firstChild;
      editor.toplevel.appendChild(draginfo.thumbstatus);
      SocialCalc.TCTDragFunctionRowSetStatus(draginfo, editor, editor.firstscrollingrow || 1);
   } else {
      if (scc.TCTDFSthumbstatushClass) draginfo.thumbstatus.className = scc.TCTDFSthumbstatushClass;
      SocialCalc.setStyles(draginfo.thumbstatus, scc.TCTDFSthumbstatushStyle);
      draginfo.thumbstatus.style.top = `${control.controlborder + scc.TCTDFStopOffseth}px`;
      draginfo.thumbstatus.style.left = `${draginfo.clientX + scc.TCTDFSleftOffseth}px`;
      editor.toplevel.appendChild(draginfo.thumbstatus);
      draginfo.thumbstatus.innerHTML = scc.s_TCTDFthumbstatusPrefixh + SocialCalc.rcColname(editor.firstscrollingcol);
   }
};

/**
 * @function TCTDragFunctionRowSetStatus
 * @memberof SocialCalc
 * @description Render partial row for thumb drag status
 * @param {Object} draginfo - Drag information object
 * @param {Object} editor - The table editor instance
 * @param {number} row - The row number to render
 */
SocialCalc.TCTDragFunctionRowSetStatus = function(draginfo, editor, row) {
   const scc = SocialCalc.Constants;
   const msg = `${scc.s_TCTDFthumbstatusPrefixv}${row} `;

   draginfo.thumbstatus.rowmsgele.innerHTML = msg;

   draginfo.thumbcontext.rowpanes = [{first: row, last: row}];
   draginfo.thumbrowshown = row;

   const ele = draginfo.thumbcontext.RenderSheet(draginfo.thumbstatus.rowpreviewele.firstChild, {type: "html"});
};

/**
 * @function TCTDragFunctionMove
 * @memberof SocialCalc
 * @description Handles drag move events for the table control thumb
 * @param {Event} event - The drag move event
 * @param {Object} draginfo - Drag information object
 * @param {Object} dobj - Drag object containing control reference
 */
SocialCalc.TCTDragFunctionMove = function(event, draginfo, dobj) {
   let first, msg;
   const control = dobj.functionobj.control;
   const thumbthickness = control.thumbthickness;
   const editor = control.editor;
   const scc = SocialCalc.Constants;

   if (dobj.vertical) {
      if (draginfo.clientY > control.scrollareaend - draginfo.offsetY - control.thumbthickness + 2)
         draginfo.clientY = control.scrollareaend - draginfo.offsetY - control.thumbthickness + 2;
      if (draginfo.clientY < control.scrollareastart - draginfo.offsetY - 1)
         draginfo.clientY = control.scrollareastart - draginfo.offsetY - 1;
      draginfo.thumbstatus.style.top = `${draginfo.clientY}px`;

      first =
         ((draginfo.clientY + draginfo.offsetY - control.scrollareastart + 1) / (control.scrollareasize - control.thumbthickness))
         * (editor.context.sheetobj.attribs.lastrow - editor.lastnonscrollingrow)
         + editor.lastnonscrollingrow + 1;
      first = Math.floor(first);
      if (first <= editor.lastnonscrollingrow) first = editor.lastnonscrollingrow + 1;
      if (first > editor.context.sheetobj.attribs.lastrow) first = editor.context.sheetobj.attribs.lastrow;
//      msg = scc.s_TCTDFthumbstatusPrefixv+first;
      if (first !== draginfo.thumbrowshown) {
         SocialCalc.TCTDragFunctionRowSetStatus(draginfo, editor, first);
      }
   } else {
      if (draginfo.clientX > control.scrollareaend - draginfo.offsetX - control.thumbthickness + 2)
         draginfo.clientX = control.scrollareaend - draginfo.offsetX - control.thumbthickness + 2;
      if (draginfo.clientX < control.scrollareastart - draginfo.offsetX - 1)
         draginfo.clientX = control.scrollareastart - draginfo.offsetX - 1;
      draginfo.thumbstatus.style.left = `${draginfo.clientX}px`;

      first =
         ((draginfo.clientX + draginfo.offsetX - control.scrollareastart + 1) / (control.scrollareasize - control.thumbthickness))
         * (editor.context.sheetobj.attribs.lastcol - editor.lastnonscrollingcol)
         + editor.lastnonscrollingcol + 1;
      first = Math.floor(first);
      if (first <= editor.lastnonscrollingcol) first = editor.lastnonscrollingcol + 1;
      if (first > editor.context.sheetobj.attribs.lastcol) first = editor.context.sheetobj.attribs.lastcol;
      msg = scc.s_TCTDFthumbstatusPrefixh + SocialCalc.rcColname(first);
      draginfo.thumbstatus.innerHTML = msg;
   }

   SocialCalc.DragFunctionPosition(event, draginfo, dobj);
};

/**
 * @function TCTDragFunctionStop
 * @memberof SocialCalc
 * @description Handles drag stop events for the table control thumb
 * @param {Event} event - The drag stop event
 * @param {Object} draginfo - Drag information object
 * @param {Object} dobj - Drag object containing control reference
 */
SocialCalc.TCTDragFunctionStop = function(event, draginfo, dobj) {
   let first;
   const control = dobj.functionobj.control;
   const editor = control.editor;

   if (dobj.vertical) {
      first =
         ((draginfo.clientY + draginfo.offsetY - control.scrollareastart + 1) / (control.scrollareasize - control.thumbthickness))
         * (editor.context.sheetobj.attribs.lastrow - editor.lastnonscrollingrow)
         + editor.lastnonscrollingrow + 1;
      first = Math.floor(first);
      if (first <= editor.lastnonscrollingrow) first = editor.lastnonscrollingrow + 1;
      if (first > editor.context.sheetobj.attribs.lastrow) first = editor.context.sheetobj.attribs.lastrow;

      editor.context.SetRowPaneFirstLast(editor.context.rowpanes.length - 1, first, first + 1);
   } else {
      first =
         ((draginfo.clientX + draginfo.offsetX - control.scrollareastart + 1) / (control.scrollareasize - control.thumbthickness))
         * (editor.context.sheetobj.attribs.lastcol - editor.lastnonscrollingcol)
         + editor.lastnonscrollingcol + 1;
      first = Math.floor(first);
      if (first <= editor.lastnonscrollingcol) first = editor.lastnonscrollingcol + 1;
      if (first > editor.context.sheetobj.attribs.lastcol) first = editor.context.sheetobj.attribs.lastcol;

      editor.context.SetColPaneFirstLast(editor.context.colpanes.length - 1, first, first + 1);
   }

   editor.FitToEditTable();

   if (draginfo.thumbstatus.rowmsgele) draginfo.thumbstatus.rowmsgele = null;
   if (draginfo.thumbstatus.rowpreviewele) draginfo.thumbstatus.rowpreviewele = null;
   editor.toplevel.removeChild(draginfo.thumbstatus);
   draginfo.thumbstatus = null;

   editor.ScheduleRender();
};

/**
 * @namespace DragFunctions
 * @description Dragging functions for general drag operations
 */

/**
 * @object DragInfo
 * @memberof SocialCalc
 * @description There is only one of these -- no "new" is done.
 * Only one dragging operation can be active at a time.
 * The registeredElements array is used to decide which item to drag.
 */
SocialCalc.DragInfo = {
   /**
    * @description One item for each draggable thing, each an object with:
    *    .element, .vertical, .horizontal, .functionobj
    * @type {Array<Object>}
    */
   registeredElements: [],

   // Items used during a drag

   /**
    * @member {Object|null} draggingElement - Item being processed (.element is the actual element)
    */
   draggingElement: null,
   
   /**
    * @member {number} startX - Starting X position
    */
   startX: 0,
   
   /**
    * @member {number} startY - Starting Y position
    */
   startY: 0,
   
   /**
    * @member {number} startZ - Starting Z position
    */
   startZ: 0,
   
   /**
    * @member {number} clientX - Modifiable version to restrict movement
    */
   clientX: 0,
   
   /**
    * @member {number} clientY - Modifiable version to restrict movement
    */
   clientY: 0,
   
   /**
    * @member {number} offsetX - X offset
    */
   offsetX: 0,
   
   /**
    * @member {number} offsetY - Y offset
    */
   offsetY: 0,
   
   /**
    * @member {number} horizontalScroll - Retrieved at drag start
    */
   horizontalScroll: 0,
   
   /**
    * @member {number} verticalScroll - Retrieved at drag start
    */
   verticalScroll: 0
};

/**
 * @function DragRegister
 * @memberof SocialCalc
 * @description Make element draggable
 * The functionobj defaults to moving the element constrained only by vertical and horizontal settings.
 * @param {HTMLElement} element - The element to make draggable
 * @param {boolean} vertical - Whether vertical dragging is allowed
 * @param {boolean} horizontal - Whether horizontal dragging is allowed
 * @param {Object} functionobj - Function object with MouseDown, MouseMove, MouseUp, and Disabled properties
 * @throws {string} When browser is not supported
 */
SocialCalc.DragRegister = function(element, vertical, horizontal, functionobj) {
   const draginfo = SocialCalc.DragInfo;

   if (!functionobj) {
      functionobj = {
         MouseDown: SocialCalc.DragFunctionStart, 
         MouseMove: SocialCalc.DragFunctionPosition,
         MouseUp: SocialCalc.DragFunctionPosition,
         Disabled: null
      };
   }

   draginfo.registeredElements.push({
      element: element, 
      vertical: vertical, 
      horizontal: horizontal, 
      functionobj: functionobj
   });

   if (element.addEventListener) { // DOM Level 2 -- Firefox, et al
      element.addEventListener("mousedown", SocialCalc.DragMouseDown, false);
   } else if (element.attachEvent) { // IE 5+
      element.attachEvent("onmousedown", SocialCalc.DragMouseDown);
   } else { // don't handle this
      throw SocialCalc.Constants.s_BrowserNotSupported;
   }
};
/**
 * @function DragUnregister
 * @memberof SocialCalc
 * @description Remove object from drag registration list
 * @param {HTMLElement} element - The element to unregister from dragging
 */
SocialCalc.DragUnregister = function(element) {
   const draginfo = SocialCalc.DragInfo;

   let i;

   if (!element) return;

   for (i = 0; i < draginfo.registeredElements.length; i++) {
      if (draginfo.registeredElements[i].element === element) {
         draginfo.registeredElements.splice(i, 1);
         if (element.removeEventListener) { // DOM Level 2 -- Firefox, et al
            element.removeEventListener("mousedown", SocialCalc.DragMouseDown, false);
         } else { // IE 5+
            element.detachEvent("onmousedown", SocialCalc.DragMouseDown);
         }
         return;
      }
   }

   return; // ignore if not in list
};

/**
 * @function DragMouseDown
 * @memberof SocialCalc
 * @description Handles mouse down events for draggable elements
 * @param {Event} event - The mouse down event
 * @returns {boolean} Returns false to prevent default behavior
 */
SocialCalc.DragMouseDown = function(event) {
   const e = event || window.event;

   const draginfo = SocialCalc.DragInfo;

   const dobj = SocialCalc.LookupElement(e.target || e.srcElement, draginfo.registeredElements);
   if (!dobj) return;

   if (dobj && dobj.functionobj && dobj.functionobj.Disabled) {
      if (dobj.functionobj.Disabled(e, draginfo, dobj)) {
         return;
      }
   }

   draginfo.draggingElement = dobj;

   const viewportinfo = SocialCalc.GetViewportInfo();
   draginfo.horizontalScroll = viewportinfo.horizontalScroll;
   draginfo.verticalScroll = viewportinfo.verticalScroll;

   draginfo.clientX = e.clientX + draginfo.horizontalScroll; // get document-relative coordinates
   draginfo.clientY = e.clientY + draginfo.verticalScroll;
   draginfo.startX = draginfo.clientX;
   draginfo.startY = draginfo.clientY;
   draginfo.startZ = dobj.element.style.zIndex;
   draginfo.offsetX = 0;
   draginfo.offsetY = 0;

   dobj.element.style.zIndex = "100";

   // Event code from JavaScript, Flanagan, 5th Edition, pg. 422
   if (document.addEventListener) { // DOM Level 2 -- Firefox, et al
      document.addEventListener("mousemove", SocialCalc.DragMouseMove, true); // capture everywhere
      document.addEventListener("mouseup", SocialCalc.DragMouseUp, true);
   } else if (dobj.element.attachEvent) { // IE 5+
      dobj.element.setCapture();
      dobj.element.attachEvent("onmousemove", SocialCalc.DragMouseMove);
      dobj.element.attachEvent("onmouseup", SocialCalc.DragMouseUp);
      dobj.element.attachEvent("onlosecapture", SocialCalc.DragMouseUp);
   }
   if (e.stopPropagation) e.stopPropagation(); // DOM Level 2
   else e.cancelBubble = true; // IE 5+
   if (e.preventDefault) e.preventDefault(); // DOM Level 2
   else e.returnValue = false; // IE 5+

   if (dobj && dobj.functionobj && dobj.functionobj.MouseDown) dobj.functionobj.MouseDown(e, draginfo, dobj);

   return false;
};

/**
 * @function DragMouseMove
 * @memberof SocialCalc
 * @description Handles mouse move events during dragging
 * @param {Event} event - The mouse move event
 * @returns {boolean} Returns false to prevent default behavior
 */
SocialCalc.DragMouseMove = function(event) {
   const e = event || window.event;

   const draginfo = SocialCalc.DragInfo;
   draginfo.clientX = e.clientX + draginfo.horizontalScroll;
   draginfo.clientY = e.clientY + draginfo.verticalScroll;

   const dobj = draginfo.draggingElement;

   if (e.stopPropagation) e.stopPropagation(); // DOM Level 2
   else e.cancelBubble = true; // IE 5+

   if (dobj && dobj.functionobj && dobj.functionobj.MouseMove) dobj.functionobj.MouseMove(e, draginfo, dobj);

   return false;
};

/**
 * @function DragMouseUp
 * @memberof SocialCalc
 * @description Handles mouse up events to end dragging
 * @param {Event} event - The mouse up event
 * @returns {boolean} Returns false to prevent default behavior
 */
SocialCalc.DragMouseUp = function(event) {
   const e = event || window.event;

   const draginfo = SocialCalc.DragInfo;
   draginfo.clientX = e.clientX + draginfo.horizontalScroll;
   draginfo.clientY = e.clientY + draginfo.verticalScroll;

   const dobj = draginfo.draggingElement;

   dobj.element.style.zIndex = draginfo.startZ;

   if (dobj && dobj.functionobj && dobj.functionobj.MouseUp) dobj.functionobj.MouseUp(e, draginfo, dobj);

   if (e.stopPropagation) e.stopPropagation(); // DOM Level 2
   else e.cancelBubble = true; // IE 5+

   if (document.removeEventListener) { // DOM Level 2
      document.removeEventListener("mousemove", SocialCalc.DragMouseMove, true);
      document.removeEventListener("mouseup", SocialCalc.DragMouseUp, true);
      // Note: In old (1.5?) versions of Firefox, this causes the browser to skip the MouseUp for
      // the button code. https://bugzilla.mozilla.org/show_bug.cgi?id=174320
      // Firefox 1.5 is <1% share (http://marketshare.hitslink.com/report.aspx?qprid=7)
   } else if (dobj.element.detachEvent) { // IE
      dobj.element.detachEvent("onlosecapture", SocialCalc.DragMouseUp);
      dobj.element.detachEvent("onmouseup", SocialCalc.DragMouseUp);
      dobj.element.detachEvent("onmousemove", SocialCalc.DragMouseMove);
      dobj.element.releaseCapture();
   }

   draginfo.draggingElement = null;

   return false;
};

/**
 * @function DragFunctionStart
 * @memberof SocialCalc
 * @description Default drag function for starting drag operations
 * @param {Event} event - The drag start event
 * @param {Object} draginfo - Drag information object
 * @param {Object} dobj - Drag object containing element and configuration
 */
SocialCalc.DragFunctionStart = function(event, draginfo, dobj) {
   let val;
   const element = dobj.functionobj.positionobj || dobj.element;

   val = element.style.top.match(/\d*/);
   draginfo.offsetY = (val ? val[0] - 0 : 0) - draginfo.clientY;
   val = element.style.left.match(/\d*/);
   draginfo.offsetX = (val ? val[0] - 0 : 0) - draginfo.clientX;
};

/**
 * @function DragFunctionPosition
 * @memberof SocialCalc
 * @description Default drag function for positioning elements during drag
 * @param {Event} event - The drag event
 * @param {Object} draginfo - Drag information object
 * @param {Object} dobj - Drag object containing element and configuration
 */
SocialCalc.DragFunctionPosition = function(event, draginfo, dobj) {
   const element = dobj.functionobj.positionobj || dobj.element;

   if (dobj.vertical) element.style.top = `${draginfo.clientY + draginfo.offsetY}px`;
   if (dobj.horizontal) element.style.left = `${draginfo.clientX + draginfo.offsetX}px`;
};
/**
 * @namespace TooltipFunctions
 * @description Tooltip functions for SocialCalc
 */

/**
 * @object TooltipInfo
 * @memberof SocialCalc
 * @description There is only one of these -- no "new" is done.
 * Only one tooltip operation can be active at a time.
 * The registeredElements array is used to identify items.
 */
SocialCalc.TooltipInfo = {
   /**
    * @description One item for each element with a tooltip, each an object with:
    *    .element, .tiptext, .functionobj
    * Currently .functionobj can only contain .offsetx and .offsety.
    * If present they are used instead of the default ones.
    * @type {Array<Object>}
    */
   registeredElements: [],

   /**
    * @member {boolean} registered - If true, an event handler has been registered for this functionality
    */
   registered: false,

   // Items used during hover over an element

   /**
    * @member {Object|null} tooltipElement - Item being processed (.element is the actual element)
    */
   tooltipElement: null,
   
   /**
    * @member {number|null} timer - Timer object waiting to see if holding over element
    */
   timer: null,
   
   /**
    * @member {HTMLElement|null} popupElement - Tooltip element being displayed
    */
   popupElement: null,
   
   /**
    * @member {number} clientX - Modifiable version to restrict movement
    */
   clientX: 0,
   
   /**
    * @member {number} clientY - Modifiable version to restrict movement
    */
   clientY: 0,
   
   /**
    * @member {number} offsetX - Modifiable version to allow positioning
    */
   offsetX: SocialCalc.Constants.TooltipOffsetX,
   
   /**
    * @member {number} offsetY - Modifiable version to allow positioning
    */
   offsetY: SocialCalc.Constants.TooltipOffsetY
};

/**
 * @function TooltipRegister
 * @memberof SocialCalc
 * @description Make element have a tooltip
 * @param {HTMLElement} element - The element to add tooltip to
 * @param {string} tiptext - The tooltip text to display
 * @param {Object} functionobj - Function object containing offsetx and offsety properties
 * @throws {string} When browser is not supported
 */
SocialCalc.TooltipRegister = function(element, tiptext, functionobj) {
   const tooltipinfo = SocialCalc.TooltipInfo;
   tooltipinfo.registeredElements.push({
      element: element, 
      tiptext: tiptext, 
      functionobj: functionobj
   });

   if (tooltipinfo.registered) return; // only need to add event listener once

   if (document.addEventListener) { // DOM Level 2 -- Firefox, et al
      document.addEventListener("mousemove", SocialCalc.TooltipMouseMove, false);
   } else if (document.attachEvent) { // IE 5+
      document.attachEvent("onmousemove", SocialCalc.TooltipMouseMove);
   } else { // don't handle this
      throw SocialCalc.Constants.s_BrowserNotSupported;
   }

   tooltipinfo.registered = true; // remember

   return;
};

/**
 * @function TooltipMouseMove
 * @memberof SocialCalc
 * @description Handles mouse move events for tooltip functionality
 * @param {Event} event - The mouse move event
 */
SocialCalc.TooltipMouseMove = function(event) {
   const e = event || window.event;

   const tooltipinfo = SocialCalc.TooltipInfo;

   tooltipinfo.viewport = SocialCalc.GetViewportInfo();
   tooltipinfo.clientX = e.clientX + tooltipinfo.viewport.horizontalScroll;
   tooltipinfo.clientY = e.clientY + tooltipinfo.viewport.verticalScroll;

   const tobj = SocialCalc.LookupElement(e.target || e.srcElement, tooltipinfo.registeredElements);

   if (tooltipinfo.timer) { // waiting to see if holding still: didn't hold still
      window.clearTimeout(tooltipinfo.timer); // cancel timer
      tooltipinfo.timer = null;
   }

   if (tooltipinfo.popupElement) { // currently displaying a tip: hide it
      SocialCalc.TooltipHide();
   }

   tooltipinfo.tooltipElement = tobj || null;

   if (!tobj || SocialCalc.ButtonInfo.buttonDown) return; // if not an object with a tip or a "button" is down, ignore

   tooltipinfo.timer = window.setTimeout(SocialCalc.TooltipWaitDone, 700);

   if (tooltipinfo.tooltipElement.element.addEventListener) { // Register event for mouse down which cancels tooltip stuff
      tooltipinfo.tooltipElement.element.addEventListener("mousedown", SocialCalc.TooltipMouseDown, false);
   } else if (tooltipinfo.tooltipElement.element.attachEvent) { // IE
      tooltipinfo.tooltipElement.element.attachEvent("onmousedown", SocialCalc.TooltipMouseDown);
   }

   return;
};

/**
 * @function TooltipMouseDown
 * @memberof SocialCalc
 * @description Handles mouse down events to cancel tooltip display
 * @param {Event} event - The mouse down event
 */
SocialCalc.TooltipMouseDown = function(event) {
   const e = event || window.event;

   const tooltipinfo = SocialCalc.TooltipInfo;

   if (tooltipinfo.timer) {
      window.clearTimeout(tooltipinfo.timer); // cancel timer
      tooltipinfo.timer = null;
   }

   if (tooltipinfo.popupElement) { // currently displaying a tip: hide it
      SocialCalc.TooltipHide();
   }

   if (tooltipinfo.tooltipElement) {
      if (tooltipinfo.tooltipElement.element.removeEventListener) { // DOM Level 2 -- Firefox, et al
         tooltipinfo.tooltipElement.element.removeEventListener("mousedown", SocialCalc.TooltipMouseDown, false);
      } else if (tooltipinfo.tooltipElement.element.attachEvent) { // IE 5+
         tooltipinfo.tooltipElement.element.detachEvent("onmousedown", SocialCalc.TooltipMouseDown);
      }
      tooltipinfo.tooltipElement = null;
   }

   return;
};

/**
 * @function TooltipDisplay
 * @memberof SocialCalc
 * @description Displays a tooltip for the specified element
 * @param {Object} tobj - Tooltip object containing element, tiptext, and functionobj
 */
SocialCalc.TooltipDisplay = function(tobj) {
   const tooltipinfo = SocialCalc.TooltipInfo;
   const scc = SocialCalc.Constants;
   const offsetX = (tobj.functionobj && ((typeof tobj.functionobj.offsetx) === "number")) ? tobj.functionobj.offsetx : tooltipinfo.offsetX;
   const offsetY = (tobj.functionobj && ((typeof tobj.functionobj.offsety) === "number")) ? tobj.functionobj.offsety : tooltipinfo.offsetY;

   tooltipinfo.popupElement = document.createElement("div");
   if (scc.TDpopupElementClass) tooltipinfo.popupElement.className = scc.TDpopupElementClass;
   SocialCalc.setStyles(tooltipinfo.popupElement, scc.TDpopupElementStyle);

   tooltipinfo.popupElement.innerHTML = tobj.tiptext;

   if (tooltipinfo.clientX > tooltipinfo.viewport.width / 2) { // on right side of screen
      tooltipinfo.popupElement.style.bottom = `${tooltipinfo.viewport.height - tooltipinfo.clientY + offsetY}px`;
      tooltipinfo.popupElement.style.right = `${tooltipinfo.viewport.width - tooltipinfo.clientX + offsetX}px`;
   } else { // on left side of screen
      tooltipinfo.popupElement.style.bottom = `${tooltipinfo.viewport.height - tooltipinfo.clientY + offsetY}px`;
      tooltipinfo.popupElement.style.left = `${tooltipinfo.clientX + offsetX}px`;
   }

   if (tooltipinfo.clientY < 50) { // make sure fits on screen if nothing above grid
      tooltipinfo.popupElement.style.bottom = `${tooltipinfo.viewport.height - tooltipinfo.clientY + offsetY - 50}px`;
   }

   document.body.appendChild(tooltipinfo.popupElement);
};

/**
 * @function TooltipHide
 * @memberof SocialCalc
 * @description Hides the currently displayed tooltip
 */
SocialCalc.TooltipHide = function() {
   const tooltipinfo = SocialCalc.TooltipInfo;

   if (tooltipinfo.popupElement) {
      tooltipinfo.popupElement.parentNode.removeChild(tooltipinfo.popupElement);
      tooltipinfo.popupElement = null;
   }
};

/**
 * @function TooltipWaitDone
 * @memberof SocialCalc
 * @description Called when the tooltip wait timer expires, displays the tooltip
 */
SocialCalc.TooltipWaitDone = function() {
   const tooltipinfo = SocialCalc.TooltipInfo;

   tooltipinfo.timer = null;

   SocialCalc.TooltipDisplay(tooltipinfo.tooltipElement);
};
/**
 * @namespace ButtonFunctions
 * @description Button functions for SocialCalc
 */

/**
 * @object ButtonInfo
 * @memberof SocialCalc
 * @description There is only one of these -- no "new" is done.
 * Only one button operation can be active at a time.
 * The registeredElements array is used to identify items.
 */
SocialCalc.ButtonInfo = {
   /**
    * @description One item for each clickable element, each an object with:
    *    .element, .normalstyle, .hoverstyle, .downstyle, .repeatinterval, .functionobj
    *
    * .functionobj is an object with optional function objects for:
    *    mouseover, mouseout, mousedown, repeatinterval, mouseup, disabled
    * @type {Array<Object>}
    */
   registeredElements: [],

   // Items used during hover over an element, clicking, repeating, etc.

   /**
    * @member {Object|null} buttonElement - Item being processed, hover or down (.element is the actual element)
    */
   buttonElement: null,
   
   /**
    * @member {boolean} doingHover - True if mouse is over one of our elements
    */
   doingHover: false,
   
   /**
    * @member {boolean} buttonDown - True if button down and buttonElement not null
    */
   buttonDown: false,
   
   /**
    * @member {number|null} timer - Timer object for repeating
    */
   timer: null,

   // Used while processing an event

   /**
    * @member {number} horizontalScroll - Horizontal scroll position
    */
   horizontalScroll: 0,
   
   /**
    * @member {number} verticalScroll - Vertical scroll position
    */
   verticalScroll: 0,
   
   /**
    * @member {number} clientX - Client X coordinate
    */
   clientX: 0,
   
   /**
    * @member {number} clientY - Client Y coordinate
    */
   clientY: 0
};

/**
 * @function ButtonRegister
 * @memberof SocialCalc
 * @description Make element clickable
 * The arguments (other than element) may be null (meaning no change for style and no repeat)
 * The paramobj has the optional normalstyle, hoverstyle, downstyle, repeatwait, repeatinterval settings
 * @param {HTMLElement} element - The element to make clickable
 * @param {Object} paramobj - Parameter object with style and repeat settings
 * @param {Object} functionobj - Function object with event handlers
 * @throws {string} When browser is not supported
 */
SocialCalc.ButtonRegister = function(element, paramobj, functionobj) {
   const buttoninfo = SocialCalc.ButtonInfo;

   if (!paramobj) paramobj = {};

   buttoninfo.registeredElements.push({
      name: paramobj.name, 
      element: element, 
      normalstyle: paramobj.normalstyle, 
      hoverstyle: paramobj.hoverstyle, 
      downstyle: paramobj.downstyle,
      repeatwait: paramobj.repeatwait, 
      repeatinterval: paramobj.repeatinterval, 
      functionobj: functionobj
   });

   if (element.addEventListener) { // DOM Level 2 -- Firefox, et al
      element.addEventListener("mousedown", SocialCalc.ButtonMouseDown, false);
      element.addEventListener("mouseover", SocialCalc.ButtonMouseOver, false);
      element.addEventListener("mouseout", SocialCalc.ButtonMouseOut, false);
   } else if (element.attachEvent) { // IE 5+
      element.attachEvent("onmousedown", SocialCalc.ButtonMouseDown);
      element.attachEvent("onmouseover", SocialCalc.ButtonMouseOver);
      element.attachEvent("onmouseout", SocialCalc.ButtonMouseOut);
   } else { // don't handle this
      throw SocialCalc.Constants.s_BrowserNotSupported;
   }

   return;
};

/**
 * @function ButtonMouseOver
 * @memberof SocialCalc
 * @description Handles mouse over events for button elements
 * @param {Event} event - The mouse over event
 */
SocialCalc.ButtonMouseOver = function(event) {
   const e = event || window.event;

   const buttoninfo = SocialCalc.ButtonInfo;

   const bobj = SocialCalc.LookupElement(e.target || e.srcElement, buttoninfo.registeredElements);

   if (!bobj) return;

   if (buttoninfo.buttonDown) {
      if (buttoninfo.buttonElement === bobj) {
         buttoninfo.doingHover = true; // keep track whether we are on the pressed button or not
      }
      return;
   }

   if (buttoninfo.buttonElement &&
          buttoninfo.buttonElement !== bobj && buttoninfo.doingHover) { // moved to a new one, undo hover there
      SocialCalc.setStyles(buttoninfo.buttonElement.element, buttoninfo.buttonElement.normalstyle);
   }

   buttoninfo.buttonElement = bobj; // remember this one is hovering
   buttoninfo.doingHover = true;

   SocialCalc.setStyles(bobj.element, bobj.hoverstyle); // set style (if provided)

   if (bobj && bobj.functionobj && bobj.functionobj.MouseOver) bobj.functionobj.MouseOver(e, buttoninfo, bobj);

   return;
};

/**
 * @function ButtonMouseOut
 * @memberof SocialCalc
 * @description Handles mouse out events for button elements
 * @param {Event} event - The mouse out event
 */
SocialCalc.ButtonMouseOut = function(event) {
   const e = event || window.event;

   const buttoninfo = SocialCalc.ButtonInfo;

   if (buttoninfo.buttonDown) {
      buttoninfo.doingHover = false; // keep track of overs and outs
      return;
   }

   const bobj = SocialCalc.LookupElement(e.target || e.srcElement, buttoninfo.registeredElements);

   if (buttoninfo.doingHover) { // if there was a hover, undo it
      if (buttoninfo.buttonElement)
         SocialCalc.setStyles(buttoninfo.buttonElement.element, buttoninfo.buttonElement.normalstyle);
      buttoninfo.buttonElement = null;
      buttoninfo.doingHover = false;
   }

   if (bobj && bobj.functionobj && bobj.functionobj.MouseOut) bobj.functionobj.MouseOut(e, buttoninfo, bobj);

   return;
};

/**
 * @function ButtonMouseDown
 * @memberof SocialCalc
 * @description Handles mouse down events for button elements
 * @param {Event} event - The mouse down event
 */
SocialCalc.ButtonMouseDown = function(event) {
   const e = event || window.event;

   const buttoninfo = SocialCalc.ButtonInfo;

   const viewportinfo = SocialCalc.GetViewportInfo();

   const bobj = SocialCalc.LookupElement(e.target || e.srcElement, buttoninfo.registeredElements);

   if (!bobj) return; // not one of our elements

   if (bobj && bobj.functionobj && bobj.functionobj.Disabled) {
      if (bobj.functionobj.Disabled(e, buttoninfo, bobj)) {
         return;
      }
   }

   buttoninfo.buttonElement = bobj;
   buttoninfo.buttonDown = true;

   SocialCalc.setStyles(bobj.element, buttoninfo.buttonElement.downstyle);

   // Register event handler for mouse up

   // Event code from JavaScript, Flanagan, 5th Edition, pg. 422
   if (document.addEventListener) { // DOM Level 2 -- Firefox, et al
      document.addEventListener("mouseup", SocialCalc.ButtonMouseUp, true); // capture everywhere
   } else if (bobj.element.attachEvent) { // IE 5+
      bobj.element.setCapture();
      bobj.element.attachEvent("onmouseup", SocialCalc.ButtonMouseUp);
      bobj.element.attachEvent("onlosecapture", SocialCalc.ButtonMouseUp);
   }
   if (e.stopPropagation) e.stopPropagation(); // DOM Level 2
   else e.cancelBubble = true; // IE 5+
   if (e.preventDefault) e.preventDefault(); // DOM Level 2
   else e.returnValue = false; // IE 5+

   buttoninfo.horizontalScroll = viewportinfo.horizontalScroll;
   buttoninfo.verticalScroll = viewportinfo.verticalScroll;
   buttoninfo.clientX = e.clientX + buttoninfo.horizontalScroll; // get document-relative coordinates
   buttoninfo.clientY = e.clientY + buttoninfo.verticalScroll;

   if (bobj && bobj.functionobj && bobj.functionobj.MouseDown) bobj.functionobj.MouseDown(e, buttoninfo, bobj);

   if (bobj.repeatwait) { // if a repeat wait is set, then starting waiting for first repetition
      buttoninfo.timer = window.setTimeout(SocialCalc.ButtonRepeat, bobj.repeatwait);
   }

   return;
};

/**
 * @function ButtonMouseUp
 * @memberof SocialCalc
 * @description Handles mouse up events for button elements
 * @param {Event} event - The mouse up event
 */
SocialCalc.ButtonMouseUp = function(event) {
   const e = event || window.event;

   const buttoninfo = SocialCalc.ButtonInfo;
   const bobj = buttoninfo.buttonElement;

   if (buttoninfo.timer) { // if repeating, cancel it
      window.clearTimeout(buttoninfo.timer); // cancel timer
      buttoninfo.timer = null;
   }

   if (!buttoninfo.buttonDown) return; // already did this (e.g., in IE, releaseCapture fires losecapture)

   if (e.stopPropagation) e.stopPropagation(); // DOM Level 2
   else e.cancelBubble = true; // IE 5+
   if (e.preventDefault) e.preventDefault(); // DOM Level 2
   else e.returnValue = false; // IE 5+

   if (document.removeEventListener) { // DOM Level 2
      document.removeEventListener("mouseup", SocialCalc.ButtonMouseUp, true);
   } else if (document.detachEvent) { // IE
      bobj.element.detachEvent("onlosecapture", SocialCalc.ButtonMouseUp);
      bobj.element.detachEvent("onmouseup", SocialCalc.ButtonMouseUp);
      bobj.element.releaseCapture();
   }

   if (buttoninfo.buttonElement.downstyle) {
      if (buttoninfo.doingHover)
         SocialCalc.setStyles(bobj.element, buttoninfo.buttonElement.hoverstyle);
      else
         SocialCalc.setStyles(bobj.element, buttoninfo.buttonElement.normalstyle);
   }

   buttoninfo.buttonDown = false;

   if (bobj && bobj.functionobj && bobj.functionobj.MouseUp) bobj.functionobj.MouseUp(e, buttoninfo, bobj);
};
/**
 * @function ButtonRepeat
 * @memberof SocialCalc
 * @description Handles button repeat functionality for buttons with repeat intervals
 */
SocialCalc.ButtonRepeat = function() {
   const buttoninfo = SocialCalc.ButtonInfo;
   const bobj = buttoninfo.buttonElement;

   if (!bobj) return;

   if (bobj && bobj.functionobj && bobj.functionobj.Repeat) bobj.functionobj.Repeat(null, buttoninfo, bobj);

   buttoninfo.timer = window.setTimeout(SocialCalc.ButtonRepeat, bobj.repeatinterval || 100);
};

/**
 * @namespace MouseWheelFunctions
 * @description MouseWheel functions for SocialCalc
 */

/**
 * @object MouseWheelInfo
 * @memberof SocialCalc
 * @description There is only one of these -- no "new" is done.
 * The mousewheel only affects the one area the mouse pointer is over
 * The registeredElements array is used to identify items.
 */
SocialCalc.MouseWheelInfo = {
   /**
    * @description One item for each element to respond to the mousewheel, each an object with:
    *    .element, .functionobj
    * @type {Array<Object>}
    */
   registeredElements: []
};

/**
 * @function MouseWheelRegister
 * @memberof SocialCalc
 * @description Make element respond to mousewheel events
 * @param {HTMLElement} element - The element to register for mousewheel events
 * @param {Object} functionobj - Function object containing WheelMove handler
 * @throws {string} When browser is not supported
 */
SocialCalc.MouseWheelRegister = function(element, functionobj) {
   const mousewheelinfo = SocialCalc.MouseWheelInfo;

   mousewheelinfo.registeredElements.push({
      element: element, 
      functionobj: functionobj
   });

   if (element.addEventListener) { // DOM Level 2 -- Firefox, et al
      element.addEventListener("DOMMouseScroll", SocialCalc.ProcessMouseWheel, false);
      element.addEventListener("mousewheel", SocialCalc.ProcessMouseWheel, false); // Opera needs this
   } else if (element.attachEvent) { // IE 5+
      element.attachEvent("onmousewheel", SocialCalc.ProcessMouseWheel);
   } else { // don't handle this
      throw SocialCalc.Constants.s_BrowserNotSupported;
   }

   return;
};

/**
 * @function ProcessMouseWheel
 * @memberof SocialCalc
 * @description Processes mouse wheel events and calls appropriate handlers
 * @param {Event} e - The mouse wheel event
 */
SocialCalc.ProcessMouseWheel = function(e) {
   const event = e || window.event;
   let delta;

   if (SocialCalc.Keyboard.passThru) return; // ignore

   const mousewheelinfo = SocialCalc.MouseWheelInfo;
   let ele = event.target || event.srcElement; // source object is often within what we want
   let wobj;

   for (wobj = null; !wobj && ele; ele = ele.parentNode) { // go up tree looking for one of our elements
      wobj = SocialCalc.LookupElement(ele, mousewheelinfo.registeredElements);
   }
   if (!wobj) return; // not one of our elements

   if (event.wheelDelta) {
      delta = event.wheelDelta / 120;
   } else {
      delta = -event.detail / 3;
   }
   if (!delta) delta = 0;

   if (wobj.functionobj && wobj.functionobj.WheelMove) wobj.functionobj.WheelMove(event, delta, mousewheelinfo, wobj);

   if (event.preventDefault) event.preventDefault();
   event.returnValue = false;
};

/**
 * @namespace KeyboardFunctions
 * @description Keyboard functions for SocialCalc
 * 
 * For more information about keyboard handling, see: http://unixpapa.com/js/key.html
 */

/**
 * @object keyboardTables
 * @memberof SocialCalc
 * @description Keyboard mapping tables for different browsers and key combinations
 */
SocialCalc.keyboardTables = {
   /**
    * @member {Object} specialKeysCommon - Common special keys mapping
    */
   specialKeysCommon: {
      8: "[backspace]", 9: "[tab]", 13: "[enter]", 25: "[tab]", 27: "[esc]", 33: "[pgup]", 34: "[pgdn]",
      35: "[end]", 36: "[home]", 37: "[aleft]", 38: "[aup]", 39: "[aright]", 40: "[adown]", 45: "[ins]",
      46: "[del]", 113: "[f2]"
   },

   /**
    * @member {Object} specialKeysIE - Internet Explorer special keys mapping
    */
   specialKeysIE: {
      8: "[backspace]", 9: "[tab]", 13: "[enter]", 25: "[tab]", 27: "[esc]", 33: "[pgup]", 34: "[pgdn]",
      35: "[end]", 36: "[home]", 37: "[aleft]", 38: "[aup]", 39: "[aright]", 40: "[adown]", 45: "[ins]",
      46: "[del]", 113: "[f2]"
   },

   /**
    * @member {Object} controlKeysIE - Internet Explorer control key combinations
    */
   controlKeysIE: {
      67: "[ctrl-c]",
      83: "[ctrl-s]",
      86: "[ctrl-v]",
      88: "[ctrl-x]",
      90: "[ctrl-z]"
   },

   /**
    * @member {Object} specialKeysOpera - Opera browser special keys mapping
    */
   specialKeysOpera: {
      8: "[backspace]", 9: "[tab]", 13: "[enter]", 25: "[tab]", 27: "[esc]", 33: "[pgup]", 34: "[pgdn]",
      35: "[end]", 36: "[home]", 37: "[aleft]", 38: "[aup]", 39: "[aright]", 40: "[adown]",
      45: "[ins]", // issues with releases before 9.5 - same as "-" ("-" changed in 9.5)
      46: "[del]", // issues with releases before 9.5 - same as "." ("." changed in 9.5)
      113: "[f2]"
   },

   /**
    * @member {Object} controlKeysOpera - Opera control key combinations
    */
   controlKeysOpera: {
      67: "[ctrl-c]",
      83: "[ctrl-s]",
      86: "[ctrl-v]",
      88: "[ctrl-x]",
      90: "[ctrl-z]"
   },

   /**
    * @member {Object} specialKeysSafari - Safari browser special keys mapping
    */
   specialKeysSafari: {
      8: "[backspace]", 9: "[tab]", 13: "[enter]", 25: "[tab]", 27: "[esc]", 63232: "[aup]", 63233: "[adown]",
      63234: "[aleft]", 63235: "[aright]", 63272: "[del]", 63273: "[home]", 63275: "[end]", 63276: "[pgup]",
      63277: "[pgdn]", 63237: "[f2]"
   },

   /**
    * @member {Object} controlKeysSafari - Safari control key combinations
    */
   controlKeysSafari: {
      99: "[ctrl-c]",
      115: "[ctrl-s]",
      118: "[ctrl-v]",
      120: "[ctrl-x]",
      122: "[ctrl-z]"
   },

   /**
    * @member {Object} ignoreKeysSafari - Safari keys to ignore
    */
   ignoreKeysSafari: {
      63236: "[f1]", 63238: "[f3]", 63239: "[f4]", 63240: "[f5]", 63241: "[f6]", 63242: "[f7]",
      63243: "[f8]", 63244: "[f9]", 63245: "[f10]", 63246: "[f11]", 63247: "[f12]", 63289: "[numlock]"
   },

   /**
    * @member {Object} specialKeysFirefox - Firefox browser special keys mapping
    */
   specialKeysFirefox: {
      8: "[backspace]", 9: "[tab]", 13: "[enter]", 25: "[tab]", 27: "[esc]", 33: "[pgup]", 34: "[pgdn]",
      35: "[end]", 36: "[home]", 37: "[aleft]", 38: "[aup]", 39: "[aright]", 40: "[adown]", 45: "[ins]",
      46: "[del]", 113: "[f2]"
   },

   /**
    * @member {Object} controlKeysFirefox - Firefox control key combinations
    */
   controlKeysFirefox: {
      99: "[ctrl-c]",
      115: "[ctrl-s]",
      118: "[ctrl-v]",
      120: "[ctrl-x]",
      122: "[ctrl-z]"
   },

   /**
    * @member {Object} ignoreKeysFirefox - Firefox keys to ignore
    */
   ignoreKeysFirefox: {
      16: "[shift]", 17: "[ctrl]", 18: "[alt]", 20: "[capslock]", 19: "[pause]", 44: "[printscreen]",
      91: "[windows]", 92: "[windows]", 112: "[f1]", 114: "[f3]", 115: "[f4]", 116: "[f5]",
      117: "[f6]", 118: "[f7]", 119: "[f8]", 120: "[f9]", 121: "[f10]", 122: "[f11]", 123: "[f12]",
      144: "[numlock]", 145: "[scrolllock]", 224: "[cmd]"
   }
};

/**
 * @object Keyboard
 * @memberof SocialCalc
 * @description Keyboard handling state and configuration
 */
SocialCalc.Keyboard = {
   /**
    * @member {boolean} areListener - If true, we have been installed as a listener for keyboard events
    */
   areListener: false,
   
   /**
    * @member {Object|null} focusTable - The table editor object that gets keystrokes or null
    */
   focusTable: null,
   
   /**
    * @member {Object|string|null} passThru - If not null, control element with focus to pass keyboard events to (has blur method), or "true"
    */
   passThru: null,
   
   /**
    * @member {boolean} didProcessKey - Did SocialCalc.ProcessKey in keydown
    */
   didProcessKey: false,
   
   /**
    * @member {boolean} statusFromProcessKey - The status from the keydown SocialCalc.ProcessKey
    */
   statusFromProcessKey: false,
   
   /**
    * @member {boolean} repeatingKeyPress - Some browsers (Opera, Gecko Mac) repeat special keys as KeyPress not KeyDown
    */
   repeatingKeyPress: false,
   
   /**
    * @member {string} chForProcessKey - Remember so can do repeat in those cases
    */
   chForProcessKey: ""
};

/**
 * @function KeyboardSetFocus
 * @memberof SocialCalc
 * @description Sets keyboard focus to the specified editor
 * @param {Object} editor - The table editor to set focus to
 */
SocialCalc.KeyboardSetFocus = function(editor) {
   SocialCalc.Keyboard.focusTable = editor;

   if (!SocialCalc.Keyboard.areListener) {
      document.onkeydown = SocialCalc.ProcessKeyDown;
      document.onkeypress = SocialCalc.ProcessKeyPress;
      SocialCalc.Keyboard.areListener = true;
   }
   if (SocialCalc.Keyboard.passThru) {
      if (SocialCalc.Keyboard.passThru.blur) {
         SocialCalc.Keyboard.passThru.blur();
      }
      SocialCalc.Keyboard.passThru = null;
   }
   window.focus();
};

/**
 * @function KeyboardFocus
 * @memberof SocialCalc
 * @description Sets keyboard focus to the window and clears passThru
 */
SocialCalc.KeyboardFocus = function() {
   SocialCalc.Keyboard.passThru = null;
   window.focus();
};

/**
 * @function ProcessKeyDown
 * @memberof SocialCalc
 * @description Processes keydown events and handles special key combinations
 * @param {Event} e - The keydown event
 * @returns {boolean} Status indicating whether to prevent default behavior
 */
SocialCalc.ProcessKeyDown = function(e) {
   const kt = SocialCalc.keyboardTables;
   kt.didProcessKey = false; // always start false
   kt.statusFromProcessKey = false;
   kt.repeatingKeyPress = false;

   let ch = "";
   let status = true;

   if (SocialCalc.Keyboard.passThru) return; // ignore

   e = e || window.event;

   // IE and Safari 3.1+ won't fire keyPress, so check for special keys here.
   if (e.which === undefined || (typeof e.keyIdentifier === "string")) {
      ch = kt.specialKeysCommon[e.keyCode];
      if (!ch) {
         if (e.ctrlKey) {
            ch = kt.controlKeysIE[e.keyCode];
         }
         if (!ch)
            return true;
      }
      status = SocialCalc.ProcessKey(ch, e);

      if (!status) {
         if (e.preventDefault) e.preventDefault();
         e.returnValue = false;
      }
   } else { 
      ch = kt.specialKeysCommon[e.keyCode];
      if (!ch) {
//         return true;
         if (e.ctrlKey || e.metaKey) {
            ch = kt.controlKeysIE[e.keyCode]; // this works here
         }
         if (!ch)
            return true;
      }

      status = SocialCalc.ProcessKey(ch, e); // process the key
      kt.didProcessKey = true; // remember what happened
      kt.statusFromProcessKey = status;
      kt.chForProcessKey = ch;
   }

   return status;
};
/**
 * @function ProcessKeyPress
 * @memberof SocialCalc
 * @description Processes keypress events with cross-browser compatibility
 * @param {Event} e - The keypress event
 * @returns {boolean} Status indicating whether to prevent default behavior
 */
SocialCalc.ProcessKeyPress = function(e) {
   const kt = SocialCalc.keyboardTables;

   let ch = "";

   e = e || window.event;

   if (SocialCalc.Keyboard.passThru) return; // ignore
   if (kt.didProcessKey) { // already processed this key
      if (kt.repeatingKeyPress) {
         return SocialCalc.ProcessKey(kt.chForProcessKey, e); // process the same key as on KeyDown
      } else {
         kt.repeatingKeyPress = true; // see if get another KeyPress before KeyDown
         return kt.statusFromProcessKey; // do what it said to do
      }
   }

   if (e.which === undefined) { // IE
      // Note: Esc and Enter will come through here, too, if not stopped at KeyDown
      ch = String.fromCharCode(e.keyCode); // convert to a character (special chars handled at ev1)
   } else { // not IE
      if (!e.which)
         return false; // ignore - special key
      if (e.charCode === undefined) { // Opera
         if (e.which !== 0) { // character
            if (e.which < 32 || e.which === 144) { // special char (144 is numlock)
               ch = kt.specialKeysOpera[e.which];
               if (ch) {
                  return true;
               }
            } else {
               if (e.ctrlKey) {
                  ch = kt.controlKeysOpera[e.keyCode];
               } else {
                  ch = String.fromCharCode(e.which);
               }
            }
         } else { // special char
            return true;
         }
      } else if (e.keyCode === 0 && e.charCode === 0) { // OLPC Fn key or something
         return; // ignore
      } else if (e.keyCode === e.charCode) { // Safari
         ch = kt.specialKeysSafari[e.keyCode];
         if (!ch) {
            if (kt.ignoreKeysSafari[e.keyCode]) // pass this through
               return true;
            if (e.metaKey) {
               ch = kt.controlKeysSafari[e.keyCode];
            } else {
               ch = String.fromCharCode(e.which);
            }
         }
      } else { // Firefox
         if (kt.specialKeysFirefox[e.keyCode]) {
            return true;
         }
         ch = String.fromCharCode(e.which);
         if (e.ctrlKey || e.metaKey) {
            ch = kt.controlKeysFirefox[e.which];
         }
      }
   }

   const status = SocialCalc.ProcessKey(ch, e);

   if (!status) {
      if (e.preventDefault) e.preventDefault();
      e.returnValue = false;
   }

   return status;
};

/* 
 * @deprecated
 * OLD ProcessKeyDown and ProcessKeyPress -- replaced for handling newer browsers, including Safari 3.1 and Opera 9.5
 *
 * @function ProcessKeyDown (OLD VERSION - DEPRECATED)
 * @memberof SocialCalc
 * @description Legacy ProcessKeyDown implementation
 * @param {Event} e - The keydown event
 * @returns {boolean} Status indicating whether to prevent default behavior

SocialCalc.ProcessKeyDown = function(e) {
   const kt = SocialCalc.keyboardTables;

   let ch = "";
   let status = true;

   if (SocialCalc.Keyboard.passThru) return; // ignore

   e = e || window.event;

   if (e.which === undefined) { // IE
      ch = kt.specialKeysIE[e.keyCode];
      if (!ch) {
         if (e.ctrlKey) {
            ch = kt.controlKeysIE[e.keyCode];
         }
         if (!ch)
            return true;
      }

      status = SocialCalc.ProcessKey(ch, e);

      if (!status) {
         if (e.preventDefault) e.preventDefault();
         e.returnValue = false;
      }
   } else { // don't do anything for other browsers - wait for keyPress
      ; // special key repeats are done as keypress in those browsers
   }

   return status;
};

 * @function ProcessKeyPress (OLD VERSION - DEPRECATED)
 * @memberof SocialCalc  
 * @description Legacy ProcessKeyPress implementation
 * @param {Event} e - The keypress event
 * @returns {boolean} Status indicating whether to prevent default behavior

SocialCalc.ProcessKeyPress = function(e) {
   const kt = SocialCalc.keyboardTables;

   let ch = "";

   if (SocialCalc.Keyboard.passThru) return; // ignore

   e = e || window.event;

   if (e.which === undefined) { // IE
      // Note: Esc and Enter will come through here, too, if not stopped at KeyDown
      ch = String.fromCharCode(e.keyCode); // convert to a character (special chars handled at ev1)
   } else { // not IE
      if (e.charCode === undefined) { // Opera
         if (e.which !== 0) { // character
            if (e.which < 32) { // special char
               ch = kt.specialKeysOpera[e.keyCode];
               if (!ch)
                  return true;
            } else {
               if (e.ctrlKey) {
                  ch = kt.controlKeysOpera[e.keyCode];
               } else {
                  ch = String.fromCharCode(e.which);
               }
            }
         } else { // special char
            ch = kt.specialKeysOpera[e.keyCode];
            if (!ch)
               return true;
         }
      } else if (e.keyCode === 0 && e.charCode === 0) { // OLPC Fn key or something
         return; // ignore
      } else if (e.keyCode === e.charCode) { // Safari
         ch = kt.specialKeysSafari[e.keyCode];
         if (!ch) {
            if (kt.ignoreKeysSafari[e.keyCode]) // pass this through
               return true;
            if (e.metaKey) {
               ch = kt.controlKeysSafari[e.keyCode];
            } else {
               ch = String.fromCharCode(e.which);
            }
         }
      } else { // Firefox
         ch = kt.specialKeysFirefox[e.keyCode];
         if (!ch) {
            if (kt.ignoreKeysFirefox[e.keyCode]) // pass this through
               return true;
            if (e.which) { // normal char
               if (e.ctrlKey || e.metaKey) {
                  ch = kt.controlKeysFirefox[e.which];
               } else {
                  ch = String.fromCharCode(e.which);
               }
            } else { // usually a special char
               return true; // old Firefox gives extra, empty keyPress for "/" - ignore
            }
         }
      }
   }

   const status = SocialCalc.ProcessKey(ch, e);

   if (!status) {
      if (e.preventDefault) e.preventDefault();
      e.returnValue = false;
   }

   return status;
};
*/

/**
 * @function ProcessKey
 * @memberof SocialCalc
 * @description Take a key representation as a character string and dispatch to appropriate routine
 * @param {string} ch - The character representation of the key
 * @param {Event} e - The keyboard event
 * @returns {boolean} Status from the key processing
 */
SocialCalc.ProcessKey = function(ch, e) {
   const ft = SocialCalc.Keyboard.focusTable;

   if (!ft) return true; // we're not handling it -- let browser do default

   return ft.EditorProcessKey(ch, e);
};
