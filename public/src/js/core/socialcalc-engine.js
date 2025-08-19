/**
 * @fileoverview The main SocialCalc code module of the SocialCalc package
 * @version 2.0.0
 * @author Dan Bricklin of Software Garden, Inc., for Socialtext, Inc.
 * @copyright (c) Copyright 2010 Socialtext, Inc. All Rights Reserved.
 * @license Artistic License 2.0 - http://socialcalc.org/licenses/al-20/
 * @description
 * This is the beginning of a library of routines for displaying and editing spreadsheet
 * data in a browser. The HTML that includes this does not need to have anything
 * specific to the spreadsheet or editor already present -- everything is dynamically
 * added to the DOM by this code, including the rendered sheet and any editing controls.
 *
 * The library has several parts:
 * - Main SocialCalc code module (this file)
 * - Table Editor module
 * - Formula module
 * - Format Number module
 *
 * Design Goals:
 * - Support multiple active spreadsheets on one page
 * - Dynamic DOM manipulation
 * - Efficient rendering for large spreadsheets
 * - Cross-browser compatibility
 *
 * Browser Compatibility:
 * - Windows Firefox (2 and 3)
 * - Internet Explorer (6 and 7)
 * - Opera (9.23+)
 * - Mac Safari (3.1)
 * - Mac Firefox (2.0.0.6+)
 */

/**
 * @namespace SocialCalc
 * @description Main SocialCalc namespace containing all classes, functions, and data
 */
var SocialCalc;
if (!SocialCalc) SocialCalc = {};

/**
 * @namespace SocialCalc.Callbacks
 * @description Callback functions for various SocialCalc operations
 */
SocialCalc.Callbacks = {
    /**
     * @description Function to expand wiki text - should be set if you want wikitext expansion
     * @type {Function|null}
     * @param {string} displayvalue - The value to be displayed
     * @param {SocialCalc.Sheet} sheetobj - The sheet object
     * @param {string} linkstyle - The link style
     * @param {string} valueformat - Format like text-wiki followed by optional sub-formats
     * @returns {string} Expanded wiki text
     */
    expand_wiki: null,

    /**
     * @description The old function to expand wiki text - may be replaced
     * @type {Function}
     * @param {string} displayvalue - The value to be displayed
     * @param {SocialCalc.Sheet} sheetobj - The sheet object
     * @param {string} linkstyle - The link style
     * @returns {string} Expanded markup text
     */
    expand_markup: function (displayvalue, sheetobj, linkstyle) {
        return SocialCalc.default_expand_markup(displayvalue, sheetobj, linkstyle);
    },

    /**
     * @description Used to create the href for a link to another "page"
     * @type {Function|null}
     * @param {string} pagename - Name of the page to link to
     * @param {string} workspacename - Name of the workspace
     * @param {string} linkstyle - Style of the link
     * @param {string} valueformat - Format of the value
     * @returns {string} The href string for the page link
     */
    MakePageLink: null,

    /**
     * @description Used to make different variations of sheetnames use the same cache slot
     * @type {Function|null}
     * @param {string} sheetname - The sheet name to normalize
     * @returns {string} Normalized sheet name
     * @default Uses lowercase normalization
     */
    NormalizeSheetName: null
};

/**
 * @class SocialCalc.Cell
 * @description Represents a single cell in a spreadsheet
 * @param {string} coord - The column/row as a string, e.g., "A1"
 * 
 * @property {string} coord - The column/row coordinate
 * @property {string|number} datavalue - Value for computation and formatting
 * @property {string|null} datatype - v=numeric, t=text, f=formula, c=constant
 * @property {string} formula - Formula without leading "=" or constant value
 * @property {string} valuetype - Main type (b=blank, n=numeric, t=text, e=error) + subtypes
 * @property {string} [displayvalue] - Rendered version with formatting applied
 * @property {Object} [parseinfo] - Cached parsed version of formula
 * @property {number} [bt] - Top border definition number
 * @property {number} [br] - Right border definition number
 * @property {number} [bb] - Bottom border definition number
 * @property {number} [bl] - Left border definition number
 * @property {number} [layout] - Layout definition number
 * @property {number} [font] - Font definition number
 * @property {number} [color] - Text color definition number
 * @property {number} [bgcolor] - Background color definition number
 * @property {number} [cellformat] - Cell format definition number
 * @property {number} [nontextvalueformat] - Custom format for non-text values
 * @property {number} [textvalueformat] - Custom format for text values
 * @property {number} [colspan] - Number of cells to span horizontally
 * @property {number} [rowspan] - Number of cells to span vertically
 * @property {string} [cssc] - Custom CSS classname
 * @property {string} [csss] - Custom CSS style definition
 * @property {string} [mod] - Modification allowed flag ("y" if present)
 * @property {string} [comment] - Cell comment string
 */
SocialCalc.Cell = function (coord) {
    this.coord = coord;
    this.datavalue = "";
    this.datatype = null;
    this.formula = "";
    this.valuetype = "b";
};

/**
 * @constant {Object} SocialCalc.CellProperties
 * @description Defines the types of cell properties
 * Type 1: Base properties
 * Type 2: Attribute properties  
 * Type 3: Special properties (e.g., displaystring, parseinfo)
 */
SocialCalc.CellProperties = {
    // Type 1: Base properties
    coord: 1,
    datavalue: 1,
    datatype: 1,
    formula: 1,
    valuetype: 1,
    errors: 1,
    comment: 1,

    // Type 2: Attribute properties
    bt: 2,
    br: 2,
    bb: 2,
    bl: 2,
    layout: 2,
    font: 2,
    color: 2,
    bgcolor: 2,
    cellformat: 2,
    nontextvalueformat: 2,
    textvalueformat: 2,
    colspan: 2,
    rowspan: 2,
    cssc: 2,
    csss: 2,
    mod: 2,

    // Type 3: Special properties
    displaystring: 3, // used to cache rendered HTML of cell contents
    parseinfo: 3, // used to cache parsed formulas
    hcolspan: 3, // spans taking hidden cols/rows into account (!!! NOT YET !!!)
    hrowspan: 3
};

/**
 * @constant {Object} SocialCalc.CellPropertiesTable
 * @description Maps cell properties to their corresponding table types
 */
SocialCalc.CellPropertiesTable = {
    bt: "borderstyle",
    br: "borderstyle",
    bb: "borderstyle",
    bl: "borderstyle",
    layout: "layout",
    font: "font",
    color: "color",
    bgcolor: "color",
    cellformat: "cellformat",
    nontextvalueformat: "valueformat",
    textvalueformat: "valueformat"
};

/**
 * @class SocialCalc.Sheet
 * @description Represents a complete spreadsheet with all its data and properties
 * 
 * @property {Function|null} statuscallback - Callback function for status updates
 * @property {*} statuscallbackparams - Parameters passed to status callback
 * @property {Object} cells - Cell objects indexed by coordinate
 * @property {Object} attribs - Sheet attributes
 * @property {Object} rowattribs - Row-specific attributes  
 * @property {Object} colattribs - Column-specific attributes
 * @property {Object} names - Named ranges and definitions
 * @property {Array} layouts - Layout definitions
 * @property {Object} layouthash - Layout hash for quick lookup
 * @property {Array} fonts - Font definitions
 * @property {Object} fonthash - Font hash for quick lookup
 * @property {Array} colors - Color definitions
 * @property {Object} colorhash - Color hash for quick lookup
 * @property {Array} borderstyles - Border style definitions
 * @property {Object} borderstylehash - Border style hash for quick lookup
 * @property {Array} cellformats - Cell format definitions
 * @property {Object} cellformathash - Cell format hash for quick lookup
 * @property {Array} valueformats - Value format definitions
 * @property {Object} valueformathash - Value format hash for quick lookup
 * @property {string} copiedfrom - Range copied from for clipboard operations
 * @property {SocialCalc.UndoStack} changes - Undo/redo stack
 * @property {boolean} renderneeded - Flag indicating if render is needed
 * @property {boolean} changedrendervalues - Flag for span/font changes
 * @property {boolean} recalcchangedavalue - Flag for recalc value changes
 */
SocialCalc.Sheet = function () {
    SocialCalc.ResetSheet(this);

    /**
     * @description Status callback function called during recalc and commands
     * @type {Function|null}
     * @param {Object} data - Current data (recalcdata during recalc, SheetCommandInfo during commands)
     * @param {string} status - Status type (calcorder, calcstep, cmdstart, etc.)
     * @param {*} arg - Status-specific argument
     * @param {*} params - Parameters from statuscallbackparams
     * 
     * Recalc status values:
     * - calcorder: {coord, total, count}
     * - calccheckdone: calclist length
     * - calcstep: {coord, total, count}
     * - calcloading: {sheetname}
     * - calcserverfunc: {funcname, coord, total, count}
     * - calcfinished: time in milliseconds
     * 
     * Command status values:
     * - cmdstart: cmdstr
     * - cmdend: (no arg)
     */
    this.statuscallback = null;
    this.statuscallbackparams = null;
};

/**
 * @function SocialCalc.ResetSheet
 * @description Resets (and/or initializes) sheet data values
 * @param {SocialCalc.Sheet} sheet - The sheet object to reset
 * @param {boolean} [reload] - Whether this is a reload operation
 */
SocialCalc.ResetSheet = function (sheet, reload) {
    // Initialize cell storage
    sheet.cells = {}; // at least one for each non-blank cell: coord: cell-object

    // Sheet-level attributes
    sheet.attribs = {
        lastcol: 1,
        lastrow: 1,
        defaultlayout: 0
    };

    // Row-specific attributes
    sheet.rowattribs = {
        hide: {}, // access by row number
        height: {}
    };

    // Column-specific attributes  
    sheet.colattribs = {
        width: {}, // access by col name
        hide: {}
    };

    // Named ranges: {desc: "optional description", definition: "B5, A1:B7, or =formula"}
    sheet.names = {};

    // Style and formatting arrays with corresponding hash tables
    sheet.layouts = [];
    sheet.layouthash = {};
    sheet.fonts = [];
    sheet.fonthash = {};
    sheet.colors = [];
    sheet.colorhash = {};
    sheet.borderstyles = [];
    sheet.borderstylehash = {};
    sheet.cellformats = [];
    sheet.cellformathash = {};
    sheet.valueformats = [];
    sheet.valueformathash = {};

    // Clipboard information
    sheet.copiedfrom = ""; // if a range, then this was loaded from a saved range as clipboard content

    // Undo/redo system
    sheet.changes = new SocialCalc.UndoStack();

    // Rendering flags
    sheet.renderneeded = false;
    sheet.changedrendervalues = true; // if true, spans and/or fonts have changed
    sheet.recalcchangedavalue = false; // true if a recalc resulted in a change to a cell's calculated value
};
/**
 * @description Prototype method to reset the sheet using the static ResetSheet function
 * @memberof SocialCalc.Sheet
 * @returns {void}
 */
SocialCalc.Sheet.prototype.ResetSheet = function () {
    SocialCalc.ResetSheet(this);
};

/**
 * @description Adds a new cell to the sheet's cell collection
 * @memberof SocialCalc.Sheet
 * @param {SocialCalc.Cell} newcell - The cell object to add
 * @returns {SocialCalc.Cell} The added cell object
 */
SocialCalc.Sheet.prototype.AddCell = function (newcell) {
    return this.cells[newcell.coord] = newcell;
};

/**
 * @description Gets an existing cell or creates a new one if it doesn't exist
 * @memberof SocialCalc.Sheet
 * @param {string} coord - The cell coordinate (e.g., "A1")
 * @returns {SocialCalc.Cell} The existing or newly created cell object
 */
SocialCalc.Sheet.prototype.GetAssuredCell = function (coord) {
    return this.cells[coord] || this.AddCell(new SocialCalc.Cell(coord));
};

/**
 * @description Parses a saved sheet string and populates this sheet with the data
 * @memberof SocialCalc.Sheet
 * @param {string} savedsheet - The saved sheet data string to parse
 * @returns {void}
 */
SocialCalc.Sheet.prototype.ParseSheetSave = function (savedsheet) {
    SocialCalc.ParseSheetSave(savedsheet, this);
};

/**
 * @description Creates a cell from string parts during parsing
 * @memberof SocialCalc.Sheet
 * @param {SocialCalc.Cell} cell - The cell object to populate
 * @param {Array<string>} parts - Array of string parts from parsing
 * @param {number} j - Starting index in the parts array
 * @returns {*} Result from the static CellFromStringParts function
 */
SocialCalc.Sheet.prototype.CellFromStringParts = function (cell, parts, j) {
    return SocialCalc.CellFromStringParts(this, cell, parts, j);
};

/**
 * @description Creates a save string representation of the sheet or a range
 * @memberof SocialCalc.Sheet
 * @param {string} [range] - Optional range to save (e.g., "A1:B10")
 * @param {boolean} [canonicalize] - Whether to canonicalize the output
 * @returns {string} The sheet save string
 */
SocialCalc.Sheet.prototype.CreateSheetSave = function (range, canonicalize) {
    return SocialCalc.CreateSheetSave(this, range, canonicalize);
};

/**
 * @description Converts a cell to its string representation
 * @memberof SocialCalc.Sheet
 * @param {SocialCalc.Cell} cell - The cell to convert
 * @returns {string} String representation of the cell
 */
SocialCalc.Sheet.prototype.CellToString = function (cell) {
    return SocialCalc.CellToString(this, cell);
};

/**
 * @description Canonicalizes the sheet data
 * @memberof SocialCalc.Sheet
 * @param {boolean} [full] - Whether to perform full canonicalization
 * @returns {*} Result from the static CanonicalizeSheet function
 */
SocialCalc.Sheet.prototype.CanonicalizeSheet = function (full) {
    return SocialCalc.CanonicalizeSheet(this, full);
};

/**
 * @description Encodes cell attributes for saving
 * @memberof SocialCalc.Sheet
 * @param {string} coord - The cell coordinate
 * @returns {string} Encoded cell attributes string
 */
SocialCalc.Sheet.prototype.EncodeCellAttributes = function (coord) {
    return SocialCalc.EncodeCellAttributes(this, coord);
};

/**
 * @description Encodes sheet-level attributes for saving
 * @memberof SocialCalc.Sheet
 * @returns {string} Encoded sheet attributes string
 */
SocialCalc.Sheet.prototype.EncodeSheetAttributes = function () {
    return SocialCalc.EncodeSheetAttributes(this);
};

/**
 * @description Decodes and applies cell attributes from a save string
 * @memberof SocialCalc.Sheet
 * @param {string} coord - The cell coordinate
 * @param {string} attribs - The encoded attributes string
 * @param {string} [range] - Optional range context
 * @returns {*} Result from the static DecodeCellAttributes function
 */
SocialCalc.Sheet.prototype.DecodeCellAttributes = function (coord, attribs, range) {
    return SocialCalc.DecodeCellAttributes(this, coord, attribs, range);
};

/**
 * @description Decodes and applies sheet-level attributes from a save string
 * @memberof SocialCalc.Sheet
 * @param {string} attribs - The encoded attributes string
 * @returns {*} Result from the static DecodeSheetAttributes function
 */
SocialCalc.Sheet.prototype.DecodeSheetAttributes = function (attribs) {
    return SocialCalc.DecodeSheetAttributes(this, attribs);
};

/**
 * @description Schedules sheet commands for execution
 * @memberof SocialCalc.Sheet
 * @param {string} cmd - The command string to execute
 * @param {boolean} [saveundo] - Whether to save undo information
 * @param {boolean} [isRemote] - Whether this is a remote command
 * @returns {*} Result from the static ScheduleSheetCommands function
 */
SocialCalc.Sheet.prototype.ScheduleSheetCommands = function (cmd, saveundo, isRemote) {
    return SocialCalc.ScheduleSheetCommands(this, cmd, saveundo, isRemote);
};

/**
 * @description Performs an undo operation on the sheet
 * @memberof SocialCalc.Sheet
 * @returns {*} Result from the static SheetUndo function
 */
SocialCalc.Sheet.prototype.SheetUndo = function () {
    return SocialCalc.SheetUndo(this);
};

/**
 * @description Performs a redo operation on the sheet
 * @memberof SocialCalc.Sheet
 * @returns {*} Result from the static SheetRedo function
 */
SocialCalc.Sheet.prototype.SheetRedo = function () {
    return SocialCalc.SheetRedo(this);
};

/**
 * @description Creates an audit string for the sheet
 * @memberof SocialCalc.Sheet
 * @returns {string} The audit string
 */
SocialCalc.Sheet.prototype.CreateAuditString = function () {
    return SocialCalc.CreateAuditString(this);
};

/**
 * @description Gets the style number for a given attribute type and style
 * @memberof SocialCalc.Sheet
 * @param {string} atype - The attribute type (font, color, etc.)
 * @param {string} style - The style definition
 * @returns {number} The style number/index
 */
SocialCalc.Sheet.prototype.GetStyleNum = function (atype, style) {
    return SocialCalc.GetStyleNum(this, atype, style);
};

/**
 * @description Gets the style string for a given attribute type and number
 * @memberof SocialCalc.Sheet
 * @param {string} atype - The attribute type (font, color, etc.)
 * @param {number} num - The style number/index
 * @returns {string} The style definition string
 */
SocialCalc.Sheet.prototype.GetStyleString = function (atype, num) {
    return SocialCalc.GetStyleString(this, atype, num);
};

/**
 * @description Recalculates all formulas in the sheet
 * @memberof SocialCalc.Sheet
 * @returns {*} Result from the static RecalcSheet function
 */
SocialCalc.Sheet.prototype.RecalcSheet = function () {
    return SocialCalc.RecalcSheet(this);
};

/**
 * @fileoverview Sheet Save Format Documentation
 * 
 * Format: linetype:param1:param2:...
 * 
 * Supported line types:
 * 
 * @example
 * version:versionname - Version of format (currently 1.4)
 * 
 * cell:coord:type:value... - Cell data with types:
 *   v:value - Numeric value
 *   t:value - Text/wiki-text (encoded)
 *   vt:fulltype:value - Value with type/subtype
 *   vtf:fulltype:value:formulatext - Formula with value/type
 *   vtc:fulltype:value:valuetext - Formatted constant
 *   vf:fvalue:formulatext - Formula (obsolete, pre v1.1)
 *   e:errortext - Error text
 *   b:top#:right#:bottom#:left# - Border definitions
 *   l:layout# - Layout definition number
 *   f:font# - Font definition number
 *   c:color# - Text color number
 *   bg:color# - Background color number
 *   cf:format# - Cell format number
 *   tvf:valueformat# - Text value format
 *   ntvf:valueformat# - Non-text value format
 *   colspan:numcols - Column span for merged cells
 *   rowspan:numrows - Row span for merged cells
 *   cssc:classname - CSS class name
 *   csss:styletext - CSS style (encoded)
 *   mod:allow - Modification flag ("y" to allow)
 *   comment:value - Cell comment (encoded)
 * 
 * col: - Column attributes
 *   w:widthval - Width (number, "auto", percentage, or blank)
 *   hide: - Hide flag (yes/no)
 * 
 * row: - Row attributes  
 *   hide - Hide flag (yes/no)
 * 
 * sheet: - Sheet-level attributes
 *   c:lastcol, r:lastrow - Last column/row numbers
 *   w:defaultcolwidth - Default column width
 *   layout:layout# - Default layout number
 *   font:font# - Default font number
 *   color:color#, bgcolor:color# - Default colors
 *   circularreferencecell:coord - Circular reference cell
 *   recalc:value - Auto-recalc setting (on/off)
 *   needsrecalc:value - Needs recalc flag (yes/no)
 * 
 * name:name:description:value - Named range definition
 * font:fontnum:value - Font definition
 * color:colornum:rgbvalue - Color definition  
 * border:bordernum:value - Border definition
 * layout:layoutnum:value - Layout definition
 * cellformat:cformatnum:value - Cell format definition
 * valueformat:vformatnum:value - Value format definition
 * copiedfrom:upperleft:bottomright - Clipboard source range
 */

/**
 * @function SocialCalc.ParseSheetSave
 * @description Parses a saved sheet string and populates the sheet object with data
 * @param {string} savedsheet - The saved sheet data string containing line-separated data
 * @param {SocialCalc.Sheet} sheetobj - The sheet object to populate
 * @throws {Error} Throws error for unknown line types, column types, or row types
 * 
 * @example
 * // Basic usage
 * const sheet = new SocialCalc.Sheet();
 * const saveData = "version:1.4\ncell:A1:v:42\nsheet:c:10:r:20";
 * SocialCalc.ParseSheetSave(saveData, sheet);
 */
SocialCalc.ParseSheetSave = function (savedsheet, sheetobj) {
    const lines = savedsheet.split(/\r\n|\n/);
    let parts = [];
    let line, i, j, t, v, coord, cell, attribs, name;
    const scc = SocialCalc.Constants;

    for (i = 0; i < lines.length; i++) {
        line = lines[i];
        parts = line.split(":");

        switch (parts[0]) {
            case "cell":
                cell = sheetobj.GetAssuredCell(parts[1]);
                j = 2;
                sheetobj.CellFromStringParts(cell, parts, j);
                break;

            case "col":
                coord = parts[1];
                j = 2;
                while (t = parts[j++]) {
                    switch (t) {
                        case "w":
                            sheetobj.colattribs.width[coord] = parts[j++]; // must be text - could be auto or %, etc.
                            break;
                        case "hide":
                            sheetobj.colattribs.hide[coord] = parts[j++];
                            break;
                        default:
                            throw `${scc.s_pssUnknownColType} '${t}'`;
                    }
                }
                break;

            case "row":
                coord = parts[1] - 0;
                j = 2;
                while (t = parts[j++]) {
                    switch (t) {
                        case "h":
                            sheetobj.rowattribs.height[coord] = parts[j++] - 0;
                            break;
                        case "hide":
                            sheetobj.rowattribs.hide[coord] = parts[j++];
                            break;
                        default:
                            throw `${scc.s_pssUnknownRowType} '${t}'`;
                    }
                }
                break;

            case "sheet":
                attribs = sheetobj.attribs;
                j = 1;
                while (t = parts[j++]) {
                    switch (t) {
                        case "c":
                            attribs.lastcol = parts[j++] - 0;
                            break;
                        case "r":
                            attribs.lastrow = parts[j++] - 0;
                            break;
                        case "w":
                            attribs.defaultcolwidth = parts[j++] + "";
                            break;
                        case "h":
                            attribs.defaultrowheight = parts[j++] - 0;
                            break;
                        case "tf":
                            attribs.defaulttextformat = parts[j++] - 0;
                            break;
                        case "ntf":
                            attribs.defaultnontextformat = parts[j++] - 0;
                            break;
                        case "layout":
                            attribs.defaultlayout = parts[j++] - 0;
                            break;
                        case "font":
                            attribs.defaultfont = parts[j++] - 0;
                            break;
                        case "tvf":
                            attribs.defaulttextvalueformat = parts[j++] - 0;
                            break;
                        case "ntvf":
                            attribs.defaultnontextvalueformat = parts[j++] - 0;
                            break;
                        case "color":
                            attribs.defaultcolor = parts[j++] - 0;
                            break;
                        case "bgcolor":
                            attribs.defaultbgcolor = parts[j++] - 0;
                            break;
                        case "circularreferencecell":
                            attribs.circularreferencecell = parts[j++];
                            break;
                        case "recalc":
                            attribs.recalc = parts[j++];
                            break;
                        case "needsrecalc":
                            attribs.needsrecalc = parts[j++];
                            break;
                        default:
                            j += 1;
                            break;
                    }
                }
                break;

            case "name":
                name = SocialCalc.decodeFromSave(parts[1]).toUpperCase();
                sheetobj.names[name] = { desc: SocialCalc.decodeFromSave(parts[2]) };
                sheetobj.names[name].definition = SocialCalc.decodeFromSave(parts[3]);
                break;

            case "layout":
                parts = lines[i].match(/^layout\:(\d+)\:(.+)$/); // layouts can have ":" in them
                sheetobj.layouts[parts[1] - 0] = parts[2];
                sheetobj.layouthash[parts[2]] = parts[1] - 0;
                break;

            case "font":
                sheetobj.fonts[parts[1] - 0] = parts[2];
                sheetobj.fonthash[parts[2]] = parts[1] - 0;
                break;

            case "color":
                sheetobj.colors[parts[1] - 0] = parts[2];
                sheetobj.colorhash[parts[2]] = parts[1] - 0;
                break;

            case "border":
                sheetobj.borderstyles[parts[1] - 0] = parts[2];
                sheetobj.borderstylehash[parts[2]] = parts[1] - 0;
                break;

            case "cellformat":
                v = SocialCalc.decodeFromSave(parts[2]);
                sheetobj.cellformats[parts[1] - 0] = v;
                sheetobj.cellformathash[v] = parts[1] - 0;
                break;

            case "valueformat":
                v = SocialCalc.decodeFromSave(parts[2]);
                sheetobj.valueformats[parts[1] - 0] = v;
                sheetobj.valueformathash[v] = parts[1] - 0;
                break;

            case "version":
                break;

            case "copiedfrom":
                sheetobj.copiedfrom = `${parts[1]}:${parts[2]}`;
                break;

            case "clipboardrange": // in save versions up to 1.3. Ignored.
            case "clipboard":
                break;

            case "":
                break;

            default:
                alert(`${scc.s_pssUnknownLineType} '${parts[0]}'`);
                throw `${scc.s_pssUnknownLineType} '${parts[0]}'`;
        }
        parts = null;
    }
};
/**
 * @function SocialCalc.CellFromStringParts
 * @description Takes string that has been split by ":" in parts, starting at item j,
 * and fills in cell assuming save format
 * @param {SocialCalc.Sheet} sheet - The sheet object containing the cell
 * @param {SocialCalc.Cell} cell - The cell object to populate with data
 * @param {Array<string>} parts - Array of string parts from parsing, split by ":"
 * @param {number} j - Starting index in the parts array
 * @throws {Error} Throws error for unknown cell types
 * 
 * @example
 * // Parse cell data from save format
 * const parts = ["cell", "A1", "v", "42", "f", "1"];
 * const cell = new SocialCalc.Cell("A1");
 * SocialCalc.CellFromStringParts(sheet, cell, parts, 2);
 */
SocialCalc.CellFromStringParts = function (sheet, cell, parts, j) {
    let t, v;

    while (t = parts[j++]) {
        switch (t) {
            case "v":
                cell.datavalue = SocialCalc.decodeFromSave(parts[j++]) - 0;
                cell.datatype = "v";
                cell.valuetype = "n";
                break;

            case "t":
                cell.datavalue = SocialCalc.decodeFromSave(parts[j++]);
                cell.datatype = "t";
                cell.valuetype = SocialCalc.Constants.textdatadefaulttype;
                break;

            case "vt":
                v = parts[j++];
                cell.valuetype = v;
                if (v.charAt(0) === "n") {
                    cell.datatype = "v";
                    cell.datavalue = SocialCalc.decodeFromSave(parts[j++]) - 0;
                } else {
                    cell.datatype = "t";
                    cell.datavalue = SocialCalc.decodeFromSave(parts[j++]);
                }
                break;

            case "vtf":
                v = parts[j++];
                cell.valuetype = v;
                if (v.charAt(0) === "n") {
                    cell.datavalue = SocialCalc.decodeFromSave(parts[j++]) - 0;
                } else {
                    cell.datavalue = SocialCalc.decodeFromSave(parts[j++]);
                }
                cell.formula = SocialCalc.decodeFromSave(parts[j++]);
                cell.datatype = "f";
                break;

            case "vtc":
                v = parts[j++];
                cell.valuetype = v;
                if (v.charAt(0) === "n") {
                    cell.datavalue = SocialCalc.decodeFromSave(parts[j++]) - 0;
                } else {
                    cell.datavalue = SocialCalc.decodeFromSave(parts[j++]);
                }
                cell.formula = SocialCalc.decodeFromSave(parts[j++]);
                cell.datatype = "c";
                break;

            case "e":
                cell.errors = SocialCalc.decodeFromSave(parts[j++]);
                break;

            case "b":
                cell.bt = parts[j++] - 0;
                cell.br = parts[j++] - 0;
                cell.bb = parts[j++] - 0;
                cell.bl = parts[j++] - 0;
                break;

            case "l":
                cell.layout = parts[j++] - 0;
                break;

            case "f":
                cell.font = parts[j++] - 0;
                break;

            case "c":
                cell.color = parts[j++] - 0;
                break;

            case "bg":
                cell.bgcolor = parts[j++] - 0;
                break;

            case "cf":
                cell.cellformat = parts[j++] - 0;
                break;

            case "ntvf":
                cell.nontextvalueformat = parts[j++] - 0;
                break;

            case "tvf":
                cell.textvalueformat = parts[j++] - 0;
                break;

            case "colspan":
                cell.colspan = parts[j++] - 0;
                break;

            case "rowspan":
                cell.rowspan = parts[j++] - 0;
                break;

            case "cssc":
                cell.cssc = parts[j++];
                break;

            case "csss":
                cell.csss = SocialCalc.decodeFromSave(parts[j++]);
                break;

            case "mod":
                j += 1;
                break;

            case "comment":
                cell.comment = SocialCalc.decodeFromSave(parts[j++]);
                break;

            default:
                throw `${SocialCalc.Constants.s_cfspUnknownCellType} '${t}'`;
        }
    }
};

/**
 * @constant {Array<string>} SocialCalc.sheetfields
 * @description Sheet field names for basic attributes
 */
SocialCalc.sheetfields = ["defaultrowheight", "defaultcolwidth", "circularreferencecell", "recalc", "needsrecalc"];

/**
 * @constant {Array<string>} SocialCalc.sheetfieldsshort
 * @description Short codes corresponding to sheet fields for save format
 */
SocialCalc.sheetfieldsshort = ["h", "w", "circularreferencecell", "recalc", "needsrecalc"];

/**
 * @constant {Array<string>} SocialCalc.sheetfieldsxlat
 * @description Sheet field names that require translation during save/load
 */
SocialCalc.sheetfieldsxlat = [
    "defaulttextformat",
    "defaultnontextformat",
    "defaulttextvalueformat",
    "defaultnontextvalueformat",
    "defaultcolor",
    "defaultbgcolor",
    "defaultfont",
    "defaultlayout"
];

/**
 * @constant {Array<string>} SocialCalc.sheetfieldsxlatshort
 * @description Short codes for translatable sheet fields
 */
SocialCalc.sheetfieldsxlatshort = ["tf", "ntf", "tvf", "ntvf", "color", "bgcolor", "font", "layout"];

/**
 * @constant {Array<string>} SocialCalc.sheetfieldsxlatxlt
 * @description Translation table types for sheet fields
 */
SocialCalc.sheetfieldsxlatxlt = [
    "cellformat",
    "cellformat",
    "valueformat",
    "valueformat",
    "color",
    "color",
    "font",
    "layout"
];

/**
 * @function SocialCalc.CreateSheetSave
 * @description Creates a text representation of the sheet data in save format
 * @param {SocialCalc.Sheet} sheetobj - The sheet object to save
 * @param {string} [range] - Optional range to save (e.g., "A1:C10"). If present, only those cells are saved as clipboard data
 * @param {boolean} [canonicalize] - Whether to canonicalize the sheet before saving
 * @returns {string} Text representation of the sheet data
 * 
 * @example
 * // Save entire sheet
 * const saveString = SocialCalc.CreateSheetSave(mySheet);
 * 
 * // Save specific range as clipboard data
 * const clipboardData = SocialCalc.CreateSheetSave(mySheet, "A1:C10", true);
 */
SocialCalc.CreateSheetSave = function (sheetobj, range, canonicalize) {
    let cell, cr1, cr2, row, col, coord, line, value, i, name;
    const result = [];
    let prange;

    sheetobj.CanonicalizeSheet(canonicalize || SocialCalc.Constants.doCanonicalizeSheet);
    const xlt = sheetobj.xlt;

    if (range) {
        prange = SocialCalc.ParseRange(range);
    } else {
        prange = {
            cr1: { row: 1, col: 1 },
            cr2: { row: xlt.maxrow, col: xlt.maxcol }
        };
    }
    cr1 = prange.cr1;
    cr2 = prange.cr2;

    result.push("version:1.5");

    // Save cells in the specified range
    for (row = cr1.row; row <= cr2.row; row++) {
        for (col = cr1.col; col <= cr2.col; col++) {
            coord = SocialCalc.crToCoord(col, row);
            cell = sheetobj.cells[coord];
            if (!cell) continue;

            line = sheetobj.CellToString(cell);
            if (line.length === 0) continue; // ignore completely empty cells

            line = `cell:${coord}${line}`;
            result.push(line);
        }
    }

    // Save column attributes
    for (col = 1; col <= xlt.maxcol; col++) {
        coord = SocialCalc.rcColname(col);
        if (sheetobj.colattribs.width[coord]) {
            result.push(`col:${coord}:w:${sheetobj.colattribs.width[coord]}`);
        }
        if (sheetobj.colattribs.hide[coord]) {
            result.push(`col:${coord}:hide:${sheetobj.colattribs.hide[coord]}`);
        }
    }

    // Save row attributes
    for (row = 1; row <= xlt.maxrow; row++) {
        if (sheetobj.rowattribs.height[row]) {
            result.push(`row:${row}:h:${sheetobj.rowattribs.height[row]}`);
        }
        if (sheetobj.rowattribs.hide[row]) {
            result.push(`row:${row}:hide:${sheetobj.rowattribs.hide[row]}`);
        }
    }

    // Save sheet attributes
    line = `sheet:c:${xlt.maxcol}:r:${xlt.maxrow}`;

    // Non-translated values
    for (i = 0; i < SocialCalc.sheetfields.length; i++) {
        value = SocialCalc.encodeForSave(sheetobj.attribs[SocialCalc.sheetfields[i]]);
        if (value) {
            line += `:${SocialCalc.sheetfieldsshort[i]}:${value}`;
        }
    }

    // Translated values
    for (i = 0; i < SocialCalc.sheetfieldsxlat.length; i++) {
        value = sheetobj.attribs[SocialCalc.sheetfieldsxlat[i]];
        if (value) {
            line += `:${SocialCalc.sheetfieldsxlatshort[i]}:${xlt[SocialCalc.sheetfieldsxlatxlt[i] + "sxlat"][value]}`;
        }
    }

    result.push(line);

    // Save style definitions
    for (i = 1; i < xlt.newborderstyles.length; i++) {
        result.push(`border:${i}:${xlt.newborderstyles[i]}`);
    }

    for (i = 1; i < xlt.newcellformats.length; i++) {
        result.push(`cellformat:${i}:${SocialCalc.encodeForSave(xlt.newcellformats[i])}`);
    }

    for (i = 1; i < xlt.newcolors.length; i++) {
        result.push(`color:${i}:${xlt.newcolors[i]}`);
    }

    for (i = 1; i < xlt.newfonts.length; i++) {
        result.push(`font:${i}:${xlt.newfonts[i]}`);
    }

    for (i = 1; i < xlt.newlayouts.length; i++) {
        result.push(`layout:${i}:${xlt.newlayouts[i]}`);
    }

    for (i = 1; i < xlt.newvalueformats.length; i++) {
        result.push(`valueformat:${i}:${SocialCalc.encodeForSave(xlt.newvalueformats[i])}`);
    }

    // Save named ranges
    for (i = 0; i < xlt.namesorder.length; i++) {
        name = xlt.namesorder[i];
        result.push(`name:${SocialCalc.encodeForSave(name).toUpperCase()}:${SocialCalc.encodeForSave(sheetobj.names[name].desc)}:${SocialCalc.encodeForSave(sheetobj.names[name].definition)}`);
    }

    // Add clipboard range info if this is a range save
    if (range) {
        result.push(`copiedfrom:${SocialCalc.crToCoord(cr1.col, cr1.row)}:${SocialCalc.crToCoord(cr2.col, cr2.row)}`);
    }

    result.push(""); // one extra to get extra \n

    delete sheetobj.xlt; // clean up

    return result.join("\n");
};

/**
 * @function SocialCalc.CellToString
 * @description Converts a cell object to its string representation for saving
 * @param {SocialCalc.Sheet} sheet - The sheet object containing the cell
 * @param {SocialCalc.Cell} cell - The cell object to convert
 * @returns {string} String representation of the cell in save format
 * 
 * @example
 * // Convert a cell to save format
 * const cellString = SocialCalc.CellToString(sheet, cell);
 * // Returns something like ":v:42:f:1:c:2"
 */
SocialCalc.CellToString = function (sheet, cell) {
    let line = "";

    if (!cell) return line;

    const value = SocialCalc.encodeForSave(cell.datavalue);

    // Handle different data types
    if (cell.datatype === "v") {
        if (cell.valuetype === "n") {
            line += `:v:${value}`;
        } else {
            line += `:vt:${cell.valuetype}:${value}`;
        }
    } else if (cell.datatype === "t") {
        if (cell.valuetype === SocialCalc.Constants.textdatadefaulttype) {
            line += `:t:${value}`;
        } else {
            line += `:vt:${cell.valuetype}:${value}`;
        }
    } else {
        const formula = SocialCalc.encodeForSave(cell.formula);
        if (cell.datatype === "f") {
            line += `:vtf:${cell.valuetype}:${value}:${formula}`;
        } else if (cell.datatype === "c") {
            line += `:vtc:${cell.valuetype}:${value}:${formula}`;
        }
    }

    // Add error information
    if (cell.errors) {
        line += `:e:${SocialCalc.encodeForSave(cell.errors)}`;
    }

    // Handle borders
    const t = cell.bt || "";
    const r = cell.br || "";
    const b = cell.bb || "";
    const l = cell.bl || "";

    if (sheet.xlt) { // if have canonical save info
        const xlt = sheet.xlt;
        if (t || r || b || l) {
            line += `:b:${xlt.borderstylesxlat[t || 0]}:${xlt.borderstylesxlat[r || 0]}:${xlt.borderstylesxlat[b || 0]}:${xlt.borderstylesxlat[l || 0]}`;
        }
        if (cell.layout) line += `:l:${xlt.layoutsxlat[cell.layout]}`;
        if (cell.font) line += `:f:${xlt.fontsxlat[cell.font]}`;
        if (cell.color) line += `:c:${xlt.colorsxlat[cell.color]}`;
        if (cell.bgcolor) line += `:bg:${xlt.colorsxlat[cell.bgcolor]}`;
        if (cell.cellformat) line += `:cf:${xlt.cellformatsxlat[cell.cellformat]}`;
        if (cell.textvalueformat) line += `:tvf:${xlt.valueformatsxlat[cell.textvalueformat]}`;
        if (cell.nontextvalueformat) line += `:ntvf:${xlt.valueformatsxlat[cell.nontextvalueformat]}`;
    } else {
        if (t || r || b || l) {
            line += `:b:${t}:${r}:${b}:${l}`;
        }
        if (cell.layout) line += `:l:${cell.layout}`;
        if (cell.font) line += `:f:${cell.font}`;
        if (cell.color) line += `:c:${cell.color}`;
        if (cell.bgcolor) line += `:bg:${cell.bgcolor}`;
        if (cell.cellformat) line += `:cf:${cell.cellformat}`;
        if (cell.textvalueformat) line += `:tvf:${cell.textvalueformat}`;
        if (cell.nontextvalueformat) line += `:ntvf:${cell.nontextvalueformat}`;
    }

    // Add span and styling information
    if (cell.colspan) line += `:colspan:${cell.colspan}`;
    if (cell.rowspan) line += `:rowspan:${cell.rowspan}`;
    if (cell.cssc) line += `:cssc:${cell.cssc}`;
    if (cell.csss) line += `:csss:${SocialCalc.encodeForSave(cell.csss)}`;
    if (cell.mod) line += `:mod:${cell.mod}`;
    if (cell.comment) line += `:comment:${SocialCalc.encodeForSave(cell.comment)}`;

    return line;
};
/**
 * @function SocialCalc.CanonicalizeSheet
 * @description Goes through the sheet and fills in sheetobj.xlt with canonicalized data
 * @param {SocialCalc.Sheet} sheetobj - The sheet object to canonicalize
 * @param {boolean} [full] - Whether to perform full canonicalization
 * 
 * @description Creates sheetobj.xlt with the following properties:
 * - .maxrow, .maxcol - lastrow and lastcol as small as possible
 * - .newlayouts - new version without unused ones, in ascending order
 * - .layoutsxlat - maps old layouts index to new one
 * - same ".new" and ".xlat" for fonts, colors, borderstyles, cell and value formats
 * - .namesorder - array with names sorted
 * 
 * If full or SocialCalc.Constants.doCanonicalizeSheet are not true, 
 * values will remain unchanged (to save time, etc.)
 * 
 * Note: sheetobj.xlt should be deleted when finished using it
 * 
 * @example
 * // Canonicalize a sheet for saving
 * SocialCalc.CanonicalizeSheet(mySheet, true);
 * // Use mySheet.xlt for processing
 * delete mySheet.xlt; // Clean up when done
 */
SocialCalc.CanonicalizeSheet = function (sheetobj, full) {
    let coord, cr, cell, filled, an, a, newa, newxlat, used, ahash, i, v;
    let maxrow = 0;
    let maxcol = 0;

    const alist = ["borderstyle", "cellformat", "color", "font", "layout", "valueformat"];
    const xlt = {};

    // Always return a sorted list of names
    xlt.namesorder = [];
    for (a in sheetobj.names) {
        xlt.namesorder.push(a);
    }
    xlt.namesorder.sort();

    // Return make-no-changes values if canonicalization not wanted
    if (!SocialCalc.Constants.doCanonicalizeSheet || !full) {
        for (an = 0; an < alist.length; an++) {
            a = alist[an];
            xlt[`new${a}s`] = sheetobj[`${a}s`];
            const l = sheetobj[`${a}s`].length;
            newxlat = new Array(l);
            newxlat[0] = "";
            for (i = 1; i < l; i++) {
                newxlat[i] = i;
            }
            xlt[`${a}sxlat`] = newxlat;
        }

        xlt.maxrow = sheetobj.attribs.lastrow;
        xlt.maxcol = sheetobj.attribs.lastcol;
        sheetobj.xlt = xlt;
        return;
    }

    // Initialize usage tracking for all attribute types
    for (an = 0; an < alist.length; an++) {
        a = alist[an];
        xlt[`${a}sUsed`] = {};
    }

    // Create convenient references to usage objects
    const colorsUsed = xlt.colorsUsed;
    const borderstylesUsed = xlt.borderstylesUsed;
    const fontsUsed = xlt.fontsUsed;
    const layoutsUsed = xlt.layoutsUsed;
    const cellformatsUsed = xlt.cellformatsUsed;
    const valueformatsUsed = xlt.valueformatsUsed;

    // Check all cells to see which values are used
    for (coord in sheetobj.cells) {
        cr = SocialCalc.coordToCr(coord);
        cell = sheetobj.cells[coord];
        filled = false;

        if (cell.valuetype && cell.valuetype !== "b") filled = true;

        // Track color usage
        if (cell.color) {
            colorsUsed[cell.color] = 1;
            filled = true;
        }

        if (cell.bgcolor) {
            colorsUsed[cell.bgcolor] = 1;
            filled = true;
        }

        // Track border style usage
        if (cell.bt) {
            borderstylesUsed[cell.bt] = 1;
            filled = true;
        }
        if (cell.br) {
            borderstylesUsed[cell.br] = 1;
            filled = true;
        }
        if (cell.bb) {
            borderstylesUsed[cell.bb] = 1;
            filled = true;
        }
        if (cell.bl) {
            borderstylesUsed[cell.bl] = 1;
            filled = true;
        }

        // Track layout usage
        if (cell.layout) {
            layoutsUsed[cell.layout] = 1;
            filled = true;
        }

        // Track font usage
        if (cell.font) {
            fontsUsed[cell.font] = 1;
            filled = true;
        }

        // Track format usage
        if (cell.cellformat) {
            cellformatsUsed[cell.cellformat] = 1;
            filled = true;
        }

        if (cell.textvalueformat) {
            valueformatsUsed[cell.textvalueformat] = 1;
            filled = true;
        }

        if (cell.nontextvalueformat) {
            valueformatsUsed[cell.nontextvalueformat] = 1;
            filled = true;
        }

        // Update max row/col if cell has content
        if (filled) {
            if (cr.row > maxrow) maxrow = cr.row;
            if (cr.col > maxcol) maxcol = cr.col;
        }
    }

    // Process sheet-level values
    for (i = 0; i < SocialCalc.sheetfieldsxlat.length; i++) {
        v = sheetobj.attribs[SocialCalc.sheetfieldsxlat[i]];
        if (v) {
            xlt[`${SocialCalc.sheetfieldsxlatxlt[i]}sUsed`][v] = 1;
        }
    }

    // Check explicit row settings
    const rowAttribs = { "height": 1, "hide": 1 };
    for (v in rowAttribs) {
        for (cr in sheetobj.rowattribs[v]) {
            if (cr > maxrow) maxrow = cr;
        }
    }

    // Check explicit column settings
    const colAttribs = { "hide": 1, "width": 1 };
    for (v in colAttribs) {
        for (coord in sheetobj.colattribs[v]) {
            cr = SocialCalc.coordToCr(`${coord}1`);
            if (cr.col > maxcol) maxcol = cr.col;
        }
    }

    // Build canonicalized attribute arrays
    for (an = 0; an < alist.length; an++) {
        a = alist[an];

        newa = [];
        used = xlt[`${a}sUsed`];
        for (v in used) {
            newa.push(sheetobj[`${a}s`][v]);
        }
        newa.sort();
        newa.unshift("");

        newxlat = [""];
        ahash = sheetobj[`${a}hash`];

        for (i = 1; i < newa.length; i++) {
            newxlat[ahash[newa[i]]] = i;
        }

        xlt[`${a}sxlat`] = newxlat;
        xlt[`new${a}s`] = newa;
    }

    xlt.maxrow = maxrow || 1;
    xlt.maxcol = maxcol || 1;

    sheetobj.xlt = xlt; // leave for use by caller
};

/**
 * @function SocialCalc.EncodeCellAttributes
 * @description Returns the cell's attributes in an object with def/val properties
 * @param {SocialCalc.Sheet} sheet - The sheet containing the cell
 * @param {string} coord - The cell coordinate (e.g., "A1")
 * @returns {Object} Object with attributes in form: {attribname: {def: boolean, val: string}}
 * 
 * @example
 * // Get cell attributes
 * const attrs = SocialCalc.EncodeCellAttributes(sheet, "A1");
 * // attrs.alignhoriz = {def: false, val: "left"}
 * // attrs.textcolor = {def: true, val: ""}  // using default
 */
SocialCalc.EncodeCellAttributes = function (sheet, coord) {
    let value, i, b, bb, parts;
    const result = {};

    /**
     * @description Initialize an attribute with default values
     * @param {string} name - The attribute name
     */
    const InitAttrib = (name) => {
        result[name] = { def: true, val: "" };
    };

    /**
     * @description Initialize multiple attributes with default values
     * @param {Array<string>} namelist - Array of attribute names
     */
    const InitAttribs = (namelist) => {
        for (let i = 0; i < namelist.length; i++) {
            InitAttrib(namelist[i]);
        }
    };

    /**
     * @description Set an attribute value
     * @param {string} name - The attribute name
     * @param {*} v - The value to set
     */
    const SetAttrib = (name, v) => {
        result[name].def = false;
        result[name].val = v || "";
    };

    /**
     * @description Set attribute value, ignoring "*" values
     * @param {string} name - The attribute name
     * @param {*} v - The value to set
     */
    const SetAttribStar = (name, v) => {
        if (v === "*") return;
        result[name].def = false;
        result[name].val = v;
    };

    const cell = sheet.GetAssuredCell(coord);

    // cellformat: alignhoriz
    InitAttrib("alignhoriz");
    if (cell.cellformat) {
        SetAttrib("alignhoriz", sheet.cellformats[cell.cellformat]);
    }

    // layout: alignvert, padtop, padright, padbottom, padleft
    InitAttribs(["alignvert", "padtop", "padright", "padbottom", "padleft"]);
    if (cell.layout) {
        parts = sheet.layouts[cell.layout].match(/^padding:\s*(\S+)\s+(\S+)\s+(\S+)\s+(\S+);vertical-align:\s*(\S+);/);
        SetAttribStar("padtop", parts[1]);
        SetAttribStar("padright", parts[2]);
        SetAttribStar("padbottom", parts[3]);
        SetAttribStar("padleft", parts[4]);
        SetAttribStar("alignvert", parts[5]);
    }

    // font: fontfamily, fontlook, fontsize
    InitAttribs(["fontfamily", "fontlook", "fontsize"]);
    if (cell.font) {
        parts = sheet.fonts[cell.font].match(/^(\*|\S+? \S+?) (\S+?) (\S.*)$/);
        SetAttribStar("fontfamily", parts[3]);
        SetAttribStar("fontsize", parts[2]);
        SetAttribStar("fontlook", parts[1]);
    }

    // color: textcolor
    InitAttrib("textcolor");
    if (cell.color) {
        SetAttrib("textcolor", sheet.colors[cell.color]);
    }

    // bgcolor: bgcolor
    InitAttrib("bgcolor");
    if (cell.bgcolor) {
        SetAttrib("bgcolor", sheet.colors[cell.bgcolor]);
    }

    // formatting: numberformat, textformat
    InitAttribs(["numberformat", "textformat"]);
    if (cell.nontextvalueformat) {
        SetAttrib("numberformat", sheet.valueformats[cell.nontextvalueformat]);
    }
    if (cell.textvalueformat) {
        SetAttrib("textformat", sheet.valueformats[cell.textvalueformat]);
    }

    // merges: colspan, rowspan
    InitAttribs(["colspan", "rowspan"]);
    SetAttrib("colspan", cell.colspan || 1);
    SetAttrib("rowspan", cell.rowspan || 1);

    // borders: bXthickness, bXstyle, bXcolor for X = t, r, b, and l
    for (i = 0; i < 4; i++) {
        b = "trbl".charAt(i);
        bb = `b${b}`;
        InitAttrib(bb);
        SetAttrib(bb, cell[bb] ? sheet.borderstyles[cell[bb]] : "");
        InitAttrib(`${bb}thickness`);
        InitAttrib(`${bb}style`);
        InitAttrib(`${bb}color`);
        if (cell[bb]) {
            parts = sheet.borderstyles[cell[bb]].match(/(\S+)\s+(\S+)\s+(\S.+)/);
            SetAttrib(`${bb}thickness`, parts[1]);
            SetAttrib(`${bb}style`, parts[2]);
            SetAttrib(`${bb}color`, parts[3]);
        }
    }

    // misc: cssc, csss, mod
    InitAttribs(["cssc", "csss", "mod"]);
    SetAttrib("cssc", cell.cssc || "");
    SetAttrib("csss", cell.csss || "");
    SetAttrib("mod", cell.mod || "n");

    return result;
};

/**
 * @function SocialCalc.EncodeSheetAttributes
 * @description Returns the sheet's attributes in an object with def/val properties
 * @param {SocialCalc.Sheet} sheet - The sheet to encode attributes for
 * @returns {Object} Object with attributes in form: {attribname: {def: boolean, val: string}}
 * 
 * @example
 * // Get sheet attributes
 * const attrs = SocialCalc.EncodeSheetAttributes(sheet);
 * // attrs.colwidth = {def: false, val: "80"}
 * // attrs.textcolor = {def: true, val: ""}  // using default
 */
SocialCalc.EncodeSheetAttributes = function (sheet) {
    let value, parts;
    const attribs = sheet.attribs;
    const result = {};

    /**
     * @description Initialize an attribute with default values
     * @param {string} name - The attribute name
     */
    const InitAttrib = (name) => {
        result[name] = { def: true, val: "" };
    };

    /**
     * @description Initialize multiple attributes with default values
     * @param {Array<string>} namelist - Array of attribute names
     */
    const InitAttribs = (namelist) => {
        for (let i = 0; i < namelist.length; i++) {
            InitAttrib(namelist[i]);
        }
    };

    /**
     * @description Set an attribute value
     * @param {string} name - The attribute name
     * @param {*} v - The value to set
     */
    const SetAttrib = (name, v) => {
        result[name].def = false;
        result[name].val = v || value;
    };

    /**
     * @description Set attribute value, ignoring "*" values
     * @param {string} name - The attribute name
     * @param {*} v - The value to set
     */
    const SetAttribStar = (name, v) => {
        if (v === "*") return;
        result[name].def = false;
        result[name].val = v;
    };

    // sizes: colwidth, rowheight
    InitAttrib("colwidth");
    if (attribs.defaultcolwidth) {
        SetAttrib("colwidth", attribs.defaultcolwidth);
    }

    InitAttrib("rowheight");
    if (attribs.rowheight) {
        SetAttrib("rowheight", attribs.defaultrowheight);
    }

    // cellformat: textalignhoriz, numberalignhoriz
    InitAttrib("textalignhoriz");
    if (attribs.defaulttextformat) {
        SetAttrib("textalignhoriz", sheet.cellformats[attribs.defaulttextformat]);
    }

    InitAttrib("numberalignhoriz");
    if (attribs.defaultnontextformat) {
        SetAttrib("numberalignhoriz", sheet.cellformats[attribs.defaultnontextformat]);
    }

    // layout: alignvert, padtop, padright, padbottom, padleft
    InitAttribs(["alignvert", "padtop", "padright", "padbottom", "padleft"]);
    if (attribs.defaultlayout) {
        parts = sheet.layouts[attribs.defaultlayout].match(/^padding:\s*(\S+)\s+(\S+)\s+(\S+)\s+(\S+);vertical-align:\s*(\S+);/);
        SetAttribStar("padtop", parts[1]);
        SetAttribStar("padright", parts[2]);
        SetAttribStar("padbottom", parts[3]);
        SetAttribStar("padleft", parts[4]);
        SetAttribStar("alignvert", parts[5]);
    }

    // font: fontfamily, fontlook, fontsize
    InitAttribs(["fontfamily", "fontlook", "fontsize"]);
    if (attribs.defaultfont) {
        parts = sheet.fonts[attribs.defaultfont].match(/^(\*|\S+? \S+?) (\S+?) (\S.*)$/);
        SetAttribStar("fontfamily", parts[3]);
        SetAttribStar("fontsize", parts[2]);
        SetAttribStar("fontlook", parts[1]);
    }

    // color: textcolor
    InitAttrib("textcolor");
    if (attribs.defaultcolor) {
        SetAttrib("textcolor", sheet.colors[attribs.defaultcolor]);
    }

    // bgcolor: bgcolor
    InitAttrib("bgcolor");
    if (attribs.defaultbgcolor) {
        SetAttrib("bgcolor", sheet.colors[attribs.defaultbgcolor]);
    }

    // formatting: numberformat, textformat
    InitAttribs(["numberformat", "textformat"]);
    if (attribs.defaultnontextvalueformat) {
        SetAttrib("numberformat", sheet.valueformats[attribs.defaultnontextvalueformat]);
    }
    if (attribs.defaulttextvalueformat) {
        SetAttrib("textformat", sheet.valueformats[attribs.defaulttextvalueformat]);
    }

    // recalc: recalc
    InitAttrib("recalc");
    if (attribs.recalc) {
        SetAttrib("recalc", attribs.recalc);
    }

    return result;
};
/**
 * @function SocialCalc.DecodeCellAttributes
 * @description Takes cell attributes in an object and returns sheet commands to make the actual attributes correspond
 * @param {SocialCalc.Sheet} sheet - The sheet containing the cell
 * @param {string} coord - The cell coordinate (e.g., "A1")
 * @param {Object} newattribs - Object with attributes in form: {attribname: {def: boolean, val: string}}
 * @param {string} [range] - Optional range to apply commands to (e.g., "A1:C10")
 * @returns {string|null} Command string to execute, or null if no changes needed
 * 
 * @example
 * // Decode cell attributes and get commands
 * const attribs = {
 *   alignhoriz: {def: false, val: "center"},
 *   textcolor: {def: true, val: ""}
 * };
 * const cmdstr = SocialCalc.DecodeCellAttributes(sheet, "A1", attribs, "A1:C10");
 * // Returns: "set A1:C10 cellformat center"
 */
SocialCalc.DecodeCellAttributes = function (sheet, coord, newattribs, range) {
    let value, b, bb;
    const cell = sheet.GetAssuredCell(coord);
    let changed = false;
    let cmdstr = "";

    /**
     * @description Check if an attribute has changed and add command if needed
     * @param {string} attribname - Name of the attribute to check
     * @param {string} oldval - Current value of the attribute
     * @param {string} cmdname - Command name to use if changed
     */
    const CheckChanges = (attribname, oldval, cmdname) => {
        let val;
        if (newattribs[attribname]) {
            if (newattribs[attribname].def) {
                val = "";
            } else {
                val = newattribs[attribname].val;
            }
            if (val !== (oldval || "")) {
                DoCmd(`${cmdname} ${val}`);
            }
        }
    };

    /**
     * @description Add a command to the command string
     * @param {string} str - The command to add
     */
    const DoCmd = (str) => {
        if (cmdstr) cmdstr += "\n";
        cmdstr += `set ${range || coord} ${str}`;
        changed = true;
    };

    // cellformat: alignhoriz
    CheckChanges("alignhoriz", sheet.cellformats[cell.cellformat], "cellformat");

    // layout: alignvert, padtop, padright, padbottom, padleft
    if (!newattribs.alignvert.def || !newattribs.padtop.def || !newattribs.padright.def ||
        !newattribs.padbottom.def || !newattribs.padleft.def) {
        value = "padding:" +
            (newattribs.padtop.def ? "* " : `${newattribs.padtop.val} `) +
            (newattribs.padright.def ? "* " : `${newattribs.padright.val} `) +
            (newattribs.padbottom.def ? "* " : `${newattribs.padbottom.val} `) +
            (newattribs.padleft.def ? "*" : newattribs.padleft.val) +
            ";vertical-align:" +
            (newattribs.alignvert.def ? "*;" : `${newattribs.alignvert.val};`);
    } else {
        value = "";
    }

    if (value !== (sheet.layouts[cell.layout] || "")) {
        DoCmd(`layout ${value}`);
    }

    // font: fontfamily, fontlook, fontsize
    if (!newattribs.fontlook.def || !newattribs.fontsize.def || !newattribs.fontfamily.def) {
        value =
            (newattribs.fontlook.def ? "* " : `${newattribs.fontlook.val} `) +
            (newattribs.fontsize.def ? "* " : `${newattribs.fontsize.val} `) +
            (newattribs.fontfamily.def ? "*" : newattribs.fontfamily.val);
    } else {
        value = "";
    }

    if (value !== (sheet.fonts[cell.font] || "")) {
        DoCmd(`font ${value}`);
    }

    // color: textcolor
    CheckChanges("textcolor", sheet.colors[cell.color], "color");

    // bgcolor: bgcolor
    CheckChanges("bgcolor", sheet.colors[cell.bgcolor], "bgcolor");

    // formatting: numberformat, textformat
    CheckChanges("numberformat", sheet.valueformats[cell.nontextvalueformat], "nontextvalueformat");
    CheckChanges("textformat", sheet.valueformats[cell.textvalueformat], "textvalueformat");

    // merges: colspan, rowspan - NOT HANDLED: IGNORED!

    // borders: bX for X = t, r, b, and l; bXthickness, bXstyle, bXcolor ignored
    for (let i = 0; i < 4; i++) {
        b = "trbl".charAt(i);
        bb = `b${b}`;
        CheckChanges(bb, sheet.borderstyles[cell[bb]], bb);
    }

    // misc: cssc, csss, mod
    CheckChanges("cssc", cell.cssc, "cssc");
    CheckChanges("csss", cell.csss, "csss");

    if (newattribs.mod) {
        if (newattribs.mod.def) {
            value = "n";
        } else {
            value = newattribs.mod.val;
        }
        if (value !== (cell.mod || "n")) {
            if (value === "n") value = ""; // restrict to "y" and "" normally
            DoCmd(`mod ${value}`);
        }
    }

    // Return commands if any changes were made
    return changed ? cmdstr : null;
};

/**
 * @function SocialCalc.DecodeSheetAttributes
 * @description Takes sheet attributes and returns commands to make the actual attributes correspond
 * @param {SocialCalc.Sheet} sheet - The sheet to modify
 * @param {Object} newattribs - Object with attributes in form: {attribname: {def: boolean, val: string}}
 * @returns {string|null} Command string to execute, or null if no changes needed
 * 
 * @example
 * // Decode sheet attributes and get commands
 * const attribs = {
 *   colwidth: {def: false, val: "100"},
 *   textcolor: {def: true, val: ""}
 * };
 * const cmdstr = SocialCalc.DecodeSheetAttributes(sheet, attribs);
 * // Returns: "set sheet defaultcolwidth 100"
 */
SocialCalc.DecodeSheetAttributes = function (sheet, newattribs) {
    let value;
    const attribs = sheet.attribs;
    let changed = false;
    let cmdstr = "";

    /**
     * @description Check if an attribute has changed and add command if needed
     * @param {string} attribname - Name of the attribute to check
     * @param {string} oldval - Current value of the attribute
     * @param {string} cmdname - Command name to use if changed
     */
    const CheckChanges = (attribname, oldval, cmdname) => {
        let val;
        if (newattribs[attribname]) {
            if (newattribs[attribname].def) {
                val = "";
            } else {
                val = newattribs[attribname].val;
            }
            if (val !== (oldval || "")) {
                DoCmd(`${cmdname} ${val}`);
            }
        }
    };

    /**
     * @description Add a command to the command string
     * @param {string} str - The command to add
     */
    const DoCmd = (str) => {
        if (cmdstr) cmdstr += "\n";
        cmdstr += `set sheet ${str}`;
        changed = true;
    };

    // sizes: colwidth, rowheight
    CheckChanges("colwidth", attribs.defaultcolwidth, "defaultcolwidth");
    CheckChanges("rowheight", attribs.defaultrowheight, "defaultrowheight");

    // cellformat: textalignhoriz, numberalignhoriz
    CheckChanges("textalignhoriz", sheet.cellformats[attribs.defaulttextformat], "defaulttextformat");
    CheckChanges("numberalignhoriz", sheet.cellformats[attribs.defaultnontextformat], "defaultnontextformat");

    // layout: alignvert, padtop, padright, padbottom, padleft
    if (!newattribs.alignvert.def || !newattribs.padtop.def || !newattribs.padright.def ||
        !newattribs.padbottom.def || !newattribs.padleft.def) {
        value = "padding:" +
            (newattribs.padtop.def ? "* " : `${newattribs.padtop.val} `) +
            (newattribs.padright.def ? "* " : `${newattribs.padright.val} `) +
            (newattribs.padbottom.def ? "* " : `${newattribs.padbottom.val} `) +
            (newattribs.padleft.def ? "*" : newattribs.padleft.val) +
            ";vertical-align:" +
            (newattribs.alignvert.def ? "*;" : `${newattribs.alignvert.val};`);
    } else {
        value = "";
    }

    if (value !== (sheet.layouts[attribs.defaultlayout] || "")) {
        DoCmd(`defaultlayout ${value}`);
    }

    // font: fontfamily, fontlook, fontsize
    if (!newattribs.fontlook.def || !newattribs.fontsize.def || !newattribs.fontfamily.def) {
        value =
            (newattribs.fontlook.def ? "* " : `${newattribs.fontlook.val} `) +
            (newattribs.fontsize.def ? "* " : `${newattribs.fontsize.val} `) +
            (newattribs.fontfamily.def ? "*" : newattribs.fontfamily.val);
    } else {
        value = "";
    }

    if (value !== (sheet.fonts[attribs.defaultfont] || "")) {
        DoCmd(`defaultfont ${value}`);
    }

    // color: textcolor
    CheckChanges("textcolor", sheet.colors[attribs.defaultcolor], "defaultcolor");

    // bgcolor: bgcolor
    CheckChanges("bgcolor", sheet.colors[attribs.defaultbgcolor], "defaultbgcolor");

    // formatting: numberformat, textformat
    CheckChanges("numberformat", sheet.valueformats[attribs.defaultnontextvalueformat], "defaultnontextvalueformat");
    CheckChanges("textformat", sheet.valueformats[attribs.defaulttextvalueformat], "defaulttextvalueformat");

    // recalc: recalc
    CheckChanges("recalc", sheet.attribs.recalc, "recalc");

    // Return commands if any changes were made
    return changed ? cmdstr : null;
};

/**
 * @namespace SocialCalc.SheetCommandInfo
 * @description Object with information used during command execution (singleton)
 * @property {SocialCalc.Sheet|null} sheetobj - Sheet being operated on
 * @property {SocialCalc.Parse|null} parseobj - Parse object with command string
 * @property {number|null} timerobj - Timer object used for timeslicing
 * @property {number} firsttimerdelay - Initial delay before starting commands (ms)
 * @property {number} timerdelay - Delay between command slices (ms)
 * @property {number} maxtimeslice - Maximum time per slice (ms)
 * @property {boolean} saveundo - Whether to save undo information
 * @property {Object} CmdExtensionCallbacks - Command extension callbacks
 * @property {string} cmdextensionbusy - Command extension busy flag
 */
SocialCalc.SheetCommandInfo = {
    sheetobj: null,
    parseobj: null,
    timerobj: null,
    firsttimerdelay: 50, // wait before starting cmds (for Chrome - to give time to update)
    timerdelay: 1, // wait between slices
    maxtimeslice: 100, // do another slice after this many milliseconds
    saveundo: false, // arg for ExecuteSheetCommand

    /**
     * @description Command extension callbacks in form: 
     * cmdname: {func: function(cmdname, data, sheet, parseobj, saveundo), data: whatever}
     */
    CmdExtensionCallbacks: {},

    /**
     * @description If length > 0, command loop waits for SocialCalc.ResumeFromCmdExtension()
     */
    cmdextensionbusy: ""
};

/**
 * @function SocialCalc.ScheduleSheetCommands
 * @description Schedules sheet commands for execution with timeslicing support
 * @param {SocialCalc.Sheet} sheet - The sheet to execute commands on
 * @param {string} cmdstr - The command string to execute
 * @param {boolean} [saveundo] - Whether to save undo information
 * @param {boolean} [isRemote] - Whether this is a remote command (affects broadcasting)
 * 
 * @description Status callback is called at the beginning (cmdstart) and end (cmdend).
 * Commands are executed with timeslicing to prevent blocking the UI.
 * 
 * @example
 * // Schedule commands for execution
 * SocialCalc.ScheduleSheetCommands(sheet, "set A1 value 42\nset B1 value 24", true);
 */
SocialCalc.ScheduleSheetCommands = function (sheet, cmdstr, saveundo, isRemote) {
    // Broadcast command if callback exists and not remote
    if (SocialCalc.Callbacks.broadcast && !isRemote) {
        if (cmdstr !== 'redisplay' &&
            cmdstr !== 'set sheet defaulttextvalueformat text-wiki' &&
            cmdstr !== 'recalc') {
            SocialCalc.Callbacks.broadcast('execute', {
                cmdtype: "scmd",
                id: sheet.sheetid,
                cmdstr: cmdstr,
                saveundo: saveundo
            });
        }
    }

    const sci = SocialCalc.SheetCommandInfo;

    sci.sheetobj = sheet;
    sci.parseobj = new SocialCalc.Parse(cmdstr);
    sci.saveundo = saveundo;

    // Notify status callback if requested
    if (sci.sheetobj.statuscallback) {
        sheet.statuscallback(sci, "cmdstart", "", sci.sheetobj.statuscallbackparams);
    }

    // Add undo step if requested
    if (sci.saveundo) {
        sci.sheetobj.changes.PushChange(""); // add a step to undo stack
    }

    // Start timer for command execution
    sci.timerobj = window.setTimeout(SocialCalc.SheetCommandsTimerRoutine, sci.firsttimerdelay);
};

/**
 * @function SocialCalc.SheetCommandsTimerRoutine
 * @description Timer routine that executes sheet commands with timeslicing
 * @description Processes commands until EOF or time slice exceeded, then yields control
 * 
 * @example
 * // This function is called automatically by the timer system
 * // It processes commands in chunks to avoid blocking the UI
 */
SocialCalc.SheetCommandsTimerRoutine = function () {
    let errortext;
    const sci = SocialCalc.SheetCommandInfo;
    const starttime = new Date();

    sci.timerobj = null;

    // Process commands until EOF or time limit reached
    while (!sci.parseobj.EOF()) {
        errortext = SocialCalc.ExecuteSheetCommand(sci.sheetobj, sci.parseobj, sci.saveundo);
        if (errortext) alert(errortext);

        sci.parseobj.NextLine();

        // Check for command extension busy state
        if (sci.cmdextensionbusy.length > 0) {
            if (sci.sheetobj.statuscallback) {
                sci.sheetobj.statuscallback(sci, "cmdextension", sci.cmdextensionbusy, sci.sheetobj.statuscallbackparams);
            }
            return;
        }

        // Yield control if taking too long
        if (((new Date()) - starttime) >= sci.maxtimeslice) {
            sci.timerobj = window.setTimeout(SocialCalc.SheetCommandsTimerRoutine, sci.timerdelay);
            return;
        }
    }

    // Notify completion
    if (sci.sheetobj.statuscallback) {
        sci.sheetobj.statuscallback(sci, "cmdend", "", sci.sheetobj.statuscallbackparams);
    }
};

/**
 * @function SocialCalc.ResumeFromCmdExtension
 * @description Resumes command execution after a command extension completes
 * @description Clears the busy flag and continues processing commands
 * 
 * @example
 * // Called by command extensions when they complete
 * SocialCalc.ResumeFromCmdExtension();
 */
SocialCalc.ResumeFromCmdExtension = function () {
    const sci = SocialCalc.SheetCommandInfo;
    sci.cmdextensionbusy = "";
    SocialCalc.SheetCommandsTimerRoutine();
};
/**
 * @function SocialCalc.ExecuteSheetCommand
 * @description Executes commands that modify sheet data with undo support and proper state management
 * @param {SocialCalc.Sheet} sheet - The sheet object to execute commands on
 * @param {SocialCalc.Parse} cmd - Parse object containing the command string
 * @param {boolean} [saveundo] - Whether to save undo information in sheet.changes
 * @returns {string} Error text if command failed, empty string if successful
 * 
 * @description Sets sheet "needsrecalc" and "changedrendervalues" flags as needed.
 * The cmd string may contain multiple commands separated by newlines.
 * Only one "step" is put on the undo stack representing all commands.
 * Text values are encoded (newline => \n, \ => \b, : => \c).
 * 
 * @description Supported commands:
 * - set sheet attributename value (plus lastcol and lastrow)
 * - set 22 attributename value
 * - set B attributename value  
 * - set A1 attributename value1 value2...
 * - set A1:B5 attributename value1 value2...
 * - erase/copy/cut/paste/fillright/filldown A1:B5 all/formulas/format
 * - loadclipboard save-encoded-clipboard-data
 * - clearclipboard
 * - merge C3:F3
 * - unmerge C3
 * - insertcol/insertrow C5
 * - deletecol/deleterow C5:E7
 * - movepaste/moveinsert A1:B5 A8 all/formulas/format
 * - sort cr1:cr2 col1 up/down col2 up/down col3 up/down
 * - name define NAME definition
 * - name desc NAME description
 * - name delete NAME
 * - recalc
 * - redisplay
 * - changedrendervalues
 * - startcmdextension extension rest-of-command
 * 
 * @example
 * // Execute a single command
 * const errorText = SocialCalc.ExecuteSheetCommand(sheet, parseObj, true);
 * if (errorText) console.error("Command failed:", errorText);
 * 
 * // Execute multiple commands
 * const cmd = new SocialCalc.Parse("set A1 value n 42\nset B1 text t Hello");
 * SocialCalc.ExecuteSheetCommand(sheet, cmd, true);
 */
SocialCalc.ExecuteSheetCommand = function (sheet, cmd, saveundo) {
    let cmdstr, cmd1, rest, what, attrib, num, pos, pos2, errortext, undostart, val;
    let cr1, cr2, col, row, cr, cell, newcell;
    let fillright, rowstart, colstart, crbase, rowoffset, coloffset, basecell;
    let clipsheet, cliprange, numcols, numrows, attribtable;
    let colend, rowend, newcolstart, newrowstart, newcolend, newrowend, rownext, colnext, colthis, cellnext;
    let lastrow, lastcol, rowbefore, colbefore, oldformula, oldcr;
    let cols, dirs, lastsortcol, i, sortlist, sortcells, sortvalues, sorttypes;
    let sortfunction, slen, valtype, originalrow, sortedcr;
    let name, v1, v2;
    let cmdextension;

    const attribs = sheet.attribs;
    const changes = sheet.changes;
    const cellProperties = SocialCalc.CellProperties;
    const scc = SocialCalc.Constants;

    /**
     * @description Internal function to parse range and update sheet bounds
     * @inner
     */
    const ParseRange = () => {
        const prange = SocialCalc.ParseRange(what);
        cr1 = prange.cr1;
        cr2 = prange.cr2;
        if (cr2.col > attribs.lastcol) attribs.lastcol = cr2.col;
        if (cr2.row > attribs.lastrow) attribs.lastrow = cr2.row;
    };

    errortext = "";
    cmdstr = cmd.RestOfStringNoMove();

    if (saveundo) {
        sheet.changes.AddDo(cmdstr);
    }

    cmd1 = cmd.NextToken();

    switch (cmd1) {
        case "set":
            what = cmd.NextToken();
            attrib = cmd.NextToken();
            rest = cmd.RestOfString();
            undostart = `set ${what} ${attrib}`;

            if (what === "sheet") {
                sheet.renderneeded = true;
                switch (attrib) {
                    case "defaultcolwidth":
                        if (saveundo) changes.AddUndo(undostart, attribs[attrib]);
                        attribs[attrib] = rest;
                        break;

                    case "defaultcolor":
                    case "defaultbgcolor":
                        if (saveundo) changes.AddUndo(undostart, sheet.GetStyleString("color", attribs[attrib]));
                        attribs[attrib] = sheet.GetStyleNum("color", rest);
                        break;

                    case "defaultlayout":
                        if (saveundo) changes.AddUndo(undostart, sheet.GetStyleString("layout", attribs[attrib]));
                        attribs[attrib] = sheet.GetStyleNum("layout", rest);
                        break;

                    case "defaultfont":
                        if (saveundo) changes.AddUndo(undostart, sheet.GetStyleString("font", attribs[attrib]));
                        if (rest === "* * *") rest = ""; // all default
                        attribs[attrib] = sheet.GetStyleNum("font", rest);
                        break;

                    case "defaulttextformat":
                    case "defaultnontextformat":
                        if (saveundo) changes.AddUndo(undostart, sheet.GetStyleString("cellformat", attribs[attrib]));
                        attribs[attrib] = sheet.GetStyleNum("cellformat", rest);
                        break;

                    case "defaulttextvalueformat":
                    case "defaultnontextvalueformat":
                        if (saveundo) changes.AddUndo(undostart, sheet.GetStyleString("valueformat", attribs[attrib]));
                        attribs[attrib] = sheet.GetStyleNum("valueformat", rest);
                        // Forget all cached display strings
                        for (cr in sheet.cells) {
                            delete sheet.cells[cr].displaystring;
                        }
                        break;

                    case "lastcol":
                    case "lastrow":
                        if (saveundo) changes.AddUndo(undostart, attribs[attrib] - 0);
                        num = rest - 0;
                        if (typeof num === "number") attribs[attrib] = num > 0 ? num : 1;
                        break;

                    case "recalc":
                        if (saveundo) changes.AddUndo(undostart, attribs[attrib]);
                        if (rest === "off") {
                            attribs.recalc = rest; // manual recalc, not auto
                        } else { // all values other than "off" mean "on"
                            delete attribs.recalc;
                        }
                        break;

                    default:
                        errortext = `${scc.s_escUnknownSheetCmd}${cmdstr}`;
                        break;
                }
            }
            // Column attributes
            else if (/(^[A-Z])([A-Z])?(:[A-Z][A-Z]?){0,1}$/i.test(what)) {
                sheet.renderneeded = true;
                what = what.toUpperCase();
                pos = what.indexOf(":");
                if (pos >= 0) {
                    cr1 = SocialCalc.coordToCr(`${what.substring(0, pos)}1`);
                    cr2 = SocialCalc.coordToCr(`${what.substring(pos + 1)}1`);
                } else {
                    cr1 = SocialCalc.coordToCr(`${what}1`);
                    cr2 = cr1;
                }
                for (col = cr1.col; col <= cr2.col; col++) {
                    if (attrib === "width") {
                        cr = SocialCalc.rcColname(col);
                        if (saveundo) changes.AddUndo(`set ${cr} width`, sheet.colattribs.width[cr]);
                        if (rest.length > 0) {
                            sheet.colattribs.width[cr] = rest;
                        } else {
                            delete sheet.colattribs.width[cr];
                        }
                    }
                }
            }
            // !!!!! need row attribs !!!!
            // Cell attributes
            else if (/([a-z]){0,1}(\d+)/i.test(what)) {
                ParseRange();
                if (cr1.row !== cr2.row || cr1.col !== cr2.col || sheet.celldisplayneeded || sheet.renderneeded) {
                    // not one cell
                    sheet.renderneeded = true;
                    sheet.celldisplayneeded = "";
                } else {
                    sheet.celldisplayneeded = SocialCalc.crToCoord(cr1.col, cr1.row);
                }

                for (row = cr1.row; row <= cr2.row; row++) {
                    for (col = cr1.col; col <= cr2.col; col++) {
                        cr = SocialCalc.crToCoord(col, row);
                        cell = sheet.GetAssuredCell(cr);
                        if (saveundo) changes.AddUndo(`set ${cr} all`, sheet.CellToString(cell));

                        if (attrib === "value") { // set coord value type numeric-value
                            pos = rest.indexOf(" ");
                            cell.datavalue = rest.substring(pos + 1) - 0;
                            delete cell.errors;
                            cell.datatype = "v";
                            cell.valuetype = rest.substring(0, pos);
                            delete cell.displaystring;
                            delete cell.parseinfo;
                            attribs.needsrecalc = "yes";
                        }
                        else if (attrib === "text") { // set coord text type text-value
                            pos = rest.indexOf(" ");
                            cell.datavalue = SocialCalc.decodeFromSave(rest.substring(pos + 1));
                            delete cell.errors;
                            cell.datatype = "t";
                            cell.valuetype = rest.substring(0, pos);
                            delete cell.displaystring;
                            delete cell.parseinfo;
                            attribs.needsrecalc = "yes";
                        }
                        else if (attrib === "formula") { // set coord formula formula-body-less-initial-=
                            cell.datavalue = 0; // until recalc
                            delete cell.errors;
                            cell.datatype = "f";
                            cell.valuetype = "e#N/A"; // until recalc
                            cell.formula = rest;
                            delete cell.displaystring;
                            delete cell.parseinfo;
                            attribs.needsrecalc = "yes";
                        }
                        else if (attrib === "constant") { // set coord constant type numeric-value source-text
                            pos = rest.indexOf(" ");
                            pos2 = rest.substring(pos + 1).indexOf(" ");
                            cell.datavalue = rest.substring(pos + 1, pos + 1 + pos2) - 0;
                            cell.valuetype = rest.substring(0, pos);
                            if (cell.valuetype.charAt(0) === "e") { // error
                                cell.errors = cell.valuetype.substring(1);
                            } else {
                                delete cell.errors;
                            }
                            cell.datatype = "c";
                            cell.formula = rest.substring(pos + pos2 + 2);
                            delete cell.displaystring;
                            delete cell.parseinfo;
                            attribs.needsrecalc = "yes";
                        }
                        else if (attrib === "empty") { // erase value
                            cell.datavalue = "";
                            delete cell.errors;
                            cell.datatype = null;
                            cell.formula = "";
                            cell.valuetype = "b";
                            delete cell.displaystring;
                            delete cell.parseinfo;
                            attribs.needsrecalc = "yes";
                        }
                        else if (attrib === "all") { // set coord all :this:val1:that:val2...
                            if (rest.length > 0) {
                                cell = new SocialCalc.Cell(cr);
                                sheet.CellFromStringParts(cell, rest.split(":"), 1);
                                sheet.cells[cr] = cell;
                            } else {
                                delete sheet.cells[cr];
                            }
                            attribs.needsrecalc = "yes";
                        }
                        else if (/^b[trbl]$/.test(attrib)) { // set coord bt 1px solid black
                            cell[attrib] = sheet.GetStyleNum("borderstyle", rest);
                            sheet.renderneeded = true; // affects more than just one cell
                        }
                        else if (attrib === "color" || attrib === "bgcolor") {
                            cell[attrib] = sheet.GetStyleNum("color", rest);
                        }
                        else if (attrib === "layout" || attrib === "cellformat") {
                            cell[attrib] = sheet.GetStyleNum(attrib, rest);
                        }
                        else if (attrib === "font") { // set coord font style weight size family
                            if (rest === "* * *") rest = "";
                            cell[attrib] = sheet.GetStyleNum("font", rest);
                        }
                        else if (attrib === "textvalueformat" || attrib === "nontextvalueformat") {
                            cell[attrib] = sheet.GetStyleNum("valueformat", rest);
                            delete cell.displaystring;
                        }
                        else if (attrib === "cssc") {
                            rest = rest.replace(/[^a-zA-Z0-9\-]/g, "");
                            cell.cssc = rest;
                        }
                        else if (attrib === "csss") {
                            rest = rest.replace(/\n/g, "");
                            cell.csss = rest;
                        }
                        else if (attrib === "mod") {
                            rest = rest.replace(/[^yY]/g, "").toLowerCase();
                            cell.mod = rest;
                        }
                        else if (attrib === "comment") {
                            cell.comment = SocialCalc.decodeFromSave(rest);
                        }
                        else {
                            errortext = `${scc.s_escUnknownSetCoordCmd}${cmdstr}`;
                        }
                    }
                }
            }
            break;

        case "merge":
            sheet.renderneeded = true;
            what = cmd.NextToken();
            rest = cmd.RestOfString();
            ParseRange();
            cell = sheet.GetAssuredCell(cr1.coord);
            if (saveundo) changes.AddUndo(`unmerge ${cr1.coord}`);

            if (cr2.col > cr1.col) cell.colspan = cr2.col - cr1.col + 1;
            else delete cell.colspan;
            if (cr2.row > cr1.row) cell.rowspan = cr2.row - cr1.row + 1;
            else delete cell.rowspan;

            sheet.changedrendervalues = true;
            break;

        case "unmerge":
            sheet.renderneeded = true;
            what = cmd.NextToken();
            rest = cmd.RestOfString();
            ParseRange();
            cell = sheet.GetAssuredCell(cr1.coord);
            if (saveundo) changes.AddUndo(`merge ${cr1.coord}:${SocialCalc.crToCoord(cr1.col + (cell.colspan || 1) - 1, cr1.row + (cell.rowspan || 1) - 1)}`);

            delete cell.colspan;
            delete cell.rowspan;

            sheet.changedrendervalues = true;
            break;

        case "erase":
        case "cut":
            sheet.renderneeded = true;
            sheet.changedrendervalues = true;
            what = cmd.NextToken();
            rest = cmd.RestOfString();
            ParseRange();

            if (saveundo) changes.AddUndo("changedrendervalues"); // to take care of undone pasted spans
            if (cmd1 === "cut") { // save copy of whole thing before erasing
                if (saveundo) changes.AddUndo("loadclipboard", SocialCalc.encodeForSave(SocialCalc.Clipboard.clipboard));
                SocialCalc.Clipboard.clipboard = SocialCalc.CreateSheetSave(sheet, what);
            }

            for (row = cr1.row; row <= cr2.row; row++) {
                for (col = cr1.col; col <= cr2.col; col++) {
                    cr = SocialCalc.crToCoord(col, row);
                    cell = sheet.GetAssuredCell(cr);
                    if (saveundo) changes.AddUndo(`set ${cr} all`, sheet.CellToString(cell));
                    if (rest === "all") {
                        delete sheet.cells[cr];
                    }
                    else if (rest === "formulas") {
                        cell.datavalue = "";
                        cell.datatype = null;
                        cell.formula = "";
                        cell.valuetype = "b";
                        delete cell.errors;
                        delete cell.displaystring;
                        delete cell.parseinfo;
                        if (cell.comment) { // comments are considered content for erasing
                            delete cell.comment;
                        }
                    }
                    else if (rest === "formats") {
                        newcell = new SocialCalc.Cell(cr); // create a new cell without attributes
                        newcell.datavalue = cell.datavalue; // copy existing values
                        newcell.datatype = cell.datatype;
                        newcell.formula = cell.formula;
                        newcell.valuetype = cell.valuetype;
                        if (cell.comment) {
                            newcell.comment = cell.comment;
                        }
                        sheet.cells[cr] = newcell; // replace
                    }
                }
            }
            attribs.needsrecalc = "yes";
            break;

        case "fillright":
        case "filldown":
            sheet.renderneeded = true;
            sheet.changedrendervalues = true;
            if (saveundo) changes.AddUndo("changedrendervalues"); // to take care of undone pasted spans
            what = cmd.NextToken();
            rest = cmd.RestOfString();
            ParseRange();
            if (cmd1 === "fillright") {
                fillright = true;
                rowstart = cr1.row;
                colstart = cr1.col + 1;
            } else {
                fillright = false;
                rowstart = cr1.row + 1;
                colstart = cr1.col;
            }

            for (row = rowstart; row <= cr2.row; row++) {
                for (col = colstart; col <= cr2.col; col++) {
                    cr = SocialCalc.crToCoord(col, row);
                    cell = sheet.GetAssuredCell(cr);
                    if (saveundo) changes.AddUndo(`set ${cr} all`, sheet.CellToString(cell));
                    if (fillright) {
                        crbase = SocialCalc.crToCoord(cr1.col, row);
                        coloffset = col - colstart + 1;
                        rowoffset = 0;
                    } else {
                        crbase = SocialCalc.crToCoord(col, cr1.row);
                        coloffset = 0;
                        rowoffset = row - rowstart + 1;
                    }
                    basecell = sheet.GetAssuredCell(crbase);

                    if (rest === "all" || rest === "formats") {
                        for (attrib in cellProperties) {
                            if (cellProperties[attrib] === 1) continue; // copy only format attributes
                            if (typeof basecell[attrib] === "undefined" || cellProperties[attrib] === 3) {
                                delete cell[attrib];
                            } else {
                                cell[attrib] = basecell[attrib];
                            }
                        }
                    }

                    if (rest === "all" || rest === "formulas") {
                        cell.datavalue = basecell.datavalue;
                        cell.datatype = basecell.datatype;
                        cell.valuetype = basecell.valuetype;
                        if (cell.datatype === "f") { // offset relative coords, even in sheet references
                            cell.formula = SocialCalc.OffsetFormulaCoords(basecell.formula, coloffset, rowoffset);
                        } else {
                            cell.formula = basecell.formula;
                        }
                        delete cell.parseinfo;
                        cell.errors = basecell.errors;
                    }
                    delete cell.displaystring;
                }
            }

            attribs.needsrecalc = "yes";
            break;

        case "copy":
            what = cmd.NextToken();
            rest = cmd.RestOfString();
            if (saveundo) changes.AddUndo("loadclipboard", SocialCalc.encodeForSave(SocialCalc.Clipboard.clipboard));
            SocialCalc.Clipboard.clipboard = SocialCalc.CreateSheetSave(sheet, what);
            break;

        case "loadclipboard":
            rest = cmd.RestOfString();
            if (saveundo) changes.AddUndo("loadclipboard", SocialCalc.encodeForSave(SocialCalc.Clipboard.clipboard));
            SocialCalc.Clipboard.clipboard = SocialCalc.decodeFromSave(rest);
            break;

        case "clearclipboard":
            if (saveundo) changes.AddUndo("loadclipboard", SocialCalc.encodeForSave(SocialCalc.Clipboard.clipboard));
            SocialCalc.Clipboard.clipboard = "";
            break;

        case "paste":
            sheet.renderneeded = true;
            sheet.changedrendervalues = true;
            if (saveundo) changes.AddUndo("changedrendervalues"); // to take care of undone pasted spans
            what = cmd.NextToken();
            rest = cmd.RestOfString();
            ParseRange();
            if (!SocialCalc.Clipboard.clipboard) {
                break;
            }
            clipsheet = new SocialCalc.Sheet(); // load clipboard contents as another sheet
            clipsheet.ParseSheetSave(SocialCalc.Clipboard.clipboard);
            cliprange = SocialCalc.ParseRange(clipsheet.copiedfrom);
            coloffset = cr1.col - cliprange.cr1.col; // get sizes, etc.
            rowoffset = cr1.row - cliprange.cr1.row;
            numcols = cliprange.cr2.col - cliprange.cr1.col + 1;
            numrows = cliprange.cr2.row - cliprange.cr1.row + 1;
            if (cr1.col + numcols - 1 > attribs.lastcol) attribs.lastcol = cr1.col + numcols - 1;
            if (cr1.row + numrows - 1 > attribs.lastrow) attribs.lastrow = cr1.row + numrows - 1;

            for (row = cr1.row; row < cr1.row + numrows; row++) {
                for (col = cr1.col; col < cr1.col + numcols; col++) {
                    cr = SocialCalc.crToCoord(col, row);
                    cell = sheet.GetAssuredCell(cr);
                    if (saveundo) changes.AddUndo(`set ${cr} all`, sheet.CellToString(cell));
                    crbase = SocialCalc.crToCoord(col - coloffset, row - rowoffset);
                    basecell = clipsheet.GetAssuredCell(crbase);

                    if (rest === "all" || rest === "formats") {
                        for (attrib in cellProperties) {
                            if (cellProperties[attrib] === 1) continue; // copy only format attributes
                            if (typeof basecell[attrib] === "undefined" || cellProperties[attrib] === 3) {
                                delete cell[attrib];
                            } else {
                                attribtable = SocialCalc.CellPropertiesTable[attrib];
                                if (attribtable && basecell[attrib]) { // table indexes to expand to strings since other sheet may have diff indexes
                                    cell[attrib] = sheet.GetStyleNum(attribtable, clipsheet.GetStyleString(attribtable, basecell[attrib]));
                                } else { // these are not table indexes
                                    cell[attrib] = basecell[attrib];
                                }
                            }
                        }
                    }

                    if (rest === "all" || rest === "formulas") {
                        cell.datavalue = basecell.datavalue;
                        cell.datatype = basecell.datatype;
                        cell.valuetype = basecell.valuetype;
                        if (cell.datatype === "f") { // offset relative coords, even in sheet references
                            cell.formula = SocialCalc.OffsetFormulaCoords(basecell.formula, coloffset, rowoffset);
                        } else {
                            cell.formula = basecell.formula;
                        }
                        delete cell.parseinfo;
                        cell.errors = basecell.errors;
                        if (basecell.comment) { // comments are pasted as part of content, though not filled, etc.
                            cell.comment = basecell.comment;
                        } else if (cell.comment) {
                            delete cell.comment;
                        }
                    }
                    delete cell.displaystring;
                }
            }

            attribs.needsrecalc = "yes";
            break;

        case "sort": // sort cr1:cr2 col1 up/down col2 up/down col3 up/down
            sheet.renderneeded = true;
            sheet.changedrendervalues = true;
            if (saveundo) changes.AddUndo("changedrendervalues"); // to take care of undone pasted spans
            what = cmd.NextToken();
            ParseRange();
            cols = []; // get columns and sort directions (or "")
            dirs = [];
            lastsortcol = 0;
            for (i = 0; i <= 3; i++) {
                cols[i] = cmd.NextToken();
                dirs[i] = cmd.NextToken();
                if (cols[i]) lastsortcol = i;
            }

            sortcells = {}; // a copy of the data which will replace the original, but in the new order
            sortlist = []; // an array of 0, 1, ..., nrows-1 needed for sorting
            sortvalues = []; // values to be sorted corresponding to sortlist
            sorttypes = []; // basic types of the values

            for (row = cr1.row; row <= cr2.row; row++) { // fill in the sort info
                for (col = cr1.col; col <= cr2.col; col++) {
                    cr = SocialCalc.crToCoord(col, row);
                    cell = sheet.cells[cr];
                    if (cell) { // only copy non-empty cells
                        sortcells[cr] = sheet.CellToString(cell);
                        if (saveundo) changes.AddUndo(`set ${cr} all`, sortcells[cr]);
                    } else {
                        if (saveundo) changes.AddUndo(`set ${cr} all`);
                    }
                }
                sortlist.push(sortlist.length);
                sortvalues.push([]);
                sorttypes.push([]);
                const slast = sorttypes.length - 1;
                for (i = 0; i <= lastsortcol; i++) {
                    cr = cols[i] + row; // get cr on this row in sort col
                    cell = sheet.GetAssuredCell(cr);
                    val = cell.datavalue;
                    valtype = cell.valuetype.charAt(0) || "b";
                    if (valtype === "t") val = val.toLowerCase();
                    sortvalues[slast].push(val);
                    sorttypes[slast].push(valtype);
                }
            }

            /**
             * @description Comparison function that handles all type variations for sorting
             * @param {number} a - First sort index
             * @param {number} b - Second sort index
             * @returns {number} -1, 0, or 1 for sort comparison
             */
            sortfunction = function (a, b) {
                let i, a1, b1, ta, tb, cresult;
                for (i = 0; i <= lastsortcol; i++) {
                    if (dirs[i] === "up") { // handle sort direction
                        a1 = a;
                        b1 = b;
                    } else {
                        a1 = b;
                        b1 = a;
                    }
                    ta = sorttypes[a1][i];
                    tb = sorttypes[b1][i];

                    if (ta === "t") { // numbers < text < errors, blank always last no matter what dir
                        if (tb === "t") {
                            a1 = sortvalues[a1][i];
                            b1 = sortvalues[b1][i];
                            cresult = a1 > b1 ? 1 : (a1 < b1 ? -1 : 0);
                        } else if (tb === "n") {
                            cresult = 1;
                        } else if (tb === "b") {
                            cresult = dirs[i] === "up" ? -1 : 1;
                        } else if (tb === "e") {
                            cresult = -1;
                        }
                    } else if (ta === "n") {
                        if (tb === "t") {
                            cresult = -1;
                        } else if (tb === "n") {
                            a1 = sortvalues[a1][i] - 0; // force to numeric, just in case
                            b1 = sortvalues[b1][i] - 0;
                            cresult = a1 > b1 ? 1 : (a1 < b1 ? -1 : 0);
                        } else if (tb === "b") {
                            cresult = dirs[i] === "up" ? -1 : 1;
                        } else if (tb === "e") {
                            cresult = -1;
                        }
                    } else if (ta === "e") {
                        if (tb === "e") {
                            a1 = sortvalues[a1][i];
                            b1 = sortvalues[b1][i];
                            cresult = a1 > b1 ? 1 : (a1 < b1 ? -1 : 0);
                        } else if (tb === "b") {
                            cresult = dirs[i] === "up" ? -1 : 1;
                        } else {
                            cresult = 1;
                        }
                    } else if (ta === "b") {
                        if (tb === "b") {
                            cresult = 0;
                        } else {
                            cresult = dirs[i] === "up" ? 1 : -1;
                        }
                    }

                    if (cresult) { // return if tested not equal, otherwise do next column
                        return cresult;
                    }
                }
                cresult = a > b ? 1 : (a < b ? -1 : 0); // equal - return position in original to maintain it
                return cresult;
            };

            sortlist.sort(sortfunction);

            for (row = cr1.row; row <= cr2.row; row++) { // copy original rows into sorted positions
                originalrow = sortlist[row - cr1.row]; // relative position where it was in original
                for (col = cr1.col; col <= cr2.col; col++) {
                    cr = SocialCalc.crToCoord(col, row);
                    sortedcr = SocialCalc.crToCoord(col, originalrow + cr1.row); // original cell to be put in new place
                    if (sortcells[sortedcr]) {
                        cell = new SocialCalc.Cell(cr);
                        sheet.CellFromStringParts(cell, sortcells[sortedcr].split(":"), 1);
                        if (cell.datatype === "f") { // offset coord refs, even to ***relative*** coords in other sheets
                            cell.formula = SocialCalc.OffsetFormulaCoords(cell.formula, 0, (row - cr1.row) - originalrow);
                        }
                        sheet.cells[cr] = cell;
                    } else {
                        delete sheet.cells[cr];
                    }
                }
            }

            attribs.needsrecalc = "yes";
            break;
        case "insertcol":
        case "insertrow":
            sheet.renderneeded = true;
            sheet.changedrendervalues = true;
            what = cmd.NextToken();
            rest = cmd.RestOfString();
            ParseRange();

            if (cmd1 === "insertcol") {
                coloffset = 1;
                colend = cr1.col;
                rowoffset = 0;
                rowend = 1;
                newcolstart = cr1.col;
                newcolend = cr1.col;
                newrowstart = 1;
                newrowend = attribs.lastrow;
                if (saveundo) changes.AddUndo(`deletecol ${cr1.coord}`);
            } else {
                coloffset = 0;
                colend = 1;
                rowoffset = 1;
                rowend = cr1.row;
                newcolstart = 1;
                newcolend = attribs.lastcol;
                newrowstart = cr1.row;
                newrowend = cr1.row;
                if (saveundo) changes.AddUndo(`deleterow ${cr1.coord}`);
            }

            // Copy the cells forward
            for (row = attribs.lastrow; row >= rowend; row--) {
                for (col = attribs.lastcol; col >= colend; col--) {
                    crbase = SocialCalc.crToCoord(col, row);
                    cr = SocialCalc.crToCoord(col + coloffset, row + rowoffset);
                    if (!sheet.cells[crbase]) { // copying empty cell
                        delete sheet.cells[cr]; // delete anything that may have been there
                    } else { // overwrite existing cell with moved contents
                        sheet.cells[cr] = sheet.cells[crbase];
                    }
                }
            }

            // Fill the "new" empty cells
            for (row = newrowstart; row <= newrowend; row++) {
                for (col = newcolstart; col <= newcolend; col++) {
                    cr = SocialCalc.crToCoord(col, row);
                    cell = new SocialCalc.Cell(cr);
                    sheet.cells[cr] = cell;
                    crbase = SocialCalc.crToCoord(col - coloffset, row - rowoffset); // copy attribs of the one before (0 gives you A or 1)
                    basecell = sheet.GetAssuredCell(crbase);
                    for (attrib in cellProperties) {
                        if (cellProperties[attrib] === 2) { // copy only format attributes
                            cell[attrib] = basecell[attrib];
                        }
                    }
                }
            }

            // Update cell references to moved cells in calculated formulas
            for (cr in sheet.cells) {
                cell = sheet.cells[cr];
                if (cell && cell.datatype === "f") {
                    cell.formula = SocialCalc.AdjustFormulaCoords(cell.formula, cr1.col, coloffset, cr1.row, rowoffset);
                }
                if (cell) {
                    delete cell.parseinfo;
                }
            }

            // Update cell references to moved cells in names
            for (name in sheet.names) {
                if (sheet.names[name]) { // works with "A1", "A1:A20", and "=formula" forms
                    v1 = sheet.names[name].definition;
                    v2 = "";
                    if (v1.charAt(0) === "=") {
                        v2 = "=";
                        v1 = v1.substring(1);
                    }
                    sheet.names[name].definition = v2 +
                        SocialCalc.AdjustFormulaCoords(v1, cr1.col, coloffset, cr1.row, rowoffset);
                }
            }

            // Copy the row attributes forward
            for (row = attribs.lastrow; row >= rowend && cmd1 === "insertrow"; row--) {
                rownext = row + rowoffset;
                for (attrib in sheet.rowattribs) {
                    val = sheet.rowattribs[attrib][row];
                    if (sheet.rowattribs[attrib][rownext] !== val) { // make assignment only if different
                        if (val) {
                            sheet.rowattribs[attrib][rownext] = val;
                        } else {
                            delete sheet.rowattribs[attrib][rownext];
                        }
                    }
                }
            }

            // Copy the column attributes forward
            for (col = attribs.lastcol; col >= colend && cmd1 === "insertcol"; col--) {
                colthis = SocialCalc.rcColname(col);
                colnext = SocialCalc.rcColname(col + coloffset);
                for (attrib in sheet.colattribs) {
                    val = sheet.colattribs[attrib][colthis];
                    if (sheet.colattribs[attrib][colnext] !== val) { // make assignment only if different
                        if (val) {
                            sheet.colattribs[attrib][colnext] = val;
                        } else {
                            delete sheet.colattribs[attrib][colnext];
                        }
                    }
                }
            }

            attribs.lastcol += coloffset;
            attribs.lastrow += rowoffset;
            attribs.needsrecalc = "yes";
            break;

        case "deletecol":
        case "deleterow":
            sheet.renderneeded = true;
            sheet.changedrendervalues = true;
            what = cmd.NextToken();
            rest = cmd.RestOfString();
            lastcol = attribs.lastcol; // save old values since ParseRange sets...
            lastrow = attribs.lastrow;
            ParseRange();

            if (cmd1 === "deletecol") {
                coloffset = cr1.col - cr2.col - 1;
                rowoffset = 0;
                colstart = cr2.col + 1;
                rowstart = 1;
            } else {
                coloffset = 0;
                rowoffset = cr1.row - cr2.row - 1;
                colstart = 1;
                rowstart = cr2.row + 1;
            }

            // Copy the cells backwards - extra so no dup of last set
            for (row = rowstart; row <= lastrow - rowoffset; row++) {
                for (col = colstart; col <= lastcol - coloffset; col++) {
                    cr = SocialCalc.crToCoord(col + coloffset, row + rowoffset);
                    if (saveundo && (row < rowstart - rowoffset || col < colstart - coloffset)) { // save cells that are overwritten as undo info
                        cell = sheet.cells[cr];
                        if (!cell) { // empty cell
                            changes.AddUndo(`erase ${cr} all`);
                        } else {
                            changes.AddUndo(`set ${cr} all`, sheet.CellToString(cell));
                        }
                    }
                    crbase = SocialCalc.crToCoord(col, row);
                    cell = sheet.cells[crbase];
                    if (!cell) { // copying empty cell
                        delete sheet.cells[cr]; // delete anything that may have been there
                    } else { // overwrite existing cell with moved contents
                        sheet.cells[cr] = cell;
                    }
                }
            }

            //!!! multiple deletes isn't setting #REF!; need to fix up #REF!'s on undo but only those!

            // Update cell references to moved cells in calculated formulas
            for (cr in sheet.cells) {
                cell = sheet.cells[cr];
                if (cell) {
                    if (cell.datatype === "f") {
                        oldformula = cell.formula;
                        cell.formula = SocialCalc.AdjustFormulaCoords(oldformula, cr1.col, coloffset, cr1.row, rowoffset);
                        if (cell.formula !== oldformula) {
                            delete cell.parseinfo;
                            if (saveundo && cell.formula.indexOf("#REF!") !== -1) { // save old version only if removed coord
                                oldcr = SocialCalc.coordToCr(cr);
                                changes.AddUndo(`set ${SocialCalc.rcColname(oldcr.col - coloffset)}${(oldcr.row - rowoffset)} formula ${oldformula}`);
                            }
                        }
                    } else {
                        delete cell.parseinfo;
                    }
                }
            }

            // Update cell references to moved cells in names
            for (name in sheet.names) {
                if (sheet.names[name]) { // works with "A1", "A1:A20", and "=formula" forms
                    v1 = sheet.names[name].definition;
                    v2 = "";
                    if (v1.charAt(0) === "=") {
                        v2 = "=";
                        v1 = v1.substring(1);
                    }
                    sheet.names[name].definition = v2 +
                        SocialCalc.AdjustFormulaCoords(v1, cr1.col, coloffset, cr1.row, rowoffset);
                }
            }

            // Copy the row attributes backwards
            for (row = rowstart; row <= lastrow - rowoffset && cmd1 === "deleterow"; row++) {
                rowbefore = row + rowoffset;
                for (attrib in sheet.rowattribs) {
                    val = sheet.rowattribs[attrib][row];
                    if (sheet.rowattribs[attrib][rowbefore] !== val) { // make assignment only if different
                        if (saveundo) changes.AddUndo(`set ${rowbefore} ${attrib}`, sheet.rowattribs[attrib][rowbefore]);
                        if (val) {
                            sheet.rowattribs[attrib][rowbefore] = val;
                        } else {
                            delete sheet.rowattribs[attrib][rowbefore];
                        }
                    }
                }
            }

            // Copy the column attributes backwards
            for (col = colstart; col <= lastcol - coloffset && cmd1 === "deletecol"; col++) {
                colthis = SocialCalc.rcColname(col);
                colbefore = SocialCalc.rcColname(col + coloffset);
                for (attrib in sheet.colattribs) {
                    val = sheet.colattribs[attrib][colthis];
                    if (sheet.colattribs[attrib][colbefore] !== val) { // make assignment only if different
                        if (saveundo) changes.AddUndo(`set ${colbefore} ${attrib}`, sheet.colattribs[attrib][colbefore]);
                        if (val) {
                            sheet.colattribs[attrib][colbefore] = val;
                        } else {
                            delete sheet.colattribs[attrib][colbefore];
                        }
                    }
                }
            }

            if (saveundo) {
                if (cmd1 === "deletecol") {
                    for (col = cr1.col; col <= cr2.col; col++) {
                        changes.AddUndo(`insertcol ${SocialCalc.rcColname(col)}`);
                    }
                } else {
                    for (row = cr1.row; row <= cr2.row; row++) {
                        changes.AddUndo(`insertrow ${row}`);
                    }
                }
            }

            if (cmd1 === "deletecol") {
                if (cr1.col <= lastcol) { // shrink sheet unless deleted phantom cols off the end
                    if (cr2.col <= lastcol) {
                        attribs.lastcol += coloffset;
                    } else {
                        attribs.lastcol = cr1.col - 1;
                    }
                }
            } else {
                if (cr1.row <= lastrow) { // shrink sheet unless deleted phantom rows off the end
                    if (cr2.row <= lastrow) {
                        attribs.lastrow += rowoffset;
                    } else {
                        attribs.lastrow = cr1.row - 1;
                    }
                }
            }
            attribs.needsrecalc = "yes";
            break;

        case "movepaste":
        case "moveinsert":
            let movingcells, dest, destcr, inserthoriz, insertvert, pushamount, movedto;

            sheet.renderneeded = true;
            sheet.changedrendervalues = true;
            if (saveundo) changes.AddUndo("changedrendervalues"); // to take care of undone pasted spans
            what = cmd.NextToken();
            dest = cmd.NextToken();
            rest = cmd.RestOfString(); // rest is all/formulas/formats
            if (rest === "") rest = "all";

            ParseRange();

            destcr = SocialCalc.coordToCr(dest);

            coloffset = destcr.col - cr1.col;
            rowoffset = destcr.row - cr1.row;
            numcols = cr2.col - cr1.col + 1;
            numrows = cr2.row - cr1.row + 1;

            // Get a copy of moving cells and erase from where they were
            movingcells = {};

            for (row = cr1.row; row <= cr2.row; row++) {
                for (col = cr1.col; col <= cr2.col; col++) {
                    cr = SocialCalc.crToCoord(col, row);
                    cell = sheet.GetAssuredCell(cr);
                    if (saveundo) changes.AddUndo(`set ${cr} all`, sheet.CellToString(cell));

                    if (!sheet.cells[cr]) { // if had nothing
                        continue; // don't save anything
                    }
                    movingcells[cr] = new SocialCalc.Cell(cr); // create new cell to copy

                    for (attrib in cellProperties) { // go through each property
                        if (typeof cell[attrib] === "undefined") { // don't copy undefined things and no need to delete
                            continue;
                        } else {
                            movingcells[cr][attrib] = cell[attrib]; // copy for potential moving
                        }
                        if (rest === "all") {
                            delete cell[attrib];
                        }
                        if (rest === "formulas") {
                            if (cellProperties[attrib] === 1 || cellProperties[attrib] === 3) {
                                delete cell[attrib];
                            }
                        }
                        if (rest === "formats") {
                            if (cellProperties[attrib] === 2) {
                                delete cell[attrib];
                            }
                        }
                    }
                    if (rest === "formulas") { // leave pristine deleted cell
                        cell.datavalue = "";
                        cell.datatype = null;
                        cell.formula = "";
                        cell.valuetype = "b";
                    }
                    if (rest === "all") { // leave nothing for move all
                        delete sheet.cells[cr];
                    }
                }
            }

            // If moveinsert, check destination OK, and calculate pushing parameters
            if (cmd1 === "moveinsert") {
                inserthoriz = false;
                insertvert = false;
                if (rowoffset === 0 && (destcr.col < cr1.col || destcr.col > cr2.col)) {
                    if (destcr.col < cr1.col) { // moving left
                        pushamount = cr1.col - destcr.col;
                        inserthoriz = -1;
                    } else {
                        destcr.col -= 1;
                        coloffset = destcr.col - cr2.col;
                        pushamount = destcr.col - cr2.col;
                        inserthoriz = 1;
                    }
                } else if (coloffset === 0 && (destcr.row < cr1.row || destcr.row > cr2.row)) {
                    if (destcr.row < cr1.row) { // moving up
                        pushamount = cr1.row - destcr.row;
                        insertvert = -1;
                    } else {
                        destcr.row -= 1;
                        rowoffset = destcr.row - cr2.row;
                        pushamount = destcr.row - cr2.row;
                        insertvert = 1;
                    }
                } else {
                    cmd1 = "movepaste"; // not allowed right now - ignore
                }
            }

            // Push any cells that need pushing
            movedto = {}; // remember what was moved where

            if (insertvert) {
                for (row = 0; row < pushamount; row++) {
                    for (col = cr1.col; col <= cr2.col; col++) {
                        if (insertvert < 0) {
                            crbase = SocialCalc.crToCoord(col, destcr.row + pushamount - row - 1); // from cell
                            cr = SocialCalc.crToCoord(col, cr2.row - row); // to cell
                        } else {
                            crbase = SocialCalc.crToCoord(col, destcr.row - pushamount + row + 1); // from cell
                            cr = SocialCalc.crToCoord(col, cr1.row + row); // to cell
                        }

                        basecell = sheet.GetAssuredCell(crbase);
                        if (saveundo) changes.AddUndo(`set ${crbase} all`, sheet.CellToString(basecell));

                        cell = sheet.GetAssuredCell(cr);
                        if (rest === "all" || rest === "formats") {
                            for (attrib in cellProperties) {
                                if (cellProperties[attrib] === 1) continue; // copy only format attributes
                                if (typeof basecell[attrib] === "undefined" || cellProperties[attrib] === 3) {
                                    delete cell[attrib];
                                } else {
                                    cell[attrib] = basecell[attrib];
                                }
                            }
                        }
                        if (rest === "all" || rest === "formulas") {
                            cell.datavalue = basecell.datavalue;
                            cell.datatype = basecell.datatype;
                            cell.valuetype = basecell.valuetype;
                            cell.formula = basecell.formula;
                            delete cell.parseinfo;
                            cell.errors = basecell.errors;
                        }
                        delete cell.displaystring;

                        movedto[crbase] = cr; // old crbase is now at cr
                    }
                }
            }
            if (inserthoriz) {
                for (col = 0; col < pushamount; col++) {
                    for (row = cr1.row; row <= cr2.row; row++) {
                        if (inserthoriz < 0) {
                            crbase = SocialCalc.crToCoord(destcr.col + pushamount - col - 1, row);
                            cr = SocialCalc.crToCoord(cr2.col - col, row);
                        } else {
                            crbase = SocialCalc.crToCoord(destcr.col - pushamount + col + 1, row);
                            cr = SocialCalc.crToCoord(cr1.col + col, row);
                        }

                        basecell = sheet.GetAssuredCell(crbase);
                        if (saveundo) changes.AddUndo(`set ${crbase} all`, sheet.CellToString(basecell));

                        cell = sheet.GetAssuredCell(cr);
                        if (rest === "all" || rest === "formats") {
                            for (attrib in cellProperties) {
                                if (cellProperties[attrib] === 1) continue; // copy only format attributes
                                if (typeof basecell[attrib] === "undefined" || cellProperties[attrib] === 3) {
                                    delete cell[attrib];
                                } else {
                                    cell[attrib] = basecell[attrib];
                                }
                            }
                        }
                        if (rest === "all" || rest === "formulas") {
                            cell.datavalue = basecell.datavalue;
                            cell.datatype = basecell.datatype;
                            cell.valuetype = basecell.valuetype;
                            cell.formula = basecell.formula;
                            delete cell.parseinfo;
                            cell.errors = basecell.errors;
                        }
                        delete cell.displaystring;

                        movedto[crbase] = cr; // old crbase is now at cr
                    }
                }
            }

            // Paste moved cells into new place
            if (destcr.col + numcols - 1 > attribs.lastcol) attribs.lastcol = destcr.col + numcols - 1;
            if (destcr.row + numrows - 1 > attribs.lastrow) attribs.lastrow = destcr.row + numrows - 1;

            for (row = cr1.row; row < cr1.row + numrows; row++) {
                for (col = cr1.col; col < cr1.col + numcols; col++) {
                    cr = SocialCalc.crToCoord(col + coloffset, row + rowoffset);
                    cell = sheet.GetAssuredCell(cr);
                    if (saveundo) changes.AddUndo(`set ${cr} all`, sheet.CellToString(cell));

                    crbase = SocialCalc.crToCoord(col, row); // get old cell to move

                    movedto[crbase] = cr; // old crbase (moved cell) will now be at cr (destination)

                    if (rest === "all" && !movingcells[crbase]) { // moving an empty cell
                        delete sheet.cells[cr]; // make the cell empty
                        continue;
                    }

                    basecell = movingcells[crbase];
                    if (!basecell) basecell = sheet.GetAssuredCell(crbase);

                    if (rest === "all" || rest === "formats") {
                        for (attrib in cellProperties) {
                            if (cellProperties[attrib] === 1) continue; // copy only format attributes
                            if (typeof basecell[attrib] === "undefined" || cellProperties[attrib] === 3) {
                                delete cell[attrib];
                            } else {
                                cell[attrib] = basecell[attrib];
                            }
                        }
                    }
                    if (rest === "all" || rest === "formulas") {
                        cell.datavalue = basecell.datavalue;
                        cell.datatype = basecell.datatype;
                        cell.valuetype = basecell.valuetype;
                        cell.formula = basecell.formula;
                        delete cell.parseinfo;
                        cell.errors = basecell.errors;
                        if (basecell.comment) { // comments are pasted as part of content, though not filled, etc.
                            cell.comment = basecell.comment;
                        } else if (cell.comment) {
                            delete cell.comment;
                        }
                    }
                    delete cell.displaystring;
                }
            }

            // Do fixups
            // Update cell references to moved cells in calculated formulas
            for (cr in sheet.cells) {
                cell = sheet.cells[cr];
                if (cell) {
                    if (cell.datatype === "f") {
                        oldformula = cell.formula;
                        cell.formula = SocialCalc.ReplaceFormulaCoords(oldformula, movedto);
                        if (cell.formula !== oldformula) {
                            delete cell.parseinfo;
                            if (saveundo && !movedto[cr]) { // moved cells are already saved for undo
                                changes.AddUndo(`set ${cr} formula ${oldformula}`);
                            }
                        }
                    } else {
                        delete cell.parseinfo;
                    }
                }
            }

            // Update cell references to moved cells in names
            for (name in sheet.names) {
                if (sheet.names[name]) { // works with "A1", "A1:A20", and "=formula" forms
                    v1 = sheet.names[name].definition;
                    oldformula = v1;
                    v2 = "";
                    if (v1.charAt(0) === "=") {
                        v2 = "=";
                        v1 = v1.substring(1);
                    }
                    sheet.names[name].definition = v2 +
                        SocialCalc.ReplaceFormulaCoords(v1, movedto);
                    if (saveundo && sheet.names[name].definition !== oldformula) { // save changes
                        changes.AddUndo(`name define ${name} ${oldformula}`);
                    }
                }
            }

            attribs.needsrecalc = "yes";
            break;

        case "name":
            what = cmd.NextToken();
            name = cmd.NextToken();
            rest = cmd.RestOfString();

            name = name.toUpperCase().replace(/[^A-Z0-9_\.]/g, "");
            if (name === "") break; // must have something

            if (what === "define") {
                if (rest === "") break; // must have something
                if (sheet.names[name]) { // already exists
                    if (saveundo) changes.AddUndo(`name define ${name} ${sheet.names[name].definition}`);
                    sheet.names[name].definition = rest;
                } else { // new
                    if (saveundo) changes.AddUndo(`name delete ${name}`);
                    sheet.names[name] = { definition: rest, desc: "" };
                }
            } else if (what === "desc") {
                if (sheet.names[name]) { // must already exist
                    if (saveundo) changes.AddUndo(`name desc ${name} ${sheet.names[name].desc}`);
                    sheet.names[name].desc = rest;
                }
            } else if (what === "delete") {
                if (saveundo) {
                    if (sheet.names[name].desc) changes.AddUndo(`name desc ${name} ${sheet.names[name].desc}`);
                    changes.AddUndo(`name define ${name} ${sheet.names[name].definition}`);
                }
                delete sheet.names[name];
            }
            attribs.needsrecalc = "yes";
            break;

        case "recalc":
            attribs.needsrecalc = "yes"; // request recalc
            sheet.recalconce = true; // even if turned off
            break;

        case "redisplay":
            sheet.renderneeded = true;
            break;

        case "changedrendervalues": // needed for undo sometimes
            sheet.changedrendervalues = true;
            break;

        case "startcmdextension": // startcmdextension extension rest-of-command
            name = cmd.NextToken();
            cmdextension = SocialCalc.SheetCommandInfo.CmdExtensionCallbacks[name];
            if (cmdextension) {
                cmdextension.func(name, cmdextension.data, sheet, cmd, saveundo);
            }
            break;

        default:
            errortext = `${scc.s_escUnknownCmd}${cmdstr}`;
            break;
    }

    /* For Debugging:
    var ustack="";
    for (var i=0;i<sheet.changes.stack.length;i++) {
       ustack+=(i-0)+":"+sheet.changes.stack[i].command[0]+" of "+sheet.changes.stack[i].command.length+"/"+sheet.changes.stack[i].undo[0]+" of "+sheet.changes.stack[i].undo.length+",";
       }
    alert(cmdstr+"|"+sheet.changes.stack.length+"--"+ustack);
    */

    return errortext;
};

/**
 * @function SocialCalc.SheetUndo
 * @description Performs an undo operation on the sheet by executing stored undo commands
 * @param {SocialCalc.Sheet} sheet - The sheet to perform undo on
 * 
 * @example
 * // Undo the last operation
 * SocialCalc.SheetUndo(mySheet);
 */
SocialCalc.SheetUndo = function (sheet) {
    const tos = sheet.changes.TOS();
    const lastone = tos ? tos.undo.length - 1 : -1;
    let cmdstr = "";

    for (let i = lastone; i >= 0; i--) { // do them backwards
        if (cmdstr) cmdstr += "\n"; // concatenate with separate lines
        cmdstr += tos.undo[i];
    }
    sheet.changes.Undo();
    sheet.ScheduleSheetCommands(cmdstr, false); // do undo operations
};

/**
 * @function SocialCalc.SheetRedo
 * @description Performs a redo operation on the sheet by re-executing stored commands
 * @param {SocialCalc.Sheet} sheet - The sheet to perform redo on
 * 
 * @example
 * // Redo the last undone operation
 * SocialCalc.SheetRedo(mySheet);
 */
SocialCalc.SheetRedo = function (sheet) {
    const didredo = sheet.changes.Redo();
    if (!didredo) {
        sheet.ScheduleSheetCommands("", false); // schedule doing nothing
        return;
    }
    const tos = sheet.changes.TOS();
    let cmdstr = "";

    for (let i = 0; tos && i < tos.command.length; i++) {
        if (cmdstr) cmdstr += "\n"; // concatenate with separate lines
        cmdstr += tos.command[i];
    }
    sheet.ScheduleSheetCommands(cmdstr, false); // do redo operations
};

/**
 * @function SocialCalc.CreateAuditString
 * @description Creates an audit string containing all commands executed on the sheet
 * @param {SocialCalc.Sheet} sheet - The sheet to create audit string for
 * @returns {string} String containing all executed commands, one per line
 * 
 * @example
 * // Get audit trail of all commands
 * const auditTrail = SocialCalc.CreateAuditString(mySheet);
 * console.log("Commands executed:", auditTrail);
 */
SocialCalc.CreateAuditString = function (sheet) {
    let result = "";
    const stack = sheet.changes.stack;
    const tos = sheet.changes.tos;

    for (let i = 0; i <= tos; i++) {
        for (let j = 0; j < stack[i].command.length; j++) {
            result += `${stack[i].command[j]}\n`;
        }
    }

    return result;
};

/**
 * @function SocialCalc.GetStyleNum
 * @description Gets the style number for a given style string, creating it if it doesn't exist
 * @param {SocialCalc.Sheet} sheet - The sheet object
 * @param {string} atype - The attribute type (font, color, layout, etc.)
 * @param {string} style - The style definition string
 * @returns {number} The style number (index) for the given style
 * 
 * @example
 * // Get or create a font style number
 * const fontNum = SocialCalc.GetStyleNum(sheet, "font", "bold 12pt Arial");
 * 
 * // Get or create a color style number
 * const colorNum = SocialCalc.GetStyleNum(sheet, "color", "rgb(255,0,0)");
 */
SocialCalc.GetStyleNum = function (sheet, atype, style) {
    if (style.length === 0) return 0; // null means use zero, which means default or global default

    let num = sheet[`${atype}hash`][style];
    if (!num) {
        if (sheet[`${atype}s`].length < 1) sheet[`${atype}s`].push("");
        num = sheet[`${atype}s`].push(style) - 1;
        sheet[`${atype}hash`][style] = num;
        sheet.changedrendervalues = true;
    }
    return num;
};

/**
 * @function SocialCalc.GetStyleString
 * @description Gets the style string for a given style number
 * @param {SocialCalc.Sheet} sheet - The sheet object
 * @param {string} atype - The attribute type (font, color, layout, etc.)
 * @param {number} num - The style number (index)
 * @returns {string|null} The style definition string, or null if num is falsy
 * 
 * @example
 * // Get font style string from number
 * const fontStyle = SocialCalc.GetStyleString(sheet, "font", 2);
 * // Returns something like "bold 12pt Arial"
 * 
 * // Get color style string from number
 * const colorStyle = SocialCalc.GetStyleString(sheet, "color", 1);
 * // Returns something like "rgb(255,0,0)"
 */
SocialCalc.GetStyleString = function (sheet, atype, num) {
    if (!num) return null; // zero, null, and undefined return null
    return sheet[`${atype}s`][num];
};
/**
 * @function SocialCalc.OffsetFormulaCoords
 * @description Changes relative cell references by offsets (even those to other worksheets so fill, paste, sort work as expected)
 * @param {string} formula - The formula to modify
 * @param {number} coloffset - Number of columns to offset references
 * @param {number} rowoffset - Number of rows to offset references
 * @returns {string} Updated formula with offset coordinates
 * 
 * @description If you don't want references to change, use absolute references ($A$1).
 * This function handles relative references and adjusts them by the specified offsets.
 * 
 * @example
 * // Offset formula references by 2 columns and 1 row
 * const newFormula = SocialCalc.OffsetFormulaCoords("A1+B2", 2, 1);
 * // Returns: "C2+D3"
 * 
 * // Absolute references remain unchanged
 * const newFormula2 = SocialCalc.OffsetFormulaCoords("$A$1+B2", 2, 1);
 * // Returns: "$A$1+D3"
 */
SocialCalc.OffsetFormulaCoords = function (formula, coloffset, rowoffset) {
    let parseinfo, ttext, ttype, i, cr, newcr;
    let updatedformula = "";

    const scf = SocialCalc.Formula;
    if (!scf) {
        return "Need SocialCalc.Formula";
    }

    const tokentype = scf.TokenType;
    const token_op = tokentype.op;
    const token_string = tokentype.string;
    const token_coord = tokentype.coord;
    const tokenOpExpansion = scf.TokenOpExpansion;

    parseinfo = scf.ParseFormulaIntoTokens(formula);

    for (i = 0; i < parseinfo.length; i++) {
        ttype = parseinfo[i].type;
        ttext = parseinfo[i].text;

        if (ttype === token_coord) {
            newcr = "";
            cr = SocialCalc.coordToCr(ttext);

            // Add col offset unless absolute column
            if (ttext.charAt(0) !== "$") {
                cr.col += coloffset;
            } else {
                newcr += "$";
            }
            newcr += SocialCalc.rcColname(cr.col);

            // Add row offset unless absolute row
            if (ttext.indexOf("$", 1) === -1) {
                cr.row += rowoffset;
            } else {
                newcr += "$";
            }
            newcr += cr.row;

            if (cr.row < 1 || cr.col < 1) {
                newcr = "#REF!";
            }
            updatedformula += newcr;
        }
        else if (ttype === token_string) {
            if (ttext.indexOf('"') >= 0) { // quotes to double
                updatedformula += `"${ttext.replace(/"/, '""')}"`;
            } else {
                updatedformula += `"${ttext}"`;
            }
        }
        else if (ttype === token_op) {
            updatedformula += tokenOpExpansion[ttext] || ttext; // make sure short tokens (e.g., "G") go back full (">=")
        }
        else { // leave everything else alone
            updatedformula += ttext;
        }
    }

    return updatedformula;
};

/**
 * @function SocialCalc.AdjustFormulaCoords
 * @description Changes all cell references to cells starting with col/row by offsets
 * @param {string} formula - The formula to modify
 * @param {number} col - Starting column for adjustment
 * @param {number} coloffset - Column offset to apply
 * @param {number} row - Starting row for adjustment
 * @param {number} rowoffset - Row offset to apply
 * @returns {string} Updated formula with adjusted coordinates
 * 
 * @description This function is typically used when inserting/deleting rows or columns
 * to update formulas that reference the affected areas.
 * 
 * @example
 * // Adjust references when inserting 2 columns at column C (3)
 * const newFormula = SocialCalc.AdjustFormulaCoords("A1+D1+E1", 3, 2, 1, 0);
 * // A1 unchanged, D1 becomes F1, E1 becomes G1
 */
SocialCalc.AdjustFormulaCoords = function (formula, col, coloffset, row, rowoffset) {
    let ttype, ttext, i, newcr, parseinfo;
    let updatedformula = "";
    let sheetref = false;

    const scf = SocialCalc.Formula;
    if (!scf) {
        return "Need SocialCalc.Formula";
    }

    const tokentype = scf.TokenType;
    const token_op = tokentype.op;
    const token_string = tokentype.string;
    const token_coord = tokentype.coord;
    const tokenOpExpansion = scf.TokenOpExpansion;

    parseinfo = SocialCalc.Formula.ParseFormulaIntoTokens(formula);

    for (i = 0; i < parseinfo.length; i++) {
        ttype = parseinfo[i].type;
        ttext = parseinfo[i].text;

        if (ttype === token_op) { // references with sheet specifier are not offset
            if (ttext === "!") {
                sheetref = true; // found a sheet reference
            } else if (ttext !== ":") { // for everything but a range, reset
                sheetref = false;
            }
            ttext = tokenOpExpansion[ttext] || ttext; // make sure short tokens (e.g., "G") go back full (">=")
        }

        if (ttype === token_coord) {
            const cr = SocialCalc.coordToCr(ttext);

            // Check if references to deleted cells become invalid
            if ((coloffset < 0 && cr.col >= col && cr.col < col - coloffset) ||
                (rowoffset < 0 && cr.row >= row && cr.row < row - rowoffset)) {
                if (!sheetref) {
                    cr.col = 0;
                    cr.row = 0;
                }
            }

            if (!sheetref) {
                if (cr.col >= col) {
                    cr.col += coloffset;
                }
                if (cr.row >= row) {
                    cr.row += rowoffset;
                }
            }

            if (ttext.charAt(0) === "$") {
                newcr = `$${SocialCalc.rcColname(cr.col)}`;
            } else {
                newcr = SocialCalc.rcColname(cr.col);
            }

            if (ttext.indexOf("$", 1) !== -1) {
                newcr += `$${cr.row}`;
            } else {
                newcr += cr.row;
            }

            if (cr.row < 1 || cr.col < 1) {
                newcr = "#REF!";
            }
            ttext = newcr;
        }
        else if (ttype === token_string) {
            if (ttext.indexOf('"') >= 0) { // quotes to double
                ttext = `"${ttext.replace(/"/, '""')}"`;
            } else {
                ttext = `"${ttext}"`;
            }
        }
        updatedformula += ttext;
    }

    return updatedformula;
};

/**
 * @function SocialCalc.ReplaceFormulaCoords
 * @description Changes all cell references to cells that are keys in movedto to be movedto[coord]
 * @param {string} formula - The formula to modify
 * @param {Object} movedto - Object mapping old coordinates to new coordinates
 * @returns {string} Updated formula with replaced coordinates
 * 
 * @description Don't change references to other sheets. Handle range extents specially.
 * This is typically used when moving cells to update formula references.
 * 
 * @example
 * // Replace moved cell references
 * const moveMap = { "A1": "C1", "B1": "D1" };
 * const newFormula = SocialCalc.ReplaceFormulaCoords("A1+B1+A2", moveMap);
 * // Returns: "C1+D1+A2" (A2 unchanged as it's not in moveMap)
 */
SocialCalc.ReplaceFormulaCoords = function (formula, movedto) {
    let ttype, ttext, i, newcr, coord, parseinfo;
    let updatedformula = "";
    let sheetref = false;

    const scf = SocialCalc.Formula;
    if (!scf) {
        return "Need SocialCalc.Formula";
    }

    const tokentype = scf.TokenType;
    const token_op = tokentype.op;
    const token_string = tokentype.string;
    const token_coord = tokentype.coord;
    const tokenOpExpansion = scf.TokenOpExpansion;

    parseinfo = SocialCalc.Formula.ParseFormulaIntoTokens(formula);

    for (i = 0; i < parseinfo.length; i++) {
        ttype = parseinfo[i].type;
        ttext = parseinfo[i].text;

        if (ttype === token_op) { // references with sheet specifier are not changed
            if (ttext === "!") {
                sheetref = true; // found a sheet reference
            } else if (ttext !== ":") { // for everything but a range, reset
                sheetref = false;
            }

            //!!!! HANDLE RANGE EXTENT MOVES

            ttext = tokenOpExpansion[ttext] || ttext; // make sure short tokens (e.g., "G") go back full (">=")
        }

        if (ttype === token_coord) {
            const cr = SocialCalc.coordToCr(ttext); // get parts
            coord = SocialCalc.crToCoord(cr.col, cr.row); // get "clean" reference

            if (movedto[coord] && !sheetref) { // this is a reference to a moved cell
                const newcr_obj = SocialCalc.coordToCr(movedto[coord]); // get new row and col

                if (ttext.charAt(0) === "$") { // copy absolute ref marks if present
                    newcr = `$${SocialCalc.rcColname(newcr_obj.col)}`;
                } else {
                    newcr = SocialCalc.rcColname(newcr_obj.col);
                }

                if (ttext.indexOf("$", 1) !== -1) {
                    newcr += `$${newcr_obj.row}`;
                } else {
                    newcr += newcr_obj.row;
                }
                ttext = newcr;
            }
        }
        else if (ttype === token_string) {
            if (ttext.indexOf('"') >= 0) { // quotes to double
                ttext = `"${ttext.replace(/"/, '""')}"`;
            } else {
                ttext = `"${ttext}"`;
            }
        }
        updatedformula += ttext;
    }

    return updatedformula;
};

// ************************
//
// Recalc Loop Code
//
// ************************

/**
 * @fileoverview Recalc Loop Implementation
 * 
 * How recalc works:
 * !!!!!!!!!!!!!!
 */

/**
 * @namespace SocialCalc.RecalcInfo
 * @description Object with global recalc info (singleton)
 * @property {SocialCalc.Sheet|null} sheet - Which sheet is being recalced
 * @property {number} currentState - Current state of recalc process
 * @property {Object} state - Allowed state values
 * @property {number|null} recalctimer - Timer ID for canceling timer
 * @property {number} maxtimeslice - Maximum milliseconds per slice of recalc time
 * @property {number} timeslicedelay - Milliseconds to wait between recalc time slices
 * @property {Date} starttime - When recalc started
 * @property {Function} LoadSheet - Function that returns true if started a load
 */
SocialCalc.RecalcInfo = {
    sheet: null, // which sheet is being recalced
    currentState: 0, // current state
    state: {
        start_calc: 1,
        order: 2,
        calc: 3,
        start_wait: 4,
        done_wait: 5
    }, // allowed state values
    recalctimer: null, // value to cancel timer
    maxtimeslice: 100, // maximum milliseconds per slice of recalc time before a wait
    timeslicedelay: 1, // milliseconds to wait between recalc time slices
    starttime: 0, // when recalc started

    /**
     * @description Function that returns true if started a load or false if not
     * @param {string} sheetname - Name of sheet to load
     * @returns {boolean} True if load started, false if not found
     */
    LoadSheet: function (sheetname) {
        return false;
    } // default returns not found
};

/**
 * @class SocialCalc.RecalcData
 * @description Object with recalc info while determining recalc order and afterward
 * @property {boolean} inrecalc - If true, doing a recalc
 * @property {Array} celllist - List with all potential cells to calculate
 * @property {number} celllistitem - Cell to check next when ordering
 * @property {Object|null} calclist - Chained list of cells to calculate
 * @property {number} calclistlength - Number of items in calclist
 * @property {string|null} firstcalc - Start of the calc list
 * @property {string|null} lastcalc - Last one on chain (used to add more to the end)
 * @property {string|null} nextcalc - Used during background recalc to make it restartable
 * @property {number} count - Number calculated
 * @property {Object} checkinfo - Attributes are coords for tracking calc order
 */
SocialCalc.RecalcData = function () { // initialize a RecalcData object
    this.inrecalc = true; // if true, doing a recalc
    this.celllist = []; // list with all potential cells to calculate
    this.celllistitem = 0; // cell to check next when ordering
    this.calclist = null; // object which is the chained list of cells to calculate
    // each in the form of "coord: nextcoord"
    // e.g., if B8 is calculated right after A8, then calclist.A8=="B8"
    // if null, need to create the list
    this.calclistlength = 0; // number of items in calclist

    this.firstcalc = null; // start of the calc list - a string or null
    this.lastcalc = null; // last one on chain (used to add more to the end)

    this.nextcalc = null; // used to keep track during background recalc to make it restartable
    this.count = 0; // number calculated

    // checkinfo is used when determining calc order:
    this.checkinfo = {}; // attributes are coords; if no attrib for a coord, it wasn't checked or doesn't need it
    // values are RecalcCheckInfo objects while checking or TRUE when complete
};

/**
 * @class SocialCalc.RecalcCheckInfo
 * @description Object that stores checking info while determining recalc order
 * @property {string|null} oldcoord - Chain back up of cells referring to cells
 * @property {number} parsepos - Which token we are up to
 * @property {boolean} inrange - If true, in the process of checking a range of coords
 * @property {boolean} inrangestart - If true, have not yet filled in range loop values
 * @property {Object|null} cr1 - Range first coord as a cr object
 * @property {Object|null} cr2 - Range second coord as a cr object
 * @property {number|null} c1 - Range column extent start
 * @property {number|null} c2 - Range column extent end
 * @property {number|null} r1 - Range row extent start
 * @property {number|null} r2 - Range row extent end
 * @property {number|null} c - Current column in range loop
 * @property {number|null} r - Current row in range loop
 */
SocialCalc.RecalcCheckInfo = function () { // initialize a RecalcCheckInfo object
    this.oldcoord = null; // chain back up of cells referring to cells
    this.parsepos = 0; // which token we are up to

    // range info
    this.inrange = false; // if true, in the process of checking a range of coords
    this.inrangestart = false; // if true, have not yet filled in range loop values
    this.cr1 = null; // range first coord as a cr object
    this.cr2 = null; // range second coord as a cr object
    this.c1 = null; // range extents
    this.c2 = null;
    this.r1 = null;
    this.r2 = null;
    this.c = null; // looping values
    this.r = null;
};

/**
 * @function SocialCalc.RecalcSheet
 * @description Recalculates the entire sheet using background processing
 * @param {SocialCalc.Sheet} sheet - The sheet to recalculate
 * 
 * @description Starts the recalc process with proper state management and status callbacks.
 * Uses timeslicing to prevent blocking the UI during long recalculations.
 * 
 * @example
 * // Recalculate a sheet
 * SocialCalc.RecalcSheet(mySheet);
 * // The recalc will run in the background with status updates
 */
SocialCalc.RecalcSheet = function (sheet) {
    const scri = SocialCalc.RecalcInfo;

    delete sheet.attribs.circularreferencecell; // reset recalc-wide things
    SocialCalc.Formula.FreshnessInfoReset();

    SocialCalc.RecalcClearTimeout();

    scri.sheet = sheet; // set values needed by background recalc
    scri.currentState = scri.state.start_calc;
    scri.starttime = new Date();

    if (sheet.statuscallback) {
        sheet.statuscallback(scri, "calcstart", null, sheet.statuscallbackparams);
    }

    SocialCalc.RecalcSetTimeout();
};

/**
 * @function SocialCalc.RecalcSetTimeout
 * @description Sets a timer for the next recalc step
 * @description Uses timeslicing to allow UI updates between calculation steps
 * 
 * @example
 * // This function is called internally by the recalc system
 * SocialCalc.RecalcSetTimeout();
 */
SocialCalc.RecalcSetTimeout = function () {
    const scri = SocialCalc.RecalcInfo;
    scri.recalctimer = window.setTimeout(SocialCalc.RecalcTimerRoutine, scri.timeslicedelay);
};

/**
 * @function SocialCalc.RecalcClearTimeout
 * @description Cancels any pending recalc timeouts
 * @description Used to stop recalc process or when starting a new one
 * 
 * @example
 * // Cancel any pending recalc operations
 * SocialCalc.RecalcClearTimeout();
 */
SocialCalc.RecalcClearTimeout = function () {
    const scri = SocialCalc.RecalcInfo;

    if (scri.recalctimer) {
        window.clearTimeout(scri.recalctimer);
        scri.recalctimer = null;
    }
};
/**
 * @function SocialCalc.RecalcLoadedSheet
 * @description Called when a sheet finishes loading with name, string, and whether it should be recalced
 * @param {string|null} sheetname - Name of the loaded sheet (null to use waiting sheet name)
 * @param {string} str - The sheet data string
 * @param {boolean} recalcneeded - Whether the sheet should be recalculated
 * 
 * @description If loaded sheet has sheet.attribs.recalc=="off", then no recalc is done.
 * If sheetname is null, then the sheetname waiting for will be used.
 * 
 * @example
 * // Load a sheet and trigger recalc if needed
 * SocialCalc.RecalcLoadedSheet("MySheet", sheetData, true);
 * 
 * // Load with waiting sheet name
 * SocialCalc.RecalcLoadedSheet(null, sheetData, false);
 */
SocialCalc.RecalcLoadedSheet = function (sheetname, str, recalcneeded) {
    const scri = SocialCalc.RecalcInfo;
    const scf = SocialCalc.Formula;

    const sheet = SocialCalc.Formula.AddSheetToCache(sheetname || scf.SheetCache.waitingForLoading, str);

    // If recalc needed, and not manual sheet, chain in this new sheet to recalc loop
    if (recalcneeded && sheet && sheet.attribs.recalc !== "off") {
        sheet.previousrecalcsheet = scri.sheet;
        scri.sheet = sheet;
        scri.currentState = scri.state.start_calc;
    }
    scf.SheetCache.waitingForLoading = null;

    SocialCalc.RecalcSetTimeout();
};

/**
 * @function SocialCalc.RecalcTimerRoutine
 * @description Handles the actual order determination and cell-by-cell recalculation in the background
 * @description Uses timeslicing to prevent blocking the UI during long recalculations
 * 
 * @returns {string|undefined} Error message if SocialCalc.Formula is not available
 * 
 * @description This function manages the recalc state machine:
 * - start_calc: Initialize recalc data and cell list
 * - order: Determine calculation order by checking dependencies
 * - calc: Execute calculations on cells in proper order
 * - start_wait/done_wait: Handle loading external sheets or server functions
 * 
 * @example
 * // This function is called automatically by the timer system
 * // It processes recalc in chunks to avoid blocking the UI
 */
SocialCalc.RecalcTimerRoutine = function () {
    let eresult, cell, coord, err, status;
    const starttime = new Date();
    let count = 0;

    const scf = SocialCalc.Formula;
    if (!scf) {
        return "Need SocialCalc.Formula";
    }

    const scri = SocialCalc.RecalcInfo;
    const sheet = scri.sheet;
    if (!sheet) {
        return;
    }
    let recalcdata = sheet.recalcdata;

    /**
     * @description Internal helper to do status callback if required
     * @param {string} status - Status type
     * @param {*} arg - Status-specific argument
     */
    const do_statuscallback = (status, arg) => {
        if (sheet.statuscallback) {
            sheet.statuscallback(recalcdata, status, arg, sheet.statuscallbackparams);
        }
    };

    SocialCalc.RecalcClearTimeout();

    // Initialize recalc process
    if (scri.currentState === scri.state.start_calc) {
        recalcdata = new SocialCalc.RecalcData();
        sheet.recalcdata = recalcdata;

        // Get list of cells to check for order
        for (coord in sheet.cells) {
            if (!coord) continue;
            recalcdata.celllist.push(coord);
        }

        recalcdata.calclist = {}; // start with empty list
        scri.currentState = scri.state.order; // drop through to determining recalc order
    }

    // Determine calculation order
    if (scri.currentState === scri.state.order) {
        // Check all cells to see if they should be on the list
        while (recalcdata.celllistitem < recalcdata.celllist.length) {
            coord = recalcdata.celllist[recalcdata.celllistitem++];
            err = SocialCalc.RecalcCheckCell(sheet, coord);

            // If taking too long, give up CPU for a while
            if (((new Date()) - starttime) >= scri.maxtimeslice) {
                do_statuscallback("calcorder", {
                    coord: coord,
                    total: recalcdata.celllist.length,
                    count: recalcdata.celllistitem
                });
                SocialCalc.RecalcSetTimeout();
                return;
            }
        }

        do_statuscallback("calccheckdone", recalcdata.calclistlength);

        recalcdata.nextcalc = recalcdata.firstcalc; // start at the beginning of the recalc chain
        scri.currentState = scri.state.calc; // loop through cells on next timer call
        SocialCalc.RecalcSetTimeout();
        return;
    }

    // Starting to wait for something (sheet load)
    if (scri.currentState === scri.state.start_wait) {
        scri.currentState = scri.state.done_wait; // finished on next timer call
        if (scri.LoadSheet) {
            status = scri.LoadSheet(scf.SheetCache.waitingForLoading);
            if (status) { // started a load operation
                return;
            }
        }
        SocialCalc.RecalcLoadedSheet(null, "", false);
        return;
    }

    // Done waiting, continue calculation
    if (scri.currentState === scri.state.done_wait) {
        scri.currentState = scri.state.calc; // loop through cells on next timer call
        SocialCalc.RecalcSetTimeout();
        return;
    }

    // Main calculation state
    if (scri.currentState !== scri.state.calc) {
        alert(`Recalc state error: ${scri.currentState}. Error in SocialCalc code.`);
    }

    // Process cells in calculation order
    coord = sheet.recalcdata.nextcalc;
    while (coord) {
        cell = sheet.cells[coord];
        eresult = scf.evaluate_parsed_formula(cell.parseinfo, sheet, false);

        // Check if waiting for sheet to load
        if (scf.SheetCache.waitingForLoading) {
            recalcdata.nextcalc = coord; // start with this cell again
            recalcdata.count += count;
            do_statuscallback("calcloading", { sheetname: scf.SheetCache.waitingForLoading });
            scri.currentState = scri.state.start_wait; // start load on next timer call
            SocialCalc.RecalcSetTimeout();
            return;
        }

        // Check if waiting for server function
        if (scf.RemoteFunctionInfo.waitingForServer) {
            recalcdata.nextcalc = coord; // start with this cell again
            recalcdata.count += count;
            do_statuscallback("calcserverfunc", {
                funcname: scf.RemoteFunctionInfo.waitingForServer,
                coord: coord,
                total: recalcdata.calclistlength,
                count: recalcdata.count
            });
            scri.currentState = scri.state.done_wait; // start load on next timer call
            return; // return and wait for next recalc timer event
        }

        // Update cell if value changed
        if (cell.datavalue !== eresult.value || cell.valuetype !== eresult.type) {
            cell.datavalue = eresult.value;
            cell.valuetype = eresult.type;
            delete cell.displaystring;
            sheet.recalcchangedavalue = true; // remember something changed
        }

        if (eresult.error) {
            cell.errors = eresult.error;
        }

        count++;
        coord = sheet.recalcdata.calclist[coord];

        // If taking too long, give up CPU for a while
        if (((new Date()) - starttime) >= scri.maxtimeslice) {
            recalcdata.nextcalc = coord; // start with next cell on chain
            recalcdata.count += count;
            do_statuscallback("calcstep", {
                coord: coord,
                total: recalcdata.calclistlength,
                count: recalcdata.count
            });
            SocialCalc.RecalcSetTimeout();
            return;
        }
    }

    // Recalc complete
    recalcdata.inrecalc = false;
    delete sheet.recalcdata; // save memory and clear out for name lookup formula evaluation
    delete sheet.attribs.needsrecalc; // remember recalc done

    // Chain back if doing recalc of loaded sheets
    scri.sheet = sheet.previousrecalcsheet || null;
    if (scri.sheet) {
        scri.currentState = scri.state.calc; // start where we left off
        SocialCalc.RecalcSetTimeout();
        return;
    }

    scf.FreshnessInfo.recalc_completed = true; // say freshness info is complete
    do_statuscallback("calcfinished", (new Date()) - scri.starttime);
};

/**
 * @function SocialCalc.RecalcCheckCell
 * @description Checks cell to put on calclist, looking at parsed tokens and dependencies
 * @param {SocialCalc.Sheet} sheet - The sheet containing the cell
 * @param {string} startcoord - The coordinate of the cell to check (e.g., "A1")
 * @returns {string} Error message if circular reference found, empty string otherwise
 * 
 * @description Analyzes cell formulas to determine calculation dependencies.
 * Builds a dependency tree and detects circular references.
 * Handles ranges, named ranges, and individual cell references.
 * 
 * @example
 * // Check a cell for dependencies and add to calc list
 * const error = SocialCalc.RecalcCheckCell(sheet, "A1");
 * if (error) console.error("Circular reference:", error);
 */
SocialCalc.RecalcCheckCell = function (sheet, startcoord) {
    let parseinfo, ttext, ttype, i, rangecoord, circref, value, pos, pos2, cell, coordvals;

    const scf = SocialCalc.Formula;
    if (!scf) {
        return "Need SocialCalc.Formula";
    }

    const tokentype = scf.TokenType;
    const token_op = tokentype.op;
    const token_name = tokentype.name;
    const token_coord = tokentype.coord;

    const recalcdata = sheet.recalcdata;
    const checkinfo = recalcdata.checkinfo;

    let sheetref = false; // if true, a sheet reference is in effect, so don't check that
    let oldcoord = null; // coord of formula that referred to this one when checking down the tree
    let coord = startcoord; // the coord of the cell we are checking

    // Start with requested cell, and then continue down or up the dependency tree
    // oldcoord (and checkinfo[coord].oldcoord) maintains the reference stack during the tree walk
    // checkinfo[coord] maintains the stack of checking looping values, e.g., token number being checked

    mainloop:
    while (coord) {
        cell = sheet.cells[coord];
        coordvals = checkinfo[coord];

        // Don't calculate if not a formula or already calculated
        if (!cell || cell.datatype !== "f" || (coordvals && typeof coordvals !== "object")) {
            coord = oldcoord; // go back up dependency tree to coord that referred to us
            if (checkinfo[coord]) oldcoord = checkinfo[coord].oldcoord;
            continue;
        }

        // Create checking information if not exists
        if (!coordvals) {
            coordvals = new SocialCalc.RecalcCheckInfo();
            checkinfo[coord] = coordvals;
        }

        // Delete errors from previous recalcs
        if (cell.errors) {
            delete cell.errors;
        }

        // Cache parsed formula
        if (!cell.parseinfo) {
            cell.parseinfo = scf.ParseFormulaIntoTokens(cell.formula);
        }
        parseinfo = cell.parseinfo;

        // Go through each token in formula
        for (i = coordvals.parsepos; i < parseinfo.length; i++) {

            // Processing a range of coords
            if (coordvals.inrange) {
                if (coordvals.inrangestart) { // first time - fill in other values
                    if (coordvals.cr1.col > coordvals.cr2.col) {
                        coordvals.c1 = coordvals.cr2.col;
                        coordvals.c2 = coordvals.cr1.col;
                    } else {
                        coordvals.c1 = coordvals.cr1.col;
                        coordvals.c2 = coordvals.cr2.col;
                    }
                    coordvals.c = coordvals.c1 - 1; // start one before

                    if (coordvals.cr1.row > coordvals.cr2.row) {
                        coordvals.r1 = coordvals.cr2.row;
                        coordvals.r2 = coordvals.cr1.row;
                    } else {
                        coordvals.r1 = coordvals.cr1.row;
                        coordvals.r2 = coordvals.cr2.row;
                    }
                    coordvals.r = coordvals.r1; // start on this row
                    coordvals.inrangestart = false;
                }

                coordvals.c += 1; // increment column
                if (coordvals.c > coordvals.c2) { // finished the columns of this row
                    coordvals.r += 1; // increment row
                    if (coordvals.r > coordvals.r2) { // finished checking the entire range
                        coordvals.inrange = false;
                        continue;
                    }
                    coordvals.c = coordvals.c1; // start at the beginning of next row
                }
                rangecoord = SocialCalc.crToCoord(coordvals.c, coordvals.r);

                // Check this range coordinate
                coordvals.parsepos = i; // remember our position
                coordvals.oldcoord = oldcoord; // remember back up chain
                oldcoord = coord; // come back to us
                coord = rangecoord;

                // Check for circular reference
                if (checkinfo[coord] && typeof checkinfo[coord] === "object") {
                    cell.errors = SocialCalc.Constants.s_caccCircRef + startcoord;
                    checkinfo[startcoord] = true;
                    if (!recalcdata.firstcalc) {
                        recalcdata.firstcalc = startcoord;
                    } else {
                        recalcdata.calclist[recalcdata.lastcalc] = startcoord;
                    }
                    recalcdata.lastcalc = startcoord;
                    recalcdata.calclistlength++;
                    sheet.attribs.circularreferencecell = `${coord}|${oldcoord}`;
                    return cell.errors;
                }
                continue mainloop;
            }

            ttype = parseinfo[i].type; // get token details
            ttext = parseinfo[i].text;

            // Handle sheet references
            if (ttype === token_op) {
                if (ttext === "!") {
                    sheetref = true; // found a sheet reference
                } else if (ttext !== ":") { // for everything but a range, reset
                    sheetref = false;
                }
            }

            // Look for named range
            if (ttype === token_name) {
                value = scf.LookupName(sheet, ttext);
                if (value.type === "range") { // only need to recurse here for range
                    pos = value.value.indexOf("|");
                    if (pos !== -1) { // range - check each cell
                        coordvals.cr1 = SocialCalc.coordToCr(value.value.substring(0, pos));
                        pos2 = value.value.indexOf("|", pos + 1);
                        coordvals.cr2 = SocialCalc.coordToCr(value.value.substring(pos + 1, pos2));
                        coordvals.inrange = true;
                        coordvals.inrangestart = true;
                        i = i - 1; // back up so will start up again here
                        continue;
                    }
                } else if (value.type === "coord") { // just a coord
                    ttype = token_coord; // treat as a coord inline
                    ttext = value.value; // and then drop through to next test
                }
            }

            // Handle coordinate tokens
            if (ttype === token_coord) {
                // Look for a range
                if (i >= 2 &&
                    parseinfo[i - 1].type === token_op && parseinfo[i - 1].text === ':' &&
                    parseinfo[i - 2].type === token_coord &&
                    !sheetref) {
                    // Range -- check each cell
                    coordvals.cr1 = SocialCalc.coordToCr(parseinfo[i - 2].text);
                    coordvals.cr2 = SocialCalc.coordToCr(ttext);
                    coordvals.inrange = true; // next time use the range looping code
                    coordvals.inrangestart = true;
                    i = i - 1; // back up so will start up again here
                    continue;
                }
                // Single cell reference
                else if (!sheetref) {
                    if (ttext.indexOf("$") !== -1) ttext = ttext.replace(/\$/g, ""); // remove any $'s
                    coordvals.parsepos = i + 1; // remember our position - come back on next token
                    coordvals.oldcoord = oldcoord; // remember back up chain
                    oldcoord = coord; // come back to us
                    coord = ttext;

                    // Check for circular reference
                    if (checkinfo[coord] && typeof checkinfo[coord] === "object") {
                        cell.errors = SocialCalc.Constants.s_caccCircRef + startcoord;
                        checkinfo[startcoord] = true;
                        if (!recalcdata.firstcalc) {
                            recalcdata.firstcalc = startcoord;
                        } else {
                            recalcdata.calclist[recalcdata.lastcalc] = startcoord;
                        }
                        recalcdata.lastcalc = startcoord;
                        recalcdata.calclistlength++;
                        sheet.attribs.circularreferencecell = `${coord}|${oldcoord}`;
                        return cell.errors;
                    }
                    continue mainloop;
                }
            }
        }

        sheetref = false; // make sure off when bump back up

        checkinfo[coord] = true; // this one is finished

        // Add to calclist
        if (!recalcdata.firstcalc) {
            recalcdata.firstcalc = coord;
        } else {
            recalcdata.calclist[recalcdata.lastcalc] = coord;
        }
        recalcdata.lastcalc = coord;
        recalcdata.calclistlength++; // count number on list

        // Go back to the formula that referred to us and continue
        coord = oldcoord;
        oldcoord = checkinfo[coord] ? checkinfo[coord].oldcoord : null;
    }

    return "";
};
/**
 * @class SocialCalc.Parse
 * @description Used by ExecuteSheetCommand to get elements of commands to execute
 * @param {string} str - The string to parse, consisting of one or more lines with tokens separated by delimiters
 * 
 * @description The string consists of one or more lines each made up of one or more tokens
 * separated by a delimiter. Provides methods to traverse and extract tokens from the string.
 * 
 * @property {string} str - The string being parsed
 * @property {number} pos - Current position in the string
 * @property {string} delimiter - Token delimiter (default: " ")
 * @property {number} lineEnd - Position of end of current line
 * 
 * @example
 * // Parse a command string
 * const parser = new SocialCalc.Parse("set A1 value n 42\nset B1 text t Hello");
 * const command = parser.NextToken(); // "set"
 * const coord = parser.NextToken();   // "A1"
 * const rest = parser.RestOfString(); // "value n 42"
 */
SocialCalc.Parse = function (str) {
    // Properties
    this.str = str;
    this.pos = 0;
    this.delimiter = " ";
    this.lineEnd = str.indexOf("\n");
    if (this.lineEnd < 0) {
        this.lineEnd = str.length;
    }
};

/**
 * @method NextToken
 * @memberof SocialCalc.Parse
 * @description Returns the next token as a string
 * @returns {string} The next token, or empty string if at end
 * 
 * @example
 * const parser = new SocialCalc.Parse("set A1 value");
 * console.log(parser.NextToken()); // "set"
 * console.log(parser.NextToken()); // "A1"
 * console.log(parser.NextToken()); // "value"
 */
SocialCalc.Parse.prototype.NextToken = function () {
    if (this.pos < 0) return "";

    let pos2 = this.str.indexOf(this.delimiter, this.pos);
    const pos1 = this.pos;

    if (pos2 > this.lineEnd) { // don't go past end of line
        pos2 = this.lineEnd;
    }

    if (pos2 >= 0) {
        this.pos = pos2 + 1;
        return this.str.substring(pos1, pos2);
    } else {
        this.pos = this.lineEnd;
        return this.str.substring(pos1, this.lineEnd);
    }
};

/**
 * @method RestOfString
 * @memberof SocialCalc.Parse
 * @description Returns everything from current point until end of line and advances position
 * @returns {string} Remaining string on current line
 * 
 * @example
 * const parser = new SocialCalc.Parse("set A1 value n 42");
 * parser.NextToken(); // "set"
 * parser.NextToken(); // "A1"
 * console.log(parser.RestOfString()); // "value n 42"
 */
SocialCalc.Parse.prototype.RestOfString = function () {
    const oldpos = this.pos;
    if (this.pos < 0 || this.pos >= this.lineEnd) return "";

    this.pos = this.lineEnd;
    return this.str.substring(oldpos, this.lineEnd);
};

/**
 * @method RestOfStringNoMove
 * @memberof SocialCalc.Parse
 * @description Returns everything from current point until end of line without advancing position
 * @returns {string} Remaining string on current line
 * 
 * @example
 * const parser = new SocialCalc.Parse("set A1 value n 42");
 * parser.NextToken(); // "set"
 * const remaining = parser.RestOfStringNoMove(); // "A1 value n 42"
 * // Position unchanged, can continue parsing from same point
 */
SocialCalc.Parse.prototype.RestOfStringNoMove = function () {
    if (this.pos < 0 || this.pos >= this.lineEnd) return "";
    return this.str.substring(this.pos, this.lineEnd);
};

/**
 * @method NextLine
 * @memberof SocialCalc.Parse
 * @description Moves current position to the next line
 * 
 * @example
 * const parser = new SocialCalc.Parse("line1\nline2\nline3");
 * parser.RestOfString(); // "line1"
 * parser.NextLine();
 * parser.RestOfString(); // "line2"
 */
SocialCalc.Parse.prototype.NextLine = function () {
    this.pos = this.lineEnd + 1;
    this.lineEnd = this.str.indexOf("\n", this.pos);
    if (this.lineEnd < 0) {
        this.lineEnd = this.str.length;
    }
};

/**
 * @method EOF
 * @memberof SocialCalc.Parse
 * @description Checks if at end of string with no more to process
 * @returns {boolean} True if at end of string, false otherwise
 * 
 * @example
 * const parser = new SocialCalc.Parse("short");
 * while (!parser.EOF()) {
 *     console.log(parser.NextToken());
 * }
 */
SocialCalc.Parse.prototype.EOF = function () {
    return this.pos < 0 || this.pos >= this.str.length;
};

/**
 * @class SocialCalc.UndoStack
 * @description Implements the behavior needed for a normal application's undo/redo stack
 * 
 * @description You add a new change sequence with PushChange. The type argument is a string
 * that can be used to lookup some general string like "typing" or "setting attribute" for
 * the menu prompts for undo/redo.
 * 
 * You add the "do" steps with AddDo. The non-null, non-undefined arguments are joined
 * together with " " to make a command string to be saved.
 * 
 * You add the undo steps as commands for the most recent change with AddUndo.
 * 
 * The Undo and Redo functions move the Top Of Stack pointer through the changes stack
 * so you can undo and redo. Doing a new PushChange removes all undone items after TOS.
 * 
 * @property {Array} stack - Array of change objects: {command: [], type: string, undo: []}
 * @property {number} tos - Top of stack position, used for undo/redo
 * @property {number} maxRedo - Maximum size of redo stack (0 = no limit)
 * @property {number} maxUndo - Maximum number of steps kept for undo (0 = no limit)
 * 
 * @example
 * const undoStack = new SocialCalc.UndoStack();
 * undoStack.PushChange("typing");
 * undoStack.AddDo("set", "A1", "value", "n", "42");
 * undoStack.AddUndo("set", "A1", "empty");
 * // Later...
 * if (undoStack.Undo()) {
 *     // Execute undo commands
 * }
 */
SocialCalc.UndoStack = function () {
    // Properties
    this.stack = []; // {command: [], type: type, undo: []} -- multiple dos and undos allowed
    this.tos = -1; // top of stack position, used for undo/redo
    this.maxRedo = 0; // Maximum size of redo stack (and audit trail) or zero if no limit
    this.maxUndo = 50; // Maximum number of steps kept for undo or zero if no limit
};

/**
 * @method PushChange
 * @memberof SocialCalc.UndoStack
 * @description Adds a new change to the stack, removing any undone items after current position
 * @param {string} type - Type of change (e.g., "typing", "formatting")
 * 
 * @example
 * undoStack.PushChange("cell edit");
 */
SocialCalc.UndoStack.prototype.PushChange = function (type) {
    // Pop off things not redone
    while (this.stack.length > 0 && this.stack.length - 1 > this.tos) {
        this.stack.pop();
    }

    this.stack.push({ command: [], type: type, undo: [] });

    // Limit number kept as audit trail
    if (this.maxRedo && this.stack.length > this.maxRedo) {
        this.stack.shift(); // remove the extra one
    }

    // Trim excess undo info
    if (this.maxUndo && this.stack.length > this.maxUndo) {
        this.stack[this.stack.length - this.maxUndo - 1].undo = []; // only need to remove one
    }

    this.tos = this.stack.length - 1;
};

/**
 * @method AddDo
 * @memberof SocialCalc.UndoStack
 * @description Adds a "do" command to the most recent change
 * @param {...*} args - Arguments to be joined with " " to create command string
 * 
 * @example
 * undoStack.AddDo("set", "A1", "value", "n", "42");
 */
SocialCalc.UndoStack.prototype.AddDo = function () {
    const args = [];
    for (let i = 0; i < arguments.length; i++) {
        if (arguments[i] != null) args.push(arguments[i]); // ignore null or undefined
    }
    const cmd = args.join(" ");
    this.stack[this.stack.length - 1].command.push(cmd);
};

/**
 * @method AddUndo
 * @memberof SocialCalc.UndoStack
 * @description Adds an "undo" command to the most recent change
 * @param {...*} args - Arguments to be joined with " " to create command string
 * 
 * @example
 * undoStack.AddUndo("set", "A1", "empty");
 */
SocialCalc.UndoStack.prototype.AddUndo = function () {
    const args = [];
    for (let i = 0; i < arguments.length; i++) {
        if (arguments[i] != null) args.push(arguments[i]); // ignore null or undefined
    }
    const cmd = args.join(" ");
    this.stack[this.stack.length - 1].undo.push(cmd);
};

/**
 * @method TOS
 * @memberof SocialCalc.UndoStack
 * @description Returns the current top of stack item
 * @returns {Object|null} The current stack item or null if empty
 * 
 * @example
 * const currentItem = undoStack.TOS();
 * if (currentItem) {
 *     console.log("Current change type:", currentItem.type);
 * }
 */
SocialCalc.UndoStack.prototype.TOS = function () {
    return this.tos >= 0 ? this.stack[this.tos] : null;
};

/**
 * @method Undo
 * @memberof SocialCalc.UndoStack
 * @description Moves the TOS pointer back one position for undo
 * @returns {boolean} True if undo is possible, false otherwise
 * 
 * @example
 * if (undoStack.Undo()) {
 *     const item = undoStack.TOS();
 *     // Execute undo commands from item.undo
 * }
 */
SocialCalc.UndoStack.prototype.Undo = function () {
    if (this.tos >= 0 && (!this.maxUndo || this.tos > this.stack.length - this.maxUndo - 1)) {
        this.tos -= 1;
        return true;
    } else {
        return false;
    }
};

/**
 * @method Redo
 * @memberof SocialCalc.UndoStack
 * @description Moves the TOS pointer forward one position for redo
 * @returns {boolean} True if redo is possible, false otherwise
 * 
 * @example
 * if (undoStack.Redo()) {
 *     const item = undoStack.TOS();
 *     // Execute redo commands from item.command
 * }
 */
SocialCalc.UndoStack.prototype.Redo = function () {
    if (this.tos < this.stack.length - 1) {
        this.tos += 1;
        return true;
    } else {
        return false;
    }
};

/**
 * @namespace SocialCalc.Clipboard
 * @description Single object that stores the clipboard, shared by all active sheets
 * @description Like the undo stack, it does not persist from one editing session to another
 * 
 * @property {string} clipboard - Empty or string in save format with "copiedfrom:" set to a range
 * 
 * @example
 * // Copy data to clipboard
 * SocialCalc.Clipboard.clipboard = sheetSaveData;
 * 
 * // Check if clipboard has data
 * if (SocialCalc.Clipboard.clipboard) {
 *     // Paste operation can proceed
 * }
 */
SocialCalc.Clipboard = {
    clipboard: "" // empty or string in save format with "copiedfrom:" set to a range
};

/**
 * @class SocialCalc.RenderContext
 * @description Context object for rendering spreadsheet data with display settings and styling
 * @param {SocialCalc.Sheet} sheetobj - The sheet object to render
 * @throws {Error} Throws error if sheetobj is missing
 * 
 * @description This class manages all the rendering settings, styling, and layout information
 * needed to display a spreadsheet. It handles panes, highlights, fonts, colors, and other
 * display attributes.
 * 
 * @property {SocialCalc.Sheet} sheetobj - The sheet being rendered
 * @property {boolean} hideRowsCols - Whether to hide rows/columns (panes only work with false)
 * @property {boolean} showGrid - Whether to show grid lines
 * @property {boolean} showRCHeaders - Whether to show row/column headers
 * @property {number} rownamewidth - Width of row name column
 * @property {number} pixelsPerRow - Assumed height per row in pixels
 * @property {Object} cellskip - Coordinates of cells covering other cells
 * @property {Object} coordToCR - For cells starting spans: coord -> {row, col}
 * @property {Array<number>} colwidth - Precomputed column widths
 * @property {number} totalwidth - Precomputed total table width
 * @property {Array<Object>} rowpanes - Pane definitions: {first: row, last: row}
 * @property {Array<Object>} colpanes - Pane definitions: {first: col, last: col}
 * @property {number} maxcol - Maximum column to display
 * @property {number} maxrow - Maximum row to display
 * @property {Object} highlights - Special display cells: coord -> highlightType
 * @property {string} cursorsuffix - Added to cursor highlights for type lookup
 * @property {Object} highlightTypes - Highlight style definitions
 * @property {string} cellIDprefix - Prefix for cell IDs in rendered HTML
 * @property {Object|null} defaultlinkstyle - Default link style object
 * @property {Object} defaultHTMLlinkstyle - Default HTML link style
 * 
 * @example
 * // Create render context for a sheet
 * const renderContext = new SocialCalc.RenderContext(mySheet);
 * renderContext.showGrid = true;
 * renderContext.showRCHeaders = true;
 * 
 * // Set up highlights
 * renderContext.highlights["A1"] = "cursor";
 * renderContext.highlights["B1:C3"] = "range";
 */
SocialCalc.RenderContext = function (sheetobj) {
    if (!sheetobj) {
        throw SocialCalc.Constants.s_rcMissingSheet;
    }

    const attribs = sheetobj.attribs;
    const scc = SocialCalc.Constants;

    // Properties
    this.sheetobj = sheetobj;
    this.hideRowsCols = false; // Rendering with panes only works with "false"
    // !!!! Note: not implemented yet in rendering, just saved as an attribute
    this.showGrid = false;
    this.showRCHeaders = false;
    this.rownamewidth = scc.defaultRowNameWidth;
    this.pixelsPerRow = scc.defaultAssumedRowHeight;

    this.cellskip = {}; // if present, coord of cell covering this cell
    this.coordToCR = {}; // for cells starting spans, coordToCR[coord]={row:row, col:col}
    this.colwidth = []; // precomputed column widths, taking into account defaults
    this.totalwidth = 0; // precomputed total table width

    this.rowpanes = []; // for each pane, {first: firstrow, last: lastrow}
    this.colpanes = []; // for each pane, {first: firstrow, last: lastrow}
    this.maxcol = 0; // max col and row to display, adding long spans, etc.
    this.maxrow = 0;

    this.highlights = {}; // for each cell with special display: coord:highlightType
    this.cursorsuffix = ""; // added to highlights[cr]=="cursor" to get type to lookup

    /**
     * @description Highlight type definitions with styles and class names
     * @type {Object}
     */
    this.highlightTypes = {
        cursor: {
            style: scc.defaultHighlightTypeCursorStyle,
            className: scc.defaultHighlightTypeCursorClass
        },
        range: {
            style: scc.defaultHighlightTypeRangeStyle,
            className: scc.defaultHighlightTypeRangeClass
        },
        cursorinsertup: {
            style: `color:#FFF;backgroundColor:#A6A6A6;backgroundRepeat:repeat-x;backgroundPosition:top left;backgroundImage:url(${scc.defaultImagePrefix}cursorinsertup.gif);`,
            className: scc.defaultHighlightTypeCursorClass
        },
        cursorinsertleft: {
            style: `color:#FFF;backgroundColor:#A6A6A6;backgroundRepeat:repeat-y;backgroundPosition:top left;backgroundImage:url(${scc.defaultImagePrefix}cursorinsertleft.gif);`,
            className: scc.defaultHighlightTypeCursorClass
        },
        range2: {
            style: `color:#000;backgroundColor:#FFF;backgroundImage:url(${scc.defaultImagePrefix}range2.gif);`,
            className: ""
        }
    };

    this.cellIDprefix = scc.defaultCellIDPrefix; // if non-null, each cell will render with an ID

    this.defaultlinkstyle = null; // default linkstyle object (allows you to pass values to link renderer)
    this.defaultHTMLlinkstyle = { type: "html" }; // default linkstyle for standalone HTML

    // Constants
    this.defaultfontstyle = scc.defaultCellFontStyle;
    this.defaultfontsize = scc.defaultCellFontSize;
    this.defaultfontfamily = scc.defaultCellFontFamily;

    this.defaultlayout = scc.defaultCellLayout;

    this.defaultpanedividerwidth = scc.defaultPaneDividerWidth;
    this.defaultpanedividerheight = scc.defaultPaneDividerHeight;

    this.gridCSS = scc.defaultGridCSS;

    this.commentClassName = scc.defaultCommentClass; // for cells with non-blank comments when showGrid is true
    this.commentCSS = scc.defaultCommentStyle; // any combination of classnames and styles may be used
    this.commentNoGridClassName = scc.defaultCommentNoGridClass; // for cells when showGrid is false
    this.commentNoGridCSS = scc.defaultCommentNoGridStyle; // any combination of classnames and styles may be used

    /**
     * @description CSS class names for various UI elements
     * @type {Object}
     */
    this.classnames = {
        colname: scc.defaultColnameClass,
        rowname: scc.defaultRownameClass,
        selectedcolname: scc.defaultSelectedColnameClass,
        selectedrowname: scc.defaultSelectedRownameClass,
        upperleft: scc.defaultUpperLeftClass,
        skippedcell: scc.defaultSkippedCellClass,
        panedivider: scc.defaultPaneDividerClass
    };

    /**
     * @description Explicit CSS styles (can be used without stylesheet)
     * @type {Object}
     */
    this.explicitStyles = {
        colname: scc.defaultColnameStyle,
        rowname: scc.defaultRownameStyle,
        selectedcolname: scc.defaultSelectedColnameStyle,
        selectedrowname: scc.defaultSelectedRownameStyle,
        upperleft: scc.defaultUpperLeftStyle,
        skippedcell: scc.defaultSkippedCellStyle,
        panedivider: scc.defaultPaneDividerStyle
    };

    // Processed info about cell skipping
    this.cellskip = null;
    this.needcellskip = true;

    // Precomputed values, filling in defaults indicated by "*"
    this.fonts = []; // for each fontnum, {style: fs, weight: fw, size: fs, family: ff}
    this.layouts = []; // for each layout, "padding:Tpx Rpx Bpx Lpx;vertical-align:va;"

    this.needprecompute = true; // need to call PrecomputeSheetFontsAndLayouts

    // Initialize with sheet object
    this.rowpanes[0] = { first: 1, last: attribs.lastrow };
    this.colpanes[0] = { first: 1, last: attribs.lastcol };
};
/**
 * @description Prototype method to precompute sheet fonts and layouts
 * @memberof SocialCalc.RenderContext
 * @returns {void}
 */
SocialCalc.RenderContext.prototype.PrecomputeSheetFontsAndLayouts = function () {
    SocialCalc.PrecomputeSheetFontsAndLayouts(this);
};

/**
 * @description Prototype method to calculate cell skip data for merged cells
 * @memberof SocialCalc.RenderContext
 * @returns {void}
 */
SocialCalc.RenderContext.prototype.CalculateCellSkipData = function () {
    SocialCalc.CalculateCellSkipData(this);
};

/**
 * @description Prototype method to calculate column width data
 * @memberof SocialCalc.RenderContext
 * @returns {void}
 */
SocialCalc.RenderContext.prototype.CalculateColWidthData = function () {
    SocialCalc.CalculateColWidthData(this);
};

/**
 * @description Sets the first and last row for a specific row pane
 * @memberof SocialCalc.RenderContext
 * @param {number} panenum - The pane number
 * @param {number} first - First row in the pane
 * @param {number} last - Last row in the pane
 */
SocialCalc.RenderContext.prototype.SetRowPaneFirstLast = function (panenum, first, last) {
    this.rowpanes[panenum] = { first: first, last: last };
};

/**
 * @description Sets the first and last column for a specific column pane
 * @memberof SocialCalc.RenderContext
 * @param {number} panenum - The pane number
 * @param {number} first - First column in the pane
 * @param {number} last - Last column in the pane
 */
SocialCalc.RenderContext.prototype.SetColPaneFirstLast = function (panenum, first, last) {
    this.colpanes[panenum] = { first: first, last: last };
};

/**
 * @description Checks if a coordinate is within a specific pane
 * @memberof SocialCalc.RenderContext
 * @param {string} coord - Cell coordinate (e.g., "A1")
 * @param {number} rowpane - Row pane number
 * @param {number} colpane - Column pane number
 * @returns {boolean} True if coordinate is in the pane
 */
SocialCalc.RenderContext.prototype.CoordInPane = function (coord, rowpane, colpane) {
    return SocialCalc.CoordInPane(this, coord, rowpane, colpane);
};

/**
 * @description Checks if a cell (by row/col numbers) is within a specific pane
 * @memberof SocialCalc.RenderContext
 * @param {number} row - Row number
 * @param {number} col - Column number
 * @param {number} rowpane - Row pane number
 * @param {number} colpane - Column pane number
 * @returns {boolean} True if cell is in the pane
 */
SocialCalc.RenderContext.prototype.CellInPane = function (row, col, rowpane, colpane) {
    return SocialCalc.CellInPane(this, row, col, rowpane, colpane);
};

/**
 * @description Initializes table object with proper styling and attributes
 * @memberof SocialCalc.RenderContext
 * @param {HTMLTableElement} tableobj - The table DOM element to initialize
 */
SocialCalc.RenderContext.prototype.InitializeTable = function (tableobj) {
    SocialCalc.InitializeTable(this, tableobj);
};

/**
 * @description Renders the complete sheet as a table
 * @memberof SocialCalc.RenderContext
 * @param {HTMLTableElement|null} oldtable - Previous table to replace (null for new table)
 * @param {Object|null} linkstyle - Link styling options
 * @returns {HTMLTableElement} The rendered table element
 */
SocialCalc.RenderContext.prototype.RenderSheet = function (oldtable, linkstyle) {
    return SocialCalc.RenderSheet(this, oldtable, linkstyle);
};

/**
 * @description Renders the column group element for table structure
 * @memberof SocialCalc.RenderContext
 * @returns {HTMLElement} The colgroup element
 */
SocialCalc.RenderContext.prototype.RenderColGroup = function () {
    return SocialCalc.RenderColGroup(this);
};

/**
 * @description Renders the column headers row
 * @memberof SocialCalc.RenderContext
 * @returns {HTMLTableRowElement|null} The header row element
 */
SocialCalc.RenderContext.prototype.RenderColHeaders = function () {
    return SocialCalc.RenderColHeaders(this);
};

/**
 * @description Renders a sizing row for column width control
 * @memberof SocialCalc.RenderContext
 * @returns {HTMLTableRowElement} The sizing row element
 */
SocialCalc.RenderContext.prototype.RenderSizingRow = function () {
    return SocialCalc.RenderSizingRow(this);
};

/**
 * @description Renders a specific row with all its cells
 * @memberof SocialCalc.RenderContext
 * @param {number} rownum - Row number to render
 * @param {number} rowpane - Row pane number
 * @param {Object|null} linkstyle - Link styling options
 * @returns {HTMLTableRowElement} The rendered row element
 */
SocialCalc.RenderContext.prototype.RenderRow = function (rownum, rowpane, linkstyle) {
    return SocialCalc.RenderRow(this, rownum, rowpane, linkstyle);
};

/**
 * @description Renders a spacing row between panes
 * @memberof SocialCalc.RenderContext
 * @returns {HTMLTableRowElement} The spacing row element
 */
SocialCalc.RenderContext.prototype.RenderSpacingRow = function () {
    return SocialCalc.RenderSpacingRow(this);
};

/**
 * @description Renders a specific cell
 * @memberof SocialCalc.RenderContext
 * @param {number} rownum - Row number
 * @param {number} colnum - Column number
 * @param {number} rowpane - Row pane number
 * @param {number} colpane - Column pane number
 * @param {boolean} noElement - Whether to skip creating DOM element
 * @param {Object|null} linkstyle - Link styling options
 * @returns {HTMLTableCellElement|null} The rendered cell element
 */
SocialCalc.RenderContext.prototype.RenderCell = function (rownum, colnum, rowpane, colpane, noElement, linkstyle) {
    return SocialCalc.RenderCell(this, rownum, colnum, rowpane, colpane, noElement, linkstyle);
};

/**
 * @function SocialCalc.PrecomputeSheetFontsAndLayouts
 * @description Precomputes font and layout information by resolving defaults and wildcards
 * @param {SocialCalc.RenderContext} context - The render context
 * 
 * @description Processes font definitions by replacing "*" placeholders with appropriate defaults.
 * Also processes layout definitions by filling in default padding and vertical alignment values.
 * This preprocessing step improves rendering performance by avoiding repeated parsing.
 * 
 * @example
 * // This is called automatically by the rendering system
 * SocialCalc.PrecomputeSheetFontsAndLayouts(renderContext);
 */
SocialCalc.PrecomputeSheetFontsAndLayouts = function (context) {
    let defaultfont, parts, layoutre, dparts, sparts, num, s, i;
    const sheetobj = context.sheetobj;
    const attribs = sheetobj.attribs;

    // Process default font if defined
    if (attribs.defaultfont) {
        defaultfont = sheetobj.fonts[attribs.defaultfont];
        defaultfont = defaultfont.replace(/^\*/, SocialCalc.Constants.defaultCellFontStyle);
        defaultfont = defaultfont.replace(/(.+)\*(.+)/, `$1${SocialCalc.Constants.defaultCellFontSize}$2`);
        defaultfont = defaultfont.replace(/\*$/, SocialCalc.Constants.defaultCellFontFamily);
        parts = defaultfont.match(/^(\S+? \S+?) (\S+?) (\S.*)$/);
        context.defaultfontstyle = parts[1];
        context.defaultfontsize = parts[2];
        context.defaultfontfamily = parts[3];
    }

    // Precompute fonts by filling in the *'s
    for (num = 1; num < sheetobj.fonts.length; num++) {
        s = sheetobj.fonts[num];
        s = s.replace(/^\*/, context.defaultfontstyle);
        s = s.replace(/(.+)\*(.+)/, `$1${context.defaultfontsize}$2`);
        s = s.replace(/\*$/, context.defaultfontfamily);
        parts = s.match(/^(\S+?) (\S+?) (\S+?) (\S.*)$/);
        context.fonts[num] = {
            style: parts[1],
            weight: parts[2],
            size: parts[3],
            family: parts[4]
        };
    }

    layoutre = /^padding:\s*(\S+)\s+(\S+)\s+(\S+)\s+(\S+);vertical-align:\s*(\S+);/;
    dparts = SocialCalc.Constants.defaultCellLayout.match(layoutre); // get built-in defaults

    if (attribs.defaultlayout) {
        sparts = sheetobj.layouts[attribs.defaultlayout].match(layoutre); // get sheet defaults, if set
    } else {
        sparts = ["", "*", "*", "*", "*", "*"];
    }

    // Precompute layouts by filling in the *'s
    for (num = 1; num < sheetobj.layouts.length; num++) {
        s = sheetobj.layouts[num];
        parts = s.match(layoutre);
        for (i = 1; i <= 5; i++) {
            if (parts[i] === "*") {
                parts[i] = (sparts[i] !== "*" ? sparts[i] : dparts[i]); // if *, sheet default or built-in
            }
        }
        context.layouts[num] = `padding:${parts[1]} ${parts[2]} ${parts[3]} ${parts[4]};vertical-align:${parts[5]};`;
    }

    context.needprecompute = false;
};

/**
 * @function SocialCalc.CalculateCellSkipData
 * @description Calculates which cells are skipped due to merged cells (colspan/rowspan)
 * @param {SocialCalc.RenderContext} context - The render context
 * 
 * @description Analyzes all cells to identify merged cells and marks cells that should be
 * skipped during rendering. Updates context.cellskip and context.coordToCR mappings.
 * Also updates maxrow and maxcol based on spans extending beyond the basic sheet bounds.
 * 
 * @example
 * // This is called automatically when rendering needs updated skip data
 * SocialCalc.CalculateCellSkipData(renderContext);
 */
SocialCalc.CalculateCellSkipData = function (context) {
    let row, col, coord, cell, colspan, rowspan, skiprow, skipcol, skipcoord;

    const sheetobj = context.sheetobj;
    context.maxrow = 0;
    context.maxcol = 0;
    context.cellskip = {}; // reset

    // Calculate cellskip data
    for (row = 1; row <= sheetobj.attribs.lastrow; row++) {
        for (col = 1; col <= sheetobj.attribs.lastcol; col++) {
            // Look for spans and set cellskip for skipped cells
            coord = SocialCalc.crToCoord(col, row);
            cell = sheetobj.cells[coord];

            // Don't look at undefined cells (they have no spans) or skipped cells
            if (cell === undefined || context.cellskip[coord]) continue;

            colspan = cell.colspan || 1;
            rowspan = cell.rowspan || 1;

            if (colspan > 1 || rowspan > 1) {
                for (skiprow = row; skiprow < row + rowspan; skiprow++) {
                    for (skipcol = col; skipcol < col + colspan; skipcol++) {
                        // Do the setting on individual cells
                        skipcoord = SocialCalc.crToCoord(skipcol, skiprow);
                        if (skipcoord === coord) {
                            // For coord, remember row and col
                            context.coordToCR[coord] = { row: row, col: col };
                        } else {
                            // For other cells, flag with coord of here
                            context.cellskip[skipcoord] = coord;
                        }
                        if (skiprow > context.maxrow) context.maxrow = skiprow;
                        if (skipcol > context.maxcol) context.maxcol = skipcol;
                    }
                }
            }
        }
    }

    context.needcellskip = false;
};

/**
 * @function SocialCalc.CalculateColWidthData
 * @description Calculates column width data for all visible columns
 * @param {SocialCalc.RenderContext} context - The render context
 * 
 * @description Processes column width information from sheet attributes and defaults.
 * Handles special width values like "blank", "auto", and percentage values.
 * Updates context.colwidth array and context.totalwidth.
 * 
 * @example
 * // This is called automatically before rendering to ensure current width data
 * SocialCalc.CalculateColWidthData(renderContext);
 */
SocialCalc.CalculateColWidthData = function (context) {
    let colnum, colname, colwidth, colpane;

    const sheetobj = context.sheetobj;
    let totalwidth = context.showRCHeaders ? context.rownamewidth - 0 : 0;

    // Calculate column width data
    for (colpane = 0; colpane < context.colpanes.length; colpane++) {
        for (colnum = context.colpanes[colpane].first; colnum <= context.colpanes[colpane].last; colnum++) {
            colname = SocialCalc.rcColname(colnum);
            colwidth = sheetobj.colattribs.width[colname] ||
                sheetobj.attribs.defaultcolwidth ||
                SocialCalc.Constants.defaultColWidth;

            if (colwidth === "blank" || colwidth === "auto") colwidth = "";

            context.colwidth[colnum] = colwidth + "";
            totalwidth += (colwidth && ((colwidth - 0) > 0)) ? (colwidth - 0) : 10;
        }
    }
    context.totalwidth = totalwidth;
};

/**
 * @function SocialCalc.InitializeTable
 * @description Initializes table object with proper styling and attributes for cross-browser compatibility
 * @param {SocialCalc.RenderContext} context - The render context
 * @param {HTMLTableElement} tableobj - The table DOM element to initialize
 * 
 * @description Sets up the table with border-collapse, proper spacing, and width.
 * Uses border-collapse to avoid holes at corners. Note that IE and Firefox handle
 * <col> elements differently under border-collapse, and Safari has issues with <col>
 * and wide text. Table-layout "fixed" can also cause problems.
 * 
 * The rendering assumes fixed column widths, even though SocialCalc allows "auto".
 * There may be issues with "auto" and it's hard to make it work cross-browser
 * with border-collapse.
 * 
 * @example
 * const table = document.createElement("table");
 * SocialCalc.InitializeTable(renderContext, table);
 */
SocialCalc.InitializeTable = function (context, tableobj) {
    /*
    Uses border-collapse so corners don't have holes
    Note: IE and Firefox handle <col> differently (IE adds borders and padding)
    under border-collapse and Safari has problems with <col> and wide text
    Tablelayout "fixed" also leads to problems
    */

    /*
    *** Discussion ***
    
    The rendering assumes fixed column widths, even though SocialCalc allows "auto".
    There may be issues with "auto" and it is hard to make it work cross-browser
    with border-collapse, etc.
    
    This and the RenderSheet routine are where in the code the specifics of
    table attributes and column size definitions are set. As the browsers settle down
    and when we decide if we don't need auto width, we may want to revisit the way the
    code does this (e.g., use table-layout:fixed).
    */

    tableobj.style.borderCollapse = "collapse";
    tableobj.cellSpacing = "0";
    tableobj.cellPadding = "0";
    tableobj.style.width = `${context.totalwidth}px`;
};

/**
 * @function SocialCalc.RenderSheet
 * @description Renders a complete sheet as a DOM table element
 * @param {SocialCalc.RenderContext} context - The render context
 * @param {HTMLTableElement|null} oldtable - Previous table to replace (null for new table)
 * @param {Object|null} linkstyle - Link styling options ("" or null for editing, object for formatting)
 * @returns {HTMLTableElement} The rendered table element
 * 
 * @description Creates a complete HTML table representation of the spreadsheet.
 * If oldtable is provided, replaces it in the DOM. Handles precomputation of
 * fonts, layouts, and cell skip data as needed. Calls EvalUserScripts if available.
 * 
 * @example
 * // Render a new sheet
 * const table = SocialCalc.RenderSheet(context, null, null);
 * document.body.appendChild(table);
 * 
 * // Replace existing table
 * const newTable = SocialCalc.RenderSheet(context, oldTable, { type: "html" });
 */
SocialCalc.RenderSheet = function (context, oldtable, linkstyle) {
    let newrow, rowpane, rownum;
    let tableobj, colgroupobj, tbodyobj, parentnode;

    // Do precompute stuff if necessary
    if (context.sheetobj.changedrendervalues) {
        context.needcellskip = true;
        context.needprecompute = true;
        context.sheetobj.changedrendervalues = false;
    }
    if (context.needcellskip) {
        context.CalculateCellSkipData();
    }
    if (context.needprecompute) {
        context.PrecomputeSheetFontsAndLayouts();
    }

    context.CalculateColWidthData(); // always make sure col width values are up to date

    // Make the table element and fill it in
    tableobj = document.createElement("table");
    context.InitializeTable(tableobj);

    colgroupobj = context.RenderColGroup();
    tableobj.appendChild(colgroupobj);

    tbodyobj = document.createElement("tbody");

    tbodyobj.appendChild(context.RenderSizingRow());

    if (context.showRCHeaders) {
        newrow = context.RenderColHeaders();
        if (newrow) tbodyobj.appendChild(newrow);
    }

    for (rowpane = 0; rowpane < context.rowpanes.length; rowpane++) {
        for (rownum = context.rowpanes[rowpane].first; rownum <= context.rowpanes[rowpane].last; rownum++) {
            newrow = context.RenderRow(rownum, rowpane, linkstyle);
            tbodyobj.appendChild(newrow);
        }
        if (rowpane < context.rowpanes.length - 1) {
            newrow = context.RenderSpacingRow();
            tbodyobj.appendChild(newrow);
        }
    }

    tableobj.appendChild(tbodyobj);

    if (oldtable) {
        parentnode = oldtable.parentNode;
        if (parentnode) parentnode.replaceChild(tableobj, oldtable);
    }

    // Call EvalUserScripts if available (defined in workbook-control.js)
    if (typeof SocialCalc.EvalUserScripts === 'function') {
        SocialCalc.EvalUserScripts();
    }

    return tableobj;
};

/**
 * @function SocialCalc.RenderRow
 * @description Renders a complete row with all its cells and headers
 * @param {SocialCalc.RenderContext} context - The render context
 * @param {number} rownum - Row number to render
 * @param {number} rowpane - Row pane number
 * @param {Object|null} linkstyle - Link styling options
 * @returns {HTMLTableRowElement} The rendered row element
 * 
 * @description Creates a table row element containing row headers (if enabled),
 * all cells in the row across all column panes, and pane dividers as needed.
 * 
 * @example
 * const row = SocialCalc.RenderRow(context, 5, 0, null);
 * tbody.appendChild(row);
 */
SocialCalc.RenderRow = function (context, rownum, rowpane, linkstyle) {
    const sheetobj = context.sheetobj;
    const result = document.createElement("tr");
    let colnum, newcol, colpane, newdiv;

    // Add row header if showing row/column headers
    if (context.showRCHeaders) {
        newcol = document.createElement("td");
        if (context.classnames) newcol.className = context.classnames.rowname;
        if (context.explicitStyles) newcol.style.cssText = context.explicitStyles.rowname;
        newcol.width = context.rownamewidth;
        // To get around Safari making top of centered row number be considered top of row
        newcol.style.verticalAlign = "top";
        newcol.innerHTML = rownum + "";
        result.appendChild(newcol);
    }

    // Render cells in all column panes
    for (colpane = 0; colpane < context.colpanes.length; colpane++) {
        for (colnum = context.colpanes[colpane].first; colnum <= context.colpanes[colpane].last; colnum++) {
            newcol = context.RenderCell(rownum, colnum, rowpane, colpane, null, linkstyle);
            if (newcol) result.appendChild(newcol);
        }

        // Add pane divider if not last pane
        if (colpane < context.colpanes.length - 1) {
            newcol = document.createElement("td");
            newcol.width = context.defaultpanedividerwidth;
            if (context.classnames.panedivider) newcol.className = context.classnames.panedivider;
            if (context.explicitStyles.panedivider) newcol.style.cssText = context.explicitStyles.panedivider;

            // Add div for Firefox to avoid squishing
            newdiv = document.createElement("div");
            newdiv.style.width = `${context.defaultpanedividerwidth}px`;
            newdiv.style.overflow = "hidden";
            newcol.appendChild(newdiv);
            result.appendChild(newcol);
        }
    }
    return result;
};

/**
 * @function SocialCalc.RenderSpacingRow
 * @description Renders a spacing row between row panes
 * @param {SocialCalc.RenderContext} context - The render context
 * @returns {HTMLTableRowElement} The spacing row element
 * 
 * @description Creates a row that provides vertical spacing between row panes.
 * The row contains cells with appropriate width and height to maintain proper
 * layout and visual separation between panes.
 * 
 * @example
 * const spacingRow = SocialCalc.RenderSpacingRow(context);
 * tbody.appendChild(spacingRow);
 */
SocialCalc.RenderSpacingRow = function (context) {
    let colnum, newcol, colpane, w;
    const result = document.createElement("tr");

    // Add row header spacing if showing headers
    if (context.showRCHeaders) {
        newcol = document.createElement("td");
        newcol.width = context.rownamewidth;
        newcol.height = context.defaultpanedividerheight;
        if (context.classnames.panedivider) newcol.className = context.classnames.panedivider;
        if (context.explicitStyles.panedivider) newcol.style.cssText = context.explicitStyles.panedivider;
        result.appendChild(newcol);
    }

    // Add spacing cells for all column panes
    for (colpane = 0; colpane < context.colpanes.length; colpane++) {
        for (colnum = context.colpanes[colpane].first; colnum <= context.colpanes[colpane].last; colnum++) {
            newcol = document.createElement("td");
            w = context.colwidth[colnum];
            if (w) newcol.width = w;
            newcol.height = context.defaultpanedividerheight;
            if (context.classnames.panedivider) newcol.className = context.classnames.panedivider;
            if (context.explicitStyles.panedivider) newcol.style.cssText = context.explicitStyles.panedivider;
            if (newcol) result.appendChild(newcol);
        }

        // Add pane divider if not last pane
        if (colpane < context.colpanes.length - 1) {
            newcol = document.createElement("td");
            newcol.width = context.defaultpanedividerwidth;
            newcol.height = context.defaultpanedividerheight;
            if (context.classnames.panedivider) newcol.className = context.classnames.panedivider;
            if (context.explicitStyles.panedivider) newcol.style.cssText = context.explicitStyles.panedivider;
            result.appendChild(newcol);
        }
    }
    return result;
};

/**
 * @function SocialCalc.RenderColHeaders
 * @description Renders the column headers row
 * @param {SocialCalc.RenderContext} context - The render context
 * @returns {HTMLTableRowElement|null} The header row element, or null if headers not shown
 * 
 * @description Creates a row containing column headers (A, B, C, etc.) if row/column
 * headers are enabled. Includes upper-left corner cell and pane dividers as needed.
 * 
 * @example
 * const headerRow = SocialCalc.RenderColHeaders(context);
 * if (headerRow) tbody.appendChild(headerRow);
 */
SocialCalc.RenderColHeaders = function (context) {
    const result = document.createElement("tr");
    let colnum, newcol, colpane;

    if (!context.showRCHeaders) return null;

    // Add upper-left corner cell
    newcol = document.createElement("td");
    if (context.classnames) newcol.className = context.classnames.upperleft;
    if (context.explicitStyles) newcol.style.cssText = context.explicitStyles.upperleft;
    newcol.width = context.rownamewidth;
    result.appendChild(newcol);

    // Add column header cells for all panes
    for (colpane = 0; colpane < context.colpanes.length; colpane++) {
        for (colnum = context.colpanes[colpane].first; colnum <= context.colpanes[colpane].last; colnum++) {
            newcol = document.createElement("td");
            if (context.classnames) newcol.className = context.classnames.colname;
            if (context.explicitStyles) newcol.style.cssText = context.explicitStyles.colname;
            newcol.innerHTML = SocialCalc.rcColname(colnum);
            result.appendChild(newcol);
        }

        // Add pane divider if not last pane
        if (colpane < context.colpanes.length - 1) {
            newcol = document.createElement("td");
            newcol.width = context.defaultpanedividerwidth;
            if (context.classnames.panedivider) newcol.className = context.classnames.panedivider;
            if (context.explicitStyles.panedivider) newcol.style.cssText = context.explicitStyles.panedivider;
            result.appendChild(newcol);
        }
    }
    return result;
};
/**
 * @function SocialCalc.RenderColGroup
 * @description Renders the colgroup element that defines column widths for the table
 * @param {SocialCalc.RenderContext} context - The render context
 * @returns {HTMLElement} The colgroup element with column definitions
 * 
 * @description Creates a colgroup element containing col elements that define
 * the width of each column in the table. Handles row headers, all column panes,
 * and pane dividers to ensure proper table layout.
 * 
 * @example
 * const colgroup = SocialCalc.RenderColGroup(context);
 * table.appendChild(colgroup);
 */
SocialCalc.RenderColGroup = function (context) {
    let colpane, colnum, newcol, t;
    const sheetobj = context.sheetobj;
    const result = document.createElement("colgroup");

    // Add column for row headers if showing them
    if (context.showRCHeaders) {
        newcol = document.createElement("col");
        newcol.width = context.rownamewidth;
        result.appendChild(newcol);
    }

    // Add columns for each pane
    for (colpane = 0; colpane < context.colpanes.length; colpane++) {
        for (colnum = context.colpanes[colpane].first; colnum <= context.colpanes[colpane].last; colnum++) {
            newcol = document.createElement("col");
            t = context.colwidth[colnum];
            if (t) newcol.width = t;
            result.appendChild(newcol);
        }

        // Add pane divider column if not last pane
        if (colpane < context.colpanes.length - 1) {
            newcol = document.createElement("col");
            newcol.width = context.defaultpanedividerwidth;
            result.appendChild(newcol);
        }
    }
    return result;
};

/**
 * @function SocialCalc.RenderSizingRow
 * @description Renders a sizing row to establish column widths
 * @param {SocialCalc.RenderContext} context - The render context
 * @returns {HTMLTableRowElement} The sizing row element
 * 
 * @description Creates a row with height=1 that contains cells with specific widths
 * to ensure proper column sizing. This is a common technique for controlling
 * table layout in older browsers.
 * 
 * @example
 * const sizingRow = SocialCalc.RenderSizingRow(context);
 * tbody.appendChild(sizingRow);
 */
SocialCalc.RenderSizingRow = function (context) {
    let colpane, colnum, newcell, t;
    const result = document.createElement("tr");

    // Add sizing cell for row header if showing headers
    if (context.showRCHeaders) {
        newcell = document.createElement("td");
        newcell.style.width = `${context.rownamewidth}px`;
        newcell.height = "1";
        result.appendChild(newcell);
    }

    // Add sizing cells for each column in each pane
    for (colpane = 0; colpane < context.colpanes.length; colpane++) {
        for (colnum = context.colpanes[colpane].first; colnum <= context.colpanes[colpane].last; colnum++) {
            newcell = document.createElement("td");
            t = context.colwidth[colnum];
            if (t) newcell.width = t;
            newcell.height = "1";
            result.appendChild(newcell);
        }

        // Add sizing cell for pane divider if not last pane
        if (colpane < context.colpanes.length - 1) {
            newcell = document.createElement("td");
            newcell.width = context.defaultpanedividerwidth;
            newcell.height = "1";
            result.appendChild(newcell);
        }
    }
    return result;
};

/**
 * @function SocialCalc.RenderCell
 * @description Renders a single spreadsheet cell with all styling and content
 * @param {SocialCalc.RenderContext} context - The render context
 * @param {number} rownum - Row number of the cell
 * @param {number} colnum - Column number of the cell
 * @param {number} rowpane - Row pane number
 * @param {number} colpane - Column pane number
 * @param {boolean} noElement - If true, creates a pseudo element instead of DOM element
 * @param {Object|null} linkstyle - Link styling options
 * @returns {HTMLTableCellElement|Object|null} The rendered cell element, pseudo element, or null if skipped
 * 
 * @description Creates a fully styled table cell with proper content, formatting,
 * borders, colors, fonts, and spans. Handles merged cells, skipped cells,
 * and various cell types. Applies highlighting and grid styles as needed.
 * 
 * @example
 * // Render a normal cell
 * const cell = SocialCalc.RenderCell(context, 5, 3, 0, 0, false, null);
 * row.appendChild(cell);
 * 
 * // Create pseudo element for calculations
 * const pseudoCell = SocialCalc.RenderCell(context, 1, 1, 0, 0, true, null);
 */
SocialCalc.RenderCell = function (context, rownum, colnum, rowpane, colpane, noElement, linkstyle) {
    const sheetobj = context.sheetobj;
    let num, t, result, span, cell, sheetattribs;
    let stylestr = "";

    rownum = rownum - 0; // make sure a number
    colnum = colnum - 0;

    const coord = SocialCalc.crToCoord(colnum, rownum);

    // Handle skipped cells (within merged cell spans)
    if (context.cellskip[coord]) {
        if (context.CoordInPane(context.cellskip[coord], rowpane, colpane)) {
            return null; // span starts in this pane -- so just skip
        }
        // Span start is scrolled away, so make a special cell
        result = noElement ? SocialCalc.CreatePseudoElement() : document.createElement("td");
        if (context.classnames.skippedcell) result.className = context.classnames.skippedcell;
        if (context.explicitStyles.skippedcell) result.style.cssText = context.explicitStyles.skippedcell;
        result.innerHTML = "&nbsp;"; // put something there so height is OK
        // !!! Really need to add borders in case there isn't anything else shown in the pane to get height
        return result;
    }

    result = noElement ? SocialCalc.CreatePseudoElement() : document.createElement("td");

    // Set cell ID if prefix is defined
    if (context.cellIDprefix) {
        result.id = context.cellIDprefix + coord;
    }

    cell = sheetobj.cells[coord];
    if (!cell) {
        cell = new SocialCalc.Cell(coord);
    }

    sheetattribs = sheetobj.attribs;
    const scc = SocialCalc.Constants;

    // Handle column spanning
    if (cell.colspan > 1) {
        span = 1;
        for (num = 1; num < cell.colspan; num++) {
            if (sheetobj.colattribs.hide[SocialCalc.rcColname(colnum + num)] !== "yes" &&
                context.CellInPane(rownum, colnum + num, rowpane, colpane)) {
                span++;
            }
        }
        result.colSpan = span;
    }

    // Handle row spanning
    if (cell.rowspan > 1) {
        span = 1;
        for (num = 1; num < cell.rowspan; num++) {
            if (sheetobj.rowattribs.hide[(rownum + num) + ""] !== "yes" &&
                context.CellInPane(rownum + num, colnum, rowpane, colpane)) {
                span++;
            }
        }
        result.rowSpan = span;
    }

    // Set cell content
    if (cell.displaystring === undefined) { // cache the display value
        cell.displaystring = SocialCalc.FormatValueForDisplay(
            sheetobj,
            cell.datavalue,
            coord,
            (linkstyle || context.defaultlinkstyle)
        );
    } else {
        // Callout to execute scripts if needed
        SocialCalc.CallOutOnRenderCell(sheetobj, cell.datavalue, coord);
    }
    result.innerHTML = cell.displaystring;

    // Apply layout styling
    num = cell.layout || sheetattribs.defaultlayout;
    if (num) {
        stylestr += context.layouts[num]; // use precomputed layout with "*"'s filled in
    } else {
        stylestr += scc.defaultCellLayout;
    }

    // Apply font styling
    num = cell.font || sheetattribs.defaultfont;
    if (num) { // get expanded font strings in context
        t = context.fonts[num]; // do each - plain "font:" style sets all sorts of other values
        stylestr += `font-style:${t.style};font-weight:${t.weight};font-size:${t.size};font-family:${t.family};`;
    } else {
        if (scc.defaultCellFontSize) {
            stylestr += `font-size:${scc.defaultCellFontSize};`;
        }
        if (scc.defaultCellFontFamily) {
            stylestr += `font-family:${scc.defaultCellFontFamily};`;
        }
    }

    // Apply text color
    num = cell.color || sheetattribs.defaultcolor;
    if (num) stylestr += `color:${sheetobj.colors[num]};`;

    // Apply background color
    num = cell.bgcolor || sheetattribs.defaultbgcolor;
    if (num) stylestr += `background-color:${sheetobj.colors[num]};`;

    // Apply text alignment
    num = cell.cellformat;
    if (num) {
        stylestr += `text-align:${sheetobj.cellformats[num]};`;
    } else {
        t = cell.valuetype.charAt(0);
        if (t === "t") {
            num = sheetattribs.defaulttextformat;
            if (num) stylestr += `text-align:${sheetobj.cellformats[num]};`;
        } else if (t === "n") {
            num = sheetattribs.defaultnontextformat;
            if (num) {
                stylestr += `text-align:${sheetobj.cellformats[num]};`;
            } else {
                stylestr += "text-align:right;";
            }
        } else {
            stylestr += "text-align:left;";
        }
    }

    // Apply borders
    num = cell.bt;
    if (num) stylestr += `border-top:${sheetobj.borderstyles[num]};`;

    num = cell.br;
    if (num) {
        stylestr += `border-right:${sheetobj.borderstyles[num]};`;
    } else if (context.showGrid) {
        if (context.CellInPane(rownum, colnum + (cell.colspan || 1), rowpane, colpane)) {
            t = SocialCalc.crToCoord(colnum + (cell.colspan || 1), rownum);
        } else {
            t = "nomatch";
        }
        if (context.cellskip[t]) t = context.cellskip[t];
        if (!sheetobj.cells[t] || !sheetobj.cells[t].bl) {
            stylestr += `border-right:${context.gridCSS}`;
        }
    }

    num = cell.bb;
    if (num) {
        stylestr += `border-bottom:${sheetobj.borderstyles[num]};`;
    } else if (context.showGrid) {
        if (context.CellInPane(rownum + (cell.rowspan || 1), colnum, rowpane, colpane)) {
            t = SocialCalc.crToCoord(colnum, rownum + (cell.rowspan || 1));
        } else {
            t = "nomatch";
        }
        if (context.cellskip[t]) t = context.cellskip[t];
        if (!sheetobj.cells[t] || !sheetobj.cells[t].bt) {
            stylestr += `border-bottom:${context.gridCSS}`;
        }
    }

    num = cell.bl;
    if (num) stylestr += `border-left:${sheetobj.borderstyles[num]};`;

    // Apply comment styling
    if (cell.comment) {
        if (context.showGrid) {
            if (context.commentClassName) {
                result.className = (result.className ? `${result.className} ` : "") + context.commentClassName;
            }
            stylestr += context.commentCSS;
        } else {
            if (context.commentNoGridClassName) {
                result.className = (result.className ? `${result.className} ` : "") + context.commentNoGridClassName;
            }
            stylestr += context.commentNoGridCSS;
        }
    }

    result.style.cssText = stylestr;

    //!!!!!!!!!
    // NOTE: csss and cssc are not supported yet.
    // csss needs to be parsed into pieces to override just the attributes specified, not all with assignment to cssText.
    // cssc just needs to set the className.

    // Apply highlighting
    t = context.highlights[coord];
    if (t) { // this is a highlighted cell: Override style appropriately
        if (t === "cursor") t += context.cursorsuffix; // cursor can take alternative forms
        if (context.highlightTypes[t].className) {
            result.className = (result.className ? `${result.className} ` : "") + context.highlightTypes[t].className;
        }
        SocialCalc.setStyles(result, context.highlightTypes[t].style);
    }

    return result;
};

/**
 * @function SocialCalc.CoordInPane
 * @description Checks if a coordinate is within a specific pane
 * @param {SocialCalc.RenderContext} context - The render context
 * @param {string} coord - Cell coordinate (e.g., "A1")
 * @param {number} rowpane - Row pane number
 * @param {number} colpane - Column pane number
 * @returns {boolean} True if coordinate is in the specified pane
 * @throws {Error} Throws error if coordToCR mapping is invalid
 * 
 * @example
 * const inPane = SocialCalc.CoordInPane(context, "B5", 0, 0);
 * if (inPane) {
 *     // Cell B5 is visible in the current pane
 * }
 */
SocialCalc.CoordInPane = function (context, coord, rowpane, colpane) {
    const coordToCR = context.coordToCR[coord];
    if (!coordToCR || !coordToCR.row || !coordToCR.col) {
        throw `Bad coordToCR for ${coord}`;
    }
    return context.CellInPane(coordToCR.row, coordToCR.col, rowpane, colpane);
};

/**
 * @function SocialCalc.CellInPane
 * @description Checks if a cell (by row/col numbers) is within a specific pane
 * @param {SocialCalc.RenderContext} context - The render context
 * @param {number} row - Row number
 * @param {number} col - Column number
 * @param {number} rowpane - Row pane number
 * @param {number} colpane - Column pane number
 * @returns {boolean} True if cell is within the specified pane boundaries
 * @throws {Error} Throws error if pane numbers are invalid
 * 
 * @example
 * const inPane = SocialCalc.CellInPane(context, 5, 3, 0, 0);
 * if (inPane) {
 *     // Cell at row 5, column 3 is in pane 0,0
 * }
 */
SocialCalc.CellInPane = function (context, row, col, rowpane, colpane) {
    const panerowlimits = context.rowpanes[rowpane];
    const panecollimits = context.colpanes[colpane];

    if (!panerowlimits || !panecollimits) {
        throw `CellInPane called with unknown panes ${rowpane}/${colpane}`;
    }

    if (row < panerowlimits.first || row > panerowlimits.last) return false;
    if (col < panecollimits.first || col > panecollimits.last) return false;

    return true;
};

/**
 * @function SocialCalc.CreatePseudoElement
 * @description Creates a pseudo DOM element for calculations without actual DOM manipulation
 * @returns {Object} A pseudo element object with basic DOM-like properties
 * 
 * @description Used when noElement is true in rendering functions to create
 * lightweight objects that mimic DOM elements for measurement and calculation
 * purposes without the overhead of actual DOM creation.
 * 
 * @example
 * const pseudoEl = SocialCalc.CreatePseudoElement();
 * pseudoEl.innerHTML = "test content";
 * pseudoEl.style.cssText = "color: red;";
 */
SocialCalc.CreatePseudoElement = function () {
    return {
        style: { cssText: "" },
        innerHTML: "",
        className: ""
    };
};
/**
 * @fileoverview Miscellaneous utility functions for SocialCalc
 * @description Various helper functions for coordinate conversion, encoding/decoding,
 * formatting, and DOM manipulation utilities.
 */

/**
 * @function SocialCalc.rcColname
 * @description Converts a column number to Excel-style column name (A, B, ..., Z, AA, AB, ...)
 * @param {number} c - Column number (1-based, max 702 for ZZ)
 * @returns {string} Column name (A, B, C, ..., Z, AA, AB, ..., ZZ)
 * 
 * @example
 * SocialCalc.rcColname(1);   // "A"
 * SocialCalc.rcColname(26);  // "Z"
 * SocialCalc.rcColname(27);  // "AA"
 * SocialCalc.rcColname(702); // "ZZ"
 */
SocialCalc.rcColname = function (c) {
    if (c > 702) c = 702; // maximum number of columns - ZZ
    if (c < 1) c = 1;

    const collow = (c - 1) % 26 + 65;
    const colhigh = Math.floor((c - 1) / 26);

    if (colhigh) {
        return String.fromCharCode(colhigh + 64) + String.fromCharCode(collow);
    } else {
        return String.fromCharCode(collow);
    }
};

/**
 * @constant {Array<string>} SocialCalc.letters
 * @description Array of letters A-Z for column name generation
 */
SocialCalc.letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M",
    "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];

/**
 * @function SocialCalc.crToCoord
 * @description Converts column and row numbers to coordinate string
 * @param {number} c - Column number (1-based)
 * @param {number} r - Row number (1-based)
 * @returns {string} Coordinate string (e.g., "A1", "B5", "AA100")
 * 
 * @example
 * SocialCalc.crToCoord(1, 1);   // "A1"
 * SocialCalc.crToCoord(2, 5);   // "B5"
 * SocialCalc.crToCoord(27, 10); // "AA10"
 */
SocialCalc.crToCoord = function (c, r) {
    if (c < 1) c = 1;
    if (c > 702) c = 702; // maximum number of columns - ZZ
    if (r < 1) r = 1;

    const collow = (c - 1) % 26;
    const colhigh = Math.floor((c - 1) / 26);

    let result;
    if (colhigh) {
        result = SocialCalc.letters[colhigh - 1] + SocialCalc.letters[collow] + r;
    } else {
        result = SocialCalc.letters[collow] + r;
    }
    return result;
};

/**
 * @namespace SocialCalc coordinate caches
 * @description Caches to improve performance of coordinate conversions
 */
SocialCalc.coordToCol = {}; // too expensive to set in crToCoord since that is called so many times
SocialCalc.coordToRow = {};

/**
 * @function SocialCalc.coordToCr
 * @description Converts coordinate string to column and row numbers
 * @param {string} cr - Coordinate string (e.g., "A1", "$B$5", "AA100")
 * @returns {Object} Object with row and col properties
 * @returns {number} returns.row - Row number (1-based)
 * @returns {number} returns.col - Column number (1-based)
 * 
 * @description Uses caching to improve performance. Handles absolute references ($A$1).
 * This was faster than using regexes and assumes well-formed input.
 * 
 * @example
 * SocialCalc.coordToCr("A1");    // {row: 1, col: 1}
 * SocialCalc.coordToCr("$B$5");  // {row: 5, col: 2}
 * SocialCalc.coordToCr("AA10");  // {row: 10, col: 27}
 */
SocialCalc.coordToCr = function (cr) {
    const r = SocialCalc.coordToRow[cr];
    if (r) return { row: r, col: SocialCalc.coordToCol[cr] };

    let c = 0, row = 0;
    for (let i = 0; i < cr.length; i++) { // this was faster than using regexes; assumes well-formed
        const ch = cr.charCodeAt(i);
        if (ch === 36); // skip $'s
        else if (ch <= 57) row = 10 * row + ch - 48;
        else if (ch >= 97) c = 26 * c + ch - 96;
        else if (ch >= 65) c = 26 * c + ch - 64;
    }

    SocialCalc.coordToCol[cr] = c;
    SocialCalc.coordToRow[cr] = row;
    return { row: row, col: c };
};

/**
 * @function SocialCalc.ParseRange
 * @description Parses a range string into start and end coordinates
 * @param {string} range - Range string (e.g., "A1:B5", "C3", "A1:A1")
 * @returns {Object} Object with cr1 and cr2 properties
 * @returns {Object} returns.cr1 - Start coordinate with row, col, and coord properties
 * @returns {Object} returns.cr2 - End coordinate with row, col, and coord properties
 * 
 * @example
 * SocialCalc.ParseRange("A1:B5");
 * // Returns: {
 * //   cr1: {row: 1, col: 1, coord: "A1"},
 * //   cr2: {row: 5, col: 2, coord: "B5"}
 * // }
 * 
 * SocialCalc.ParseRange("C3");
 * // Returns: {
 * //   cr1: {row: 3, col: 3, coord: "C3"},
 * //   cr2: {row: 3, col: 3, coord: "C3"}
 * // }
 */
SocialCalc.ParseRange = function (range) {
    let pos, cr, cr1, cr2;

    if (!range) range = "A1:A1"; // error return, hopefully benign
    range = range.toUpperCase();
    pos = range.indexOf(":");

    if (pos >= 0) {
        cr = range.substring(0, pos);
        cr1 = SocialCalc.coordToCr(cr);
        cr1.coord = cr;
        cr = range.substring(pos + 1);
        cr2 = SocialCalc.coordToCr(cr);
        cr2.coord = cr;
    } else {
        cr1 = SocialCalc.coordToCr(range);
        cr1.coord = range;
        cr2 = SocialCalc.coordToCr(range);
        cr2.coord = range;
    }
    return { cr1: cr1, cr2: cr2 };
};

/**
 * @function SocialCalc.decodeFromSave
 * @description Decodes strings from save format by unescaping special characters
 * @param {*} s - String to decode (returns unchanged if not string)
 * @returns {*} Decoded string or original value
 * 
 * @description Converts escaped characters back to original:
 * - \\c -> :
 * - \\n -> newline
 * - \\b -> \\
 * 
 * @example
 * SocialCalc.decodeFromSave("Hello\\cWorld");    // "Hello:World"
 * SocialCalc.decodeFromSave("Line1\\nLine2");    // "Line1\nLine2"
 * SocialCalc.decodeFromSave("Backslash\\bHere"); // "Backslash\Here"
 */
SocialCalc.decodeFromSave = function (s) {
    if (typeof s !== "string") return s;
    if (s.indexOf("\\") === -1) return s; // for performance reasons: replace nothing takes up time

    let r = s.replace(/\\c/g, ":");
    r = r.replace(/\\n/g, "\n");
    return r.replace(/\\b/g, "\\");
};

/**
 * @function SocialCalc.decodeFromAjax
 * @description Decodes strings from AJAX format by unescaping special characters
 * @param {*} s - String to decode (returns unchanged if not string)
 * @returns {*} Decoded string or original value
 * 
 * @description Similar to decodeFromSave but also handles:
 * - \\e -> ]]
 * 
 * @example
 * SocialCalc.decodeFromAjax("Hello\\cWorld");     // "Hello:World"
 * SocialCalc.decodeFromAjax("End\\eTag");         // "End]]Tag"
 * SocialCalc.decodeFromAjax("Line1\\nLine2");     // "Line1\nLine2"
 */
SocialCalc.decodeFromAjax = function (s) {
    if (typeof s !== "string") return s;
    if (s.indexOf("\\") === -1) return s; // for performance reasons: replace nothing takes up time

    let r = s.replace(/\\c/g, ":");
    r = r.replace(/\\n/g, "\n");
    r = r.replace(/\\e/g, "]]");
    return r.replace(/\\b/g, "\\");
};

/**
 * @function SocialCalc.encodeForSave
 * @description Encodes strings for save format by escaping special characters
 * @param {*} s - String to encode (returns unchanged if not string)
 * @returns {*} Encoded string or original value
 * 
 * @description Converts special characters to escaped versions:
 * - \\ -> \\b
 * - : -> \\c
 * - newline -> \\n
 * 
 * @example
 * SocialCalc.encodeForSave("Hello:World");    // "Hello\\cWorld"
 * SocialCalc.encodeForSave("Line1\nLine2");   // "Line1\\nLine2"
 * SocialCalc.encodeForSave("Backslash\\Here"); // "Backslash\\bHere"
 */
SocialCalc.encodeForSave = function (s) {
    if (typeof s !== "string") return s;

    // For performance reasons: replace nothing takes up time
    if (s.indexOf("\\") !== -1) {
        s = s.replace(/\\/g, "\\b");
    }
    if (s.indexOf(":") !== -1) {
        s = s.replace(/:/g, "\\c");
    }
    if (s.indexOf("\n") !== -1) {
        s = s.replace(/\n/g, "\\n");
    }
    return s;
};

/**
 * @function SocialCalc.special_chars
 * @description Escapes HTML special characters for safe display
 * @param {string} string - String to escape
 * @returns {string} HTML-escaped string
 * 
 * @description Escapes: &, <, >, " for HTML safety
 * Only performs replacements if special characters are present for performance.
 * 
 * @example
 * SocialCalc.special_chars("<script>alert('xss')</script>");
 * // Returns: "&lt;script&gt;alert('xss')&lt;/script&gt;"
 * 
 * SocialCalc.special_chars('Say "Hello" & goodbye');
 * // Returns: 'Say &quot;Hello&quot; &amp; goodbye'
 */
SocialCalc.special_chars = function (string) {
    if (/[&<>"]/.test(string)) { // only do "slow" replaces if something to replace
        string = string.replace(/&/g, "&amp;");
        string = string.replace(/</g, "&lt;");
        string = string.replace(/>/g, "&gt;");
        string = string.replace(/"/g, "&quot;");
    }
    return string;
};

/**
 * @function SocialCalc.Lookup
 * @description Performs binary-style lookup in a sorted array
 * @param {*} value - Value to find in the list
 * @param {Array} list - Sorted array to search
 * @returns {number|null} Index of largest value <= input value, or null if none found
 * 
 * @description Finds the index of the largest value in the list that is <= value.
 * Used for lookups in sorted arrays like tax tables or grade boundaries.
 * 
 * @example
 * const grades = [60, 70, 80, 90];
 * SocialCalc.Lookup(75, grades); // Returns 1 (70 is largest <= 75)
 * SocialCalc.Lookup(95, grades); // Returns 3 (all values are smaller)
 * SocialCalc.Lookup(50, grades); // Returns null (no values <= 50)
 */
SocialCalc.Lookup = function (value, list) {
    for (let i = 0; i < list.length; i++) {
        if (list[i] > value) {
            if (i > 0) return i - 1;
            else return null;
        }
    }
    return list.length - 1; // if all smaller, matches last
};

/**
 * @function SocialCalc.setStyles
 * @description Sets CSS styles on an element from a CSS text string
 * @param {HTMLElement} element - DOM element to apply styles to
 * @param {string|null} cssText - CSS style string (e.g., "color:red;font-size:12px")
 * 
 * @description Takes a CSS style string and applies each style property to the element.
 * Handles pseudo style strings where property names may use camelCase (e.g., textAlign).
 * Safe to call with null cssText.
 * 
 * @example
 * const div = document.createElement('div');
 * SocialCalc.setStyles(div, "color:red;fontSize:12px;textAlign:center");
 * // Sets div.style.color = "red", div.style.fontSize = "12px", etc.
 */
SocialCalc.setStyles = function (element, cssText) {
    if (!cssText) return;

    const parts = cssText.split(";");
    for (let part = 0; part < parts.length; part++) {
        const pos = parts[part].indexOf(":"); // find first colon (could be one in url)
        if (pos !== -1) {
            const name = parts[part].substring(0, pos);
            const value = parts[part].substring(pos + 1);
            if (name && value) { // if non-null name and value, set style
                element.style[name] = value;
            }
        }
    }
};

/**
 * @function SocialCalc.GetViewportInfo
 * @description Gets viewport dimensions and scroll positions
 * @returns {Object} Viewport information object
 * @returns {number} returns.width - Viewport width in pixels
 * @returns {number} returns.height - Viewport height in pixels  
 * @returns {number} returns.horizontalScroll - Horizontal scroll offset in pixels
 * @returns {number} returns.verticalScroll - Vertical scroll offset in pixels
 * 
 * @description Cross-browser compatible viewport information getter.
 * Based on Flanagan, JavaScript, 5th Edition, page 276.
 * 
 * @example
 * const viewport = SocialCalc.GetViewportInfo();
 * console.log(`Viewport: ${viewport.width}x${viewport.height}`);
 * console.log(`Scroll: ${viewport.horizontalScroll}, ${viewport.verticalScroll}`);
 */
SocialCalc.GetViewportInfo = function () {
    const result = {};

    if (window.innerWidth) { // all but IE
        result.width = window.innerWidth;
        result.height = window.innerHeight;
        result.horizontalScroll = window.pageXOffset;
        result.verticalScroll = window.pageYOffset;
    } else {
        if (document.documentElement && document.documentElement.clientWidth) {
            result.width = document.documentElement.clientWidth;
            result.height = document.documentElement.clientHeight;
            result.horizontalScroll = document.documentElement.scrollLeft;
            result.verticalScroll = document.documentElement.scrollTop;
        } else if (document.body.clientWidth) {
            result.width = document.body.clientWidth;
            result.height = document.body.clientHeight;
            result.horizontalScroll = document.body.scrollLeft;
            result.verticalScroll = document.body.scrollTop;
        }
    }

    return result;
};

/**
 * @function SocialCalc.GetElementPosition
 * @description Gets the absolute position of an element in the document
 * @param {HTMLElement} element - DOM element to get position for
 * @returns {Object} Position object
 * @returns {number} returns.left - Left offset from document edge in pixels
 * @returns {number} returns.top - Top offset from document edge in pixels
 * 
 * @description Based on Goodman's JavaScript & DHTML Cookbook, 2nd Edition, page 415.
 * Does not account for scroll positions.
 * 
 * @example
 * const pos = SocialCalc.GetElementPosition(document.getElementById('myDiv'));
 * console.log(`Element at: ${pos.left}, ${pos.top}`);
 */
SocialCalc.GetElementPosition = function (element) {
    let offsetLeft = 0;
    let offsetTop = 0;

    while (element) {
        offsetLeft += element.offsetLeft;
        offsetTop += element.offsetTop;
        element = element.offsetParent;
    }

    return { left: offsetLeft, top: offsetTop };
};

/**
 * @function SocialCalc.GetElementPositionWithScroll
 * @description Gets element position accounting for scroll offsets throughout the DOM tree
 * @param {HTMLElement} element - DOM element to get position for
 * @returns {Object} Position object
 * @returns {number} returns.left - Left position adjusted for scroll offsets
 * @returns {number} returns.top - Top position adjusted for scroll offsets
 * 
 * @description More accurate than GetElementPosition as it accounts for scroll
 * offsets by traversing the entire DOM tree up to the HTML element.
 * 
 * @example
 * const pos = SocialCalc.GetElementPositionWithScroll(scrolledElement);
 * console.log(`True position: ${pos.left}, ${pos.top}`);
 */
SocialCalc.GetElementPositionWithScroll = function (element) {
    let offsetLeft = 0;
    let offsetTop = 0;
    let offsetElement = element;

    while (element) {
        if (element.tagName === "HTML") break;

        if (element === offsetElement) {
            offsetLeft += element.offsetLeft;
            offsetTop += element.offsetTop;
            offsetElement = element.offsetParent;
        }

        if (element.scrollLeft) {
            offsetLeft -= element.scrollLeft;
        }
        if (element.scrollTop) {
            offsetTop -= element.scrollTop;
        }

        element = element.parentNode;
    }

    return { left: offsetLeft, top: offsetTop };
};

/**
 * @function SocialCalc.LookupElement
 * @description Finds an array element that contains a specific DOM element
 * @param {HTMLElement} element - DOM element to find
 * @param {Array<Object>} array - Array of objects with "element" property
 * @returns {Object|null} Array element containing the DOM element, or null if not found
 * 
 * @example
 * const items = [
 *   {element: div1, data: "first"},
 *   {element: div2, data: "second"}
 * ];
 * const found = SocialCalc.LookupElement(div1, items);
 * // Returns: {element: div1, data: "first"}
 */
SocialCalc.LookupElement = function (element, array) {
    for (let i = 0; i < array.length; i++) {
        if (array[i].element === element) return array[i];
    }
    return null;
};

/**
 * @function SocialCalc.AssignID
 * @description Optionally assigns an ID with a prefix to a DOM element
 * @param {Object} obj - Object that may have an idPrefix property
 * @param {HTMLElement} element - DOM element to assign ID to
 * @param {string} id - ID suffix to append to prefix
 * 
 * @description Only assigns ID if obj has a non-empty idPrefix attribute.
 * Final ID will be obj.idPrefix + id.
 * 
 * @example
 * const manager = {idPrefix: "sheet_"};
 * SocialCalc.AssignID(manager, div, "cell_A1");
 * // Sets div.id = "sheet_cell_A1"
 * 
 * const noPrefix = {};
 * SocialCalc.AssignID(noPrefix, div, "test");
 * // Does nothing (no idPrefix)
 */
SocialCalc.AssignID = function (obj, element, id) {
    if (obj.idPrefix) { // Object must have a non-empty idPrefix attribute
        element.id = obj.idPrefix + id;
    }
};

/**
 * @function SocialCalc.GetCellContents
 * @description Gets the contents of a cell with appropriate prefix
 * @param {SocialCalc.Sheet} sheetobj - Sheet containing the cell
 * @param {string} coord - Cell coordinate (e.g., "A1")
 * @returns {string} Cell contents with prefix ("'", "=", etc.) or empty string
 * 
 * @description Returns cell contents formatted for editing:
 * - Values: returned as-is
 * - Text: prefixed with '
 * - Formulas: prefixed with =
 * - Constants: returned as formula without prefix
 * 
 * @example
 * SocialCalc.GetCellContents(sheet, "A1"); 
 * // For value 42: "42"
 * // For text "Hello": "'Hello" 
 * // For formula SUM(B1:B5): "=SUM(B1:B5)"
 * // For constant 3.14159: "3.14159"
 */
SocialCalc.GetCellContents = function (sheetobj, coord) {
    let result = "";
    const cellobj = sheetobj.cells[coord];

    if (cellobj) {
        switch (cellobj.datatype) {
            case "v":
                result = cellobj.datavalue + "";
                break;
            case "t":
                result = "'" + cellobj.datavalue;
                break;
            case "f":
                result = "=" + cellobj.formula;
                break;
            case "c":
                result = cellobj.formula;
                break;
            default:
                break;
        }
    }

    return result;
};

/**
 * @function SocialCalc.FormatValueForDisplay
 * @description Formats a cell value for display in HTML
 * @param {SocialCalc.Sheet} sheetobj - Sheet containing the cell
 * @param {*} value - Cell value (numeric or text)
 * @param {string} cr - Cell coordinate
 * @param {Object|null} linkstyle - Link styling options for wiki-text expansion
 * @returns {string} HTML-formatted display value
 * 
 * @description Converts cell values to HTML display format based on cell properties.
 * Handles different value types (text, numeric, error) and applies appropriate
 * formatting based on cell's value format settings. Special formats like "formula"
 * and "forcetext" are handled appropriately.
 * 
 * @example
 * // Format a numeric value
 * const display1 = SocialCalc.FormatValueForDisplay(sheet, 42.5, "A1", null);
 * 
 * // Format text with wiki markup
 * const display2 = SocialCalc.FormatValueForDisplay(sheet, "Hello World", "B1", {type: "wiki"});
 * 
 * // Format formula display
 * const display3 = SocialCalc.FormatValueForDisplay(sheet, "=SUM(A1:A5)", "C1", null);
 */
SocialCalc.FormatValueForDisplay = function (sheetobj, value, cr, linkstyle) {
    let valueformat, valuetype, valuesubtype;
    let displayvalue;

    const sheetattribs = sheetobj.attribs;
    const scc = SocialCalc.Constants;

    let cell = sheetobj.cells[cr];

    if (!cell) { // get an empty cell if not there
        cell = new SocialCalc.Cell(cr);
    }

    displayvalue = value;

    valuetype = cell.valuetype || ""; // get type of value to determine formatting
    valuesubtype = valuetype.substring(1);
    valuetype = valuetype.charAt(0);

    // Handle errors
    if (cell.errors || valuetype === "e") {
        displayvalue = cell.errors || valuesubtype || "Error in cell";
        return displayvalue;
    }

    // Handle text values
    if (valuetype === "t") {
        valueformat = sheetobj.valueformats[cell.textvalueformat - 0] ||
            sheetobj.valueformats[sheetattribs.defaulttextvalueformat - 0] || "";

        if (valueformat === "formula") {
            if (cell.datatype === "f") {
                displayvalue = SocialCalc.special_chars("=" + cell.formula) || "&nbsp;";
            } else if (cell.datatype === "c") {
                displayvalue = SocialCalc.special_chars("'" + cell.formula) || "&nbsp;";
            } else {
                displayvalue = SocialCalc.special_chars("'" + displayvalue) || "&nbsp;";
            }
            return displayvalue;
        }

        displayvalue = SocialCalc.format_text_for_display(displayvalue, cell.valuetype, valueformat, sheetobj, linkstyle);
        if (valueformat === "text-html") {
            SocialCalc.ScriptCheck(sheetobj.sheetid, cr, value);
        }
    }
    // Handle numeric values
    else if (valuetype === "n") {
        valueformat = cell.nontextvalueformat;
        if (valueformat == null || valueformat === "") {
            valueformat = sheetattribs.defaultnontextvalueformat;
        }
        valueformat = sheetobj.valueformats[valueformat - 0];
        if (valueformat == null || valueformat === "none") {
            valueformat = "";
        }

        if (valueformat === "formula") {
            if (cell.datatype === "f") {
                displayvalue = SocialCalc.special_chars("=" + cell.formula) || "&nbsp;";
            } else if (cell.datatype === "c") {
                displayvalue = SocialCalc.special_chars("'" + cell.formula) || "&nbsp;";
            } else {
                displayvalue = SocialCalc.special_chars("'" + displayvalue) || "&nbsp;";
            }
            return displayvalue;
        } else if (valueformat === "forcetext") {
            if (cell.datatype === "f") {
                displayvalue = SocialCalc.special_chars("=" + cell.formula) || "&nbsp;";
            } else if (cell.datatype === "c") {
                displayvalue = SocialCalc.special_chars(cell.formula) || "&nbsp;";
            } else {
                displayvalue = SocialCalc.special_chars(displayvalue) || "&nbsp;";
            }
            return displayvalue;
        }

        displayvalue = SocialCalc.format_number_for_display(displayvalue, cell.valuetype, valueformat);
    }
    // Handle unknown/blank values
    else {
        displayvalue = "&nbsp;";
    }

    return displayvalue;
};
/**
 * @function SocialCalc.format_text_for_display
 * @description Formats text values for display with various text formatting options
 * @param {*} rawvalue - The raw text value to format
 * @param {string} valuetype - The value type string (e.g., "th", "tw", "tl")
 * @param {string} valueformat - The format specification
 * @param {SocialCalc.Sheet} sheetobj - The sheet object for context
 * @param {Object|null} linkstyle - Link styling options
 * @returns {string} HTML-formatted display value
 * 
 * @description Handles various text formatting types including:
 * - text-html: Raw HTML output
 * - text-wiki: Wiki markup processing
 * - text-url: URL links
 * - text-link: Extended link capabilities
 * - text-image: Image tags
 * - text-custom: Custom format with placeholders (@r, @s, @u)
 * - hidden: Hidden content
 * - Default: Plain text with HTML escaping
 * 
 * @example
 * // Format as HTML
 * const html = SocialCalc.format_text_for_display("<b>Bold</b>", "th", "text-html", sheet, null);
 * 
 * // Format as URL
 * const url = SocialCalc.format_text_for_display("http://example.com", "tl", "text-url", sheet, null);
 * 
 * // Custom format
 * const custom = SocialCalc.format_text_for_display("Hello", "t", "text-custom:@s World", sheet, null);
 */
SocialCalc.format_text_for_display = function (rawvalue, valuetype, valueformat, sheetobj, linkstyle) {
    let valuesubtype, dvsc, dvue, textval;
    let displayvalue;

    valuesubtype = valuetype.substring(1);
    displayvalue = rawvalue;

    // Normalize valueformat
    if (valueformat === "none" || valueformat == null) valueformat = "";
    if (!/^(text-|custom|hidden)/.test(valueformat)) valueformat = "";

    // Determine format from type if not specified
    if (valueformat === "" || valueformat === "General") {
        if (valuesubtype === "h") valueformat = "text-html";
        if (valuesubtype === "w" || valuesubtype === "r") valueformat = "text-wiki";
        if (valuesubtype === "l") valueformat = "text-link";
        if (!valuesubtype) valueformat = "text-plain";
    }

    // Process based on format type
    if (valueformat === "text-html") {
        // HTML - output as it is
        ;
    }
    else if (SocialCalc.Callbacks.expand_wiki && /^text-wiki/.test(valueformat)) {
        // Do general wiki markup
        displayvalue = SocialCalc.Callbacks.expand_wiki(displayvalue, sheetobj, linkstyle, valueformat);
    }
    else if (valueformat === "text-wiki") {
        // Wiki-text - encode then output (or use old markup callback)
        displayvalue = (SocialCalc.Callbacks.expand_markup &&
            SocialCalc.Callbacks.expand_markup(displayvalue, sheetobj, linkstyle)) ||
            SocialCalc.special_chars(displayvalue);
    }
    else if (valueformat === "text-url") {
        // Text is a URL for a link
        dvsc = SocialCalc.special_chars(displayvalue);
        dvue = encodeURI(displayvalue);
        displayvalue = `<a href="${dvue}">${dvsc}</a>`;
    }
    else if (valueformat === "text-link") {
        // More extensive link capabilities for regular web links
        displayvalue = SocialCalc.expand_text_link(displayvalue, sheetobj, linkstyle, valueformat);
    }
    else if (valueformat === "text-image") {
        // Text is a URL for an image
        dvue = encodeURI(displayvalue);
        displayvalue = `<img src="${dvue}">`;
    }
    else if (valueformat.substring(0, 12) === "text-custom:") {
        // Custom text format: @r = text raw, @s = special chars, @u = url encoded
        dvsc = SocialCalc.special_chars(displayvalue);
        dvsc = dvsc.replace(/  /g, "&nbsp; "); // keep multiple spaces
        dvsc = dvsc.replace(/\n/g, "<br>");   // keep line breaks
        dvue = encodeURI(displayvalue);

        textval = {};
        textval.r = displayvalue;
        textval.s = dvsc;
        textval.u = dvue;

        displayvalue = valueformat.substring(12); // remove "text-custom:"
        displayvalue = displayvalue.replace(/@(r|s|u)/g, (a, c) => textval[c]); // replace placeholders
    }
    else if (valueformat.substring(0, 6) === "custom") {
        // Custom format
        displayvalue = SocialCalc.special_chars(displayvalue);
        displayvalue = displayvalue.replace(/  /g, "&nbsp; "); // keep multiple spaces
        displayvalue = displayvalue.replace(/\n/g, "<br>");    // keep line breaks
        displayvalue += " (custom format)";
    }
    else if (valueformat === "hidden") {
        displayvalue = "&nbsp;";
    }
    else {
        // Plain text
        displayvalue = SocialCalc.special_chars(displayvalue);
        displayvalue = displayvalue.replace(/  /g, "&nbsp; "); // keep multiple spaces
        displayvalue = displayvalue.replace(/\n/g, "<br>");    // keep line breaks
    }

    return displayvalue;
};

/**
 * @function SocialCalc.format_number_for_display
 * @description Formats numeric values for display with various number formatting options
 * @param {*} rawvalue - The raw numeric value to format
 * @param {string} valuetype - The value type string (e.g., "n", "n%", "n$")
 * @param {string} valueformat - The format specification
 * @returns {string} Formatted display value
 * 
 * @description Handles various numeric formatting types including:
 * - Auto format based on value subtype (%, $, dates, times, logical)
 * - Custom number formats using FormatNumber.formatNumberWithFormat
 * - Special formats like "logical" and "hidden"
 * 
 * @example
 * // Format percentage
 * const pct = SocialCalc.format_number_for_display(0.25, "n%", "Auto");
 * // Returns: "25.0%"
 * 
 * // Format currency
 * const money = SocialCalc.format_number_for_display(1234.56, "n$", "Auto");
 * // Returns: "$1,234.56"
 * 
 * // Custom format
 * const custom = SocialCalc.format_number_for_display(42.5, "n", "#,##0.00");
 * // Returns: "42.50"
 */
SocialCalc.format_number_for_display = function (rawvalue, valuetype, valueformat) {
    let value, valuesubtype;
    const scc = SocialCalc.Constants;

    value = rawvalue - 0;
    valuesubtype = valuetype.substring(1);

    // Handle default formats based on subtype
    if (valueformat === "Auto" || valueformat === "") {
        if (valuesubtype === "%") {
            valueformat = "#,##0.0%";
        }
        else if (valuesubtype === '$') {
            valueformat = '[$]#,##0.00';
        }
        else if (valuesubtype === 'dt') {
            valueformat = scc.defaultFormatdt;
        }
        else if (valuesubtype === 'd') {
            valueformat = scc.defaultFormatd;
        }
        else if (valuesubtype === 't') {
            valueformat = scc.defaultFormatt;
        }
        else if (valuesubtype === 'l') {
            valueformat = 'logical';
        }
        else {
            valueformat = "General";
        }
    }

    // Handle special formats
    if (valueformat === "logical") {
        return value ? scc.defaultDisplayTRUE : scc.defaultDisplayFALSE;
    }

    if (valueformat === "hidden") {
        return "&nbsp;";
    }

    // Use standard number formatting
    return SocialCalc.FormatNumber.formatNumberWithFormat(rawvalue, valueformat, "");
};

/**
 * @function SocialCalc.DetermineValueType
 * @description Analyzes a raw value to determine its type and convert it appropriately
 * @param {*} rawvalue - The raw input value to analyze
 * @returns {Object} Object containing the converted value and its type
 * @returns {*} returns.value - The converted value (number or string)
 * @returns {string} returns.type - The determined type ("t", "n", "n%", "n$", "nd", "nt", "tl", etc.)
 * 
 * @description Tries to follow spreadsheet VALUE() function spec. Recognizes:
 * - Numbers (including scientific notation)
 * - Percentages (15.1%)
 * - Currency ($1.49)
 * - Numbers with commas (1,234.56)
 * - Dates (MM/DD/YYYY, YYYY-MM-DD)
 * - Times (HH:MM, HH:MM:SS)
 * - Fractions (1 1/2)
 * - Constants (TRUE, FALSE, #N/A, etc.)
 * - URLs (http://...)
 * 
 * @example
 * SocialCalc.DetermineValueType("42.5");        // {value: 42.5, type: "n"}
 * SocialCalc.DetermineValueType("25%");         // {value: 0.25, type: "n%"}
 * SocialCalc.DetermineValueType("$1,234.56");   // {value: 1234.56, type: "n$"}
 * SocialCalc.DetermineValueType("12/25/2023");  // {value: [julian_date], type: "nd"}
 * SocialCalc.DetermineValueType("TRUE");        // {value: 1, type: "nl"}
 * SocialCalc.DetermineValueType("http://example.com"); // {value: "http://example.com", type: "tl"}
 */
SocialCalc.DetermineValueType = function (rawvalue) {
    let value = rawvalue + "";
    let type = "t";
    let tvalue, matches, year, hour, minute, second, denom, num, intgr, constr;

    // Remove leading and trailing blanks
    tvalue = value.replace(/^\s+/, "").replace(/\s+$/, "");

    if (value.length === 0) {
        type = "";
    }
    else if (value.match(/^\s+$/)) {
        // Just blanks - leave type "t"
        ;
    }
    else if (tvalue.match(/^[-+]?\d*(?:\.)?\d*(?:[eE][-+]?\d+)?$/)) {
        // General number, including scientific notation
        value = tvalue - 0;
        if (isNaN(value)) {
            // Leave alone - catches things like plain "-"
            value = rawvalue + "";
        } else {
            type = "n";
        }
    }
    else if (tvalue.match(/^[-+]?\d*(?:\.)?\d*\s*%$/)) {
        // Percent form: 15.1%
        value = (tvalue.slice(0, -1) - 0) / 100;
        type = "n%";
    }
    else if (tvalue.match(/^[-+]?\$\s*\d*(?:\.)?\d*\s*$/) && tvalue.match(/\d/)) {
        // Dollar format: $1.49
        value = tvalue.replace(/\$/, "") - 0;
        type = "n$";
    }
    else if (tvalue.match(/^[-+]?(\d*,\d*)+(?:\.)?\d*$/)) {
        // Number format ignoring commas: 1,234.49
        value = tvalue.replace(/,/g, "") - 0;
        type = "n";
    }
    else if (tvalue.match(/^[-+]?(\d*,\d*)+(?:\.)?\d*\s*%$/)) {
        // Percent with commas: 1,234.49%
        value = (tvalue.replace(/[%,]/g, "") - 0) / 100;
        type = "n%";
    }
    else if (tvalue.match(/^[-+]?\$\s*(\d*,\d*)+(?:\.)?\d*$/) && tvalue.match(/\d/)) {
        // Dollar and commas: $1,234.49
        value = tvalue.replace(/[\$,]/g, "") - 0;
        type = "n$";
    }
    else if (matches = value.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{1,4})\s*$/)) {
        // MM/DD/YYYY, MM-DD-YYYY
        year = matches[3] - 0;
        year = year < 1000 ? year + 2000 : year;
        value = SocialCalc.FormatNumber.convert_date_gregorian_to_julian(year, matches[1] - 0, matches[2] - 0) - 2415019;
        type = "nd";
    }
    else if (matches = value.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\s*$/)) {
        // YYYY-MM-DD, YYYY/MM/DD
        year = matches[1] - 0;
        year = year < 1000 ? year + 2000 : year;
        value = SocialCalc.FormatNumber.convert_date_gregorian_to_julian(year, matches[2] - 0, matches[3] - 0) - 2415019;
        type = "nd";
    }
    else if (matches = value.match(/^(\d{1,2}):(\d{1,2})\s*$/)) {
        // HH:MM
        hour = matches[1] - 0;
        minute = matches[2] - 0;
        if (hour < 24 && minute < 60) {
            value = hour / 24 + minute / (24 * 60);
            type = "nt";
        }
    }
    else if (matches = value.match(/^(\d{1,2}):(\d{1,2}):(\d{1,2})\s*$/)) {
        // HH:MM:SS
        hour = matches[1] - 0;
        minute = matches[2] - 0;
        second = matches[3] - 0;
        if (hour < 24 && minute < 60 && second < 60) {
            value = hour / 24 + minute / (24 * 60) + second / (24 * 60 * 60);
            type = "nt";
        }
    }
    else if (matches = value.match(/^\s*([-+]?\d+) (\d+)\/(\d+)\s*$/)) {
        // Mixed fraction: 1 1/2
        intgr = matches[1] - 0;
        num = matches[2] - 0;
        denom = matches[3] - 0;
        if (denom && denom > 0) {
            value = intgr + (intgr < 0 ? -num / denom : num / denom);
            type = "n";
        }
    }
    else if (constr = SocialCalc.InputConstants[value.toUpperCase()]) {
        // Special constants like "FALSE", "#N/A"
        num = constr.indexOf(",");
        value = constr.substring(0, num) - 0;
        type = constr.substring(num + 1);
    }
    else if (tvalue.length > 7 && tvalue.substring(0, 7).toLowerCase() === "http://") {
        // URL
        value = tvalue;
        type = "tl";
    }

    return { value: value, type: type };
};

/**
 * @constant {Object} SocialCalc.InputConstants
 * @description Constants that convert string inputs to specific values and types
 * @description Format: "STRING": "value,type" where value is numeric and type is the value type
 * 
 * @example
 * // These constants are recognized by DetermineValueType:
 * // "TRUE" -> {value: 1, type: "nl"}
 * // "FALSE" -> {value: 0, type: "nl"}  
 * // "#N/A" -> {value: 0, type: "e#N/A"}
 */
SocialCalc.InputConstants = {
    "TRUE": "1,nl",
    "FALSE": "0,nl",
    "#N/A": "0,e#N/A",
    "#NULL!": "0,e#NULL!",
    "#NUM!": "0,e#NUM!",
    "#DIV/0!": "0,e#DIV/0!",
    "#VALUE!": "0,e#VALUE!",
    "#REF!": "0,e#REF!",
    "#NAME?": "0,e#NAME?"
};

/**
 * @function SocialCalc.default_expand_markup
 * @description Default wiki-text processor (placeholder implementation)
 * @param {string} displayvalue - The text value to process
 * @param {SocialCalc.Sheet} sheetobj - Sheet object for context
 * @param {Object|null} linkstyle - Link styling options
 * @returns {string} Processed HTML text
 * 
 * @description This is a basic placeholder implementation. Applications should
 * replace the reference to this function with their own wiki markup processor.
 * Currently only handles HTML escaping, space preservation, and line breaks.
 * 
 * @example
 * const processed = SocialCalc.default_expand_markup("Hello\nWorld  Test", sheet, null);
 * // Returns: "Hello<br>World&nbsp; Test"
 */
SocialCalc.default_expand_markup = function (displayvalue, sheetobj, linkstyle) {
    let result = displayvalue;

    result = SocialCalc.special_chars(result); // HTML escape special chars
    result = result.replace(/  /g, "&nbsp; "); // preserve multiple spaces
    result = result.replace(/\n/g, "<br>");    // preserve line breaks

    return result; // Very basic processing by default

    // Commented out more advanced wiki markup that could be implemented:
    // result = result.replace(/('*)'''(.*?)'''/g, "$1<b>$2<\/b>"); // Wiki-style bold
    // result = result.replace(/''(.*?)''/g, "<i>$1<\/i>");         // Wiki-style italics

    // return result;
};

/**
 * @function SocialCalc.expand_text_link
 * @description Parses and formats link text into HTML anchor tags
 * @param {string} displayvalue - The link text to parse and format
 * @param {SocialCalc.Sheet} sheetobj - Sheet object for context
 * @param {Object|null} linkstyle - Link styling options
 * @param {string} valueformat - The value format specification
 * @returns {string} HTML anchor tag
 * 
 * @description Handles various link formats and generates appropriate HTML.
 * Uses ParseCellLinkText to extract URL, description, and other link properties.
 * Supports both regular web links and page links via callbacks.
 * 
 * @example
 * // Simple URL
 * const link1 = SocialCalc.expand_text_link("http://example.com", sheet, null, "text-link");
 * 
 * // URL with description
 * const link2 = SocialCalc.expand_text_link("My Site<http://example.com>", sheet, null, "text-link");
 * 
 * // Page link
 * const link3 = SocialCalc.expand_text_link("Home Page[HomePage]", sheet, null, "text-link");
 */
SocialCalc.expand_text_link = function (displayvalue, sheetobj, linkstyle, valueformat) {
    let desc, tb, str;
    const scc = SocialCalc.Constants;
    let url = "";

    const parts = SocialCalc.ParseCellLinkText(displayvalue + "");

    // Determine description
    if (parts.desc) {
        desc = SocialCalc.special_chars(parts.desc);
    } else {
        desc = parts.pagename ? scc.defaultPageLinkFormatString : scc.defaultLinkFormatString;
    }

    // Remove http:// from description unless explicit
    if (displayvalue.length > 7 &&
        displayvalue.substring(0, 7).toLowerCase() === "http://" &&
        displayvalue.charAt(displayvalue.length - 1) !== ">") {
        desc = desc.substring(7);
    }

    // Set target attribute
    tb = (parts.newwin || !linkstyle) ? ' target="_blank"' : "";

    // Generate URL
    if (parts.pagename) {
        if (SocialCalc.Callbacks.MakePageLink) {
            url = SocialCalc.Callbacks.MakePageLink(parts.pagename, parts.workspacename, linkstyle, valueformat);
        }
        // Alternative URL generation commented out:
        // else if (parts.workspace) {
        //     url = "/" + encodeURI(parts.workspace) + "/" + encodeURI(parts.pagename);
        // }
        // else {
        //     url = parts.pagename;
        // }
    } else {
        url = encodeURI(parts.url);
    }

    str = `<a href="${url}"${tb}>${desc}</a>`;
    return str;
};

/**
 * @function SocialCalc.ParseCellLinkText
 * @description Parses various link text formats into component parts
 * @param {string} str - The link text string to parse
 * @returns {Object} Parsed link components
 * @returns {string} returns.url - The URL (for web links)
 * @returns {string} returns.desc - The description text
 * @returns {boolean} returns.newwin - Whether to open in new window
 * @returns {string} returns.pagename - The page name (for page links)
 * @returns {string} returns.workspace - The workspace name (for page links)
 * 
 * @description Supports multiple link formats:
 * 
 * Web links:
 * - url
 * - <url>
 * - desc<url>
 * - "desc"<url>
 * - <<url>> (opens in new window)
 * 
 * Page links:
 * - [page name]
 * - "desc"[page name]
 * - desc[page name]
 * - {workspace [page name]}
 * - "desc"{workspace [page name]}
 * - [[page name]] (opens in new window)
 * 
 * @example
 * // Parse simple URL
 * SocialCalc.ParseCellLinkText("http://example.com");
 * // Returns: {url: "http://example.com", desc: "", newwin: false, pagename: "", workspace: ""}
 * 
 * // Parse URL with description
 * SocialCalc.ParseCellLinkText("My Site<http://example.com>");
 * // Returns: {url: "http://example.com", desc: "My Site", newwin: false, pagename: "", workspace: ""}
 * 
 * // Parse page link with workspace
 * SocialCalc.ParseCellLinkText("Home{MySpace [HomePage]}");
 * // Returns: {url: "", desc: "Home", newwin: false, pagename: "HomePage", workspace: "MySpace"}
 */
SocialCalc.ParseCellLinkText = function (str) {
    const result = { url: "", desc: "", newwin: false, pagename: "", workspace: "" };

    let pageform = false;
    let urlend = str.length - 1;
    let descstart = 0;
    const lastlt = str.lastIndexOf("<");
    const lastbrkt = str.lastIndexOf("[");
    const lastbrace = str.lastIndexOf("{");
    let descend = -1;
    let wsend;

    // Determine if this is a plain URL or has markup
    if ((str.charAt(urlend) !== ">" || lastlt === -1) &&
        (str.charAt(urlend) !== "]" || lastbrkt === -1) &&
        (str.charAt(urlend) !== "}" || str.charAt(urlend - 1) !== "]" ||
            lastbrace === -1 || lastbrkt === -1 || lastbrkt < lastbrace)) {
        // Plain URL
        urlend++;
        descend = urlend;
    } else {
        // Some markup present
        if (str.charAt(urlend) === ">") {
            // URL form: <url> or <<url>>
            descend = lastlt - 1;
            if (lastlt > 0 && str.charAt(descend) === "<" && str.charAt(urlend - 1) === ">") {
                descend--;
                urlend--;
                result.newwin = true;
            }
        }
        else if (str.charAt(urlend) === "]") {
            // Plain page form: [page] or [[page]]
            descend = lastbrkt - 1;
            pageform = true;
            if (lastbrkt > 0 && str.charAt(descend) === "[" && str.charAt(urlend - 1) === "]") {
                descend--;
                urlend--;
                result.newwin = true;
            }
        }
        else if (str.charAt(urlend) === "}") {
            // Page and workspace form: {workspace [page]}
            descend = lastbrace - 1;
            pageform = true;
            wsend = lastbrkt;
            urlend--;
            if (lastbrkt > 0 && str.charAt(lastbrkt - 1) === "[" && str.charAt(urlend - 1) === "]") {
                wsend = lastbrkt - 1;
                urlend--;
                result.newwin = true;
            }
            // Trim trailing space in workspace name
            if (str.charAt(wsend - 1) === " ") {
                wsend--;
            }
            result.workspace = str.substring(lastbrace + 1, wsend) || "";
        }

        // Trim trailing space on description
        if (str.charAt(descend) === " ") {
            descend--;
        }

        // Handle quoted descriptions
        if (str.charAt(descstart) === '"' && str.charAt(descend) === '"') {
            descstart++;
            descend--;
        }
    }

    // Extract page name or URL
    if (pageform) {
        result.pagename = str.substring(lastbrkt + 1, urlend) || "";
    } else {
        result.url = str.substring(lastlt + 1, urlend) || "";
    }

    // Extract description
    if (descend >= descstart) {
        result.desc = str.substring(descstart, descend + 1);
    }

    return result;
};
/**
 * @function SocialCalc.ConvertSaveToOtherFormat
 * @description Converts a sheet save string to various output formats
 * @param {string} savestr - The sheet save string to convert
 * @param {string} outputformat - Target format: "scsave", "html", "csv", "tab"
 * @param {boolean} dorecalc - Whether to recalc after loading (OBSOLETE - throws error if true)
 * @returns {string} Converted string in the specified format
 * @throws {Error} Throws error if dorecalc is true (no longer supported)
 * 
 * @description Converts SocialCalc sheet data to different formats:
 * - "scsave": Returns original save string unchanged
 * - "html": Renders as HTML table
 * - "csv": Comma-separated values with proper quoting
 * - "tab": Tab-delimited values with proper quoting
 * 
 * Note: dorecalc parameter is obsolete as of 9/10/08 since recalc is now async.
 * 
 * @example
 * // Convert to HTML
 * const html = SocialCalc.ConvertSaveToOtherFormat(saveData, "html", false);
 * 
 * // Convert to CSV
 * const csv = SocialCalc.ConvertSaveToOtherFormat(saveData, "csv", false);
 * 
 * // Convert to tab-delimited
 * const tsv = SocialCalc.ConvertSaveToOtherFormat(saveData, "tab", false);
 */
SocialCalc.ConvertSaveToOtherFormat = function (savestr, outputformat, dorecalc) {
    let sheet, context, clipextents, div, ele, row, col, cr, cell, str;
    let result = "";

    // Return unchanged for scsave format
    if (outputformat === "scsave") {
        return savestr;
    }

    // Return empty for empty input
    if (savestr === "") {
        return "";
    }

    // Parse the sheet data
    sheet = new SocialCalc.Sheet();
    sheet.ParseSheetSave(savestr);

    // Recalc is no longer supported (async issues)
    if (dorecalc) {
        throw new Error("SocialCalc.ConvertSaveToOtherFormat: Not doing recalc.");
    }

    // Determine the range to convert
    if (sheet.copiedfrom) {
        clipextents = SocialCalc.ParseRange(sheet.copiedfrom);
    } else {
        clipextents = {
            cr1: { row: 1, col: 1 },
            cr2: { row: sheet.attribs.lastrow, col: sheet.attribs.lastcol }
        };
    }

    // Handle HTML output format
    if (outputformat === "html") {
        context = new SocialCalc.RenderContext(sheet);
        if (sheet.copiedfrom) {
            context.rowpanes[0] = {
                first: clipextents.cr1.row,
                last: clipextents.cr2.row
            };
            context.colpanes[0] = {
                first: clipextents.cr1.col,
                last: clipextents.cr2.col
            };
        }
        div = document.createElement("div");
        ele = context.RenderSheet(null, context.defaultHTMLlinkstyle);
        div.appendChild(ele);

        result = div.innerHTML;

        // Cleanup references
        delete context;
        delete sheet;
        delete ele;
        delete div;

        return result;
    }

    // Handle CSV and tab-delimited formats
    for (row = clipextents.cr1.row; row <= clipextents.cr2.row; row++) {
        for (col = clipextents.cr1.col; col <= clipextents.cr2.col; col++) {
            cr = SocialCalc.crToCoord(col, row);
            cell = sheet.GetAssuredCell(cr);

            // Get cell value as string
            if (cell.errors) {
                str = cell.errors;
            } else {
                str = cell.datavalue + ""; // convert to string
            }

            // Format for CSV
            if (outputformat === "csv") {
                // Handle quotes in CSV
                if (str.indexOf('"') !== -1) {
                    str = str.replace(/"/g, '""'); // double quotes
                }
                // Add quotes if contains comma, space, newline, or quote
                if (/[, \n"]/.test(str)) {
                    str = `"${str}"`; // add quotes
                }
                // Add comma separator (except for first column)
                if (col > clipextents.cr1.col) {
                    str = `,${str}`; // add commas
                }
            }
            // Format for tab-delimited
            else if (outputformat === "tab") {
                // Handle multiline content
                if (str.indexOf('\n') !== -1) { // if multiple lines
                    if (str.indexOf('"') !== -1) {
                        str = str.replace(/"/g, '""'); // double quotes
                    }
                    str = `"${str}"`; // add quotes
                }
                // Add tab separator (except for first column)
                if (col > clipextents.cr1.col) {
                    str = `\t${str}`; // add tabs
                }
            }
            result += str;
        }
        result += "\n";
    }

    return result;
};

/**
 * @function SocialCalc.ConvertOtherFormatToSave
 * @description Converts various input formats to SocialCalc save format
 * @param {string} inputstr - The input string to convert
 * @param {string} inputformat - Source format: "scsave", "csv", "tab"
 * @returns {string} SocialCalc save format string
 * 
 * @description Parses different input formats and converts them to SocialCalc's
 * internal save format:
 * - "scsave": Returns input unchanged
 * - "csv": Parses comma-separated values with quote handling
 * - "tab": Parses tab-delimited values with quote handling
 * 
 * Handles quoted fields, escaped quotes, and multiline content properly.
 * 
 * @example
 * // Convert CSV to save format
 * const saveData = SocialCalc.ConvertOtherFormatToSave(csvString, "csv");
 * 
 * // Convert tab-delimited to save format
 * const saveData2 = SocialCalc.ConvertOtherFormatToSave(tsvString, "tab");
 * 
 * // Pass through save format
 * const saveData3 = SocialCalc.ConvertOtherFormatToSave(saveString, "scsave");
 */
SocialCalc.ConvertOtherFormatToSave = function (inputstr, inputformat) {
    let sheet, lines, i, line, value, inquote, j, ch, row, col, cr, maxc;
    let result = "";

    /**
     * @description Internal helper to add a cell to the sheet
     * @inner
     */
    const AddCell = () => {
        col++;
        if (col > maxc) maxc = col;
        cr = SocialCalc.crToCoord(col, row);
        SocialCalc.SetConvertedCell(sheet, cr, value);
        value = "";
    };

    // Return unchanged for scsave format
    if (inputformat === "scsave") {
        return inputstr;
    }

    sheet = new SocialCalc.Sheet();
    lines = inputstr.split(/\r\n|\n/);
    maxc = 0;

    // Handle CSV format
    if (inputformat === "csv") {
        row = 0;
        inquote = false;
        value = "";

        for (i = 0; i < lines.length; i++) {
            // Ignore extra null line at end
            if (i === lines.length - 1 && lines[i] === "") {
                break;
            }

            if (inquote) {
                // If in quote, continue from where left off with newline
                value += "\n";
            } else {
                // Start new row
                value = "";
                row++;
                col = 0;
            }

            line = lines[i];
            for (j = 0; j < line.length; j++) {
                ch = line.charAt(j);

                if (ch === '"') {
                    if (inquote) {
                        if (j < line.length - 1 && line.charAt(j + 1) === '"') {
                            // Double quotes - add single quote to value
                            j++; // skip the second one
                            value += '"';
                        } else {
                            // End of quoted section
                            inquote = false;
                            if (j === line.length - 1) {
                                AddCell();
                            }
                        }
                    } else {
                        // Start of quoted section
                        inquote = true;
                    }
                    continue;
                }

                if (ch === "," && !inquote) {
                    AddCell();
                } else {
                    value += ch;
                }

                // End of line and not in quote
                if (j === line.length - 1 && !inquote) {
                    AddCell();
                }
            }
        }

        // Set sheet bounds if we have data
        if (maxc > 0) {
            sheet.attribs.lastrow = row;
            sheet.attribs.lastcol = maxc;
            result = sheet.CreateSheetSave(`A1:${SocialCalc.crToCoord(maxc, row)}`);
        }
    }

    // Handle tab-delimited format
    if (inputformat === "tab") {
        row = 0;
        inquote = false;
        value = "";

        for (i = 0; i < lines.length; i++) {
            // Ignore extra null line at end
            if (i === lines.length - 1 && lines[i] === "") {
                break;
            }

            if (inquote) {
                // If in quote, continue from where left off with newline
                value += "\n";
            } else {
                // Start new row
                value = "";
                row++;
                col = 0;
            }

            line = lines[i];
            for (j = 0; j < line.length; j++) {
                ch = line.charAt(j);

                if (ch === '"') {
                    if (inquote) {
                        if (j < line.length - 1) {
                            if (line.charAt(j + 1) === '"') {
                                // Double quotes - add single quote
                                j++; // skip the second one
                                value += '"';
                            } else if (line.charAt(j + 1) === '\t') {
                                // End of quoted item followed by tab
                                j++;
                                inquote = false;
                                AddCell();
                            }
                        } else {
                            // At end of line
                            inquote = false;
                            AddCell();
                        }
                        continue;
                    }
                    // Quote at start of item
                    if (value === "") {
                        inquote = true;
                        continue;
                    }
                }

                if (ch === "\t" && !inquote) {
                    AddCell();
                } else {
                    value += ch;
                }

                // End of line and not in quote
                if (j === line.length - 1 && !inquote) {
                    AddCell();
                }
            }
        }

        // Set sheet bounds if we have data
        if (maxc > 0) {
            sheet.attribs.lastrow = row;
            sheet.attribs.lastcol = maxc;
            result = sheet.CreateSheetSave(`A1:${SocialCalc.crToCoord(maxc, row)}`);
        }
    }

    return result;
};

/**
 * @function SocialCalc.SetConvertedCell
 * @description Sets a cell with a value and type determined from raw input
 * @param {SocialCalc.Sheet} sheet - The sheet to modify
 * @param {string} cr - Cell coordinate (e.g., "A1")
 * @param {string} rawvalue - Raw string value to analyze and convert
 * 
 * @description Analyzes the raw value using DetermineValueType and sets the cell
 * with appropriate datatype, valuetype, and datavalue. Handles three cases:
 * - Pure numbers: stored as value type
 * - Text: stored as text type
 * - Special numbers (%, $, dates): stored as constant with original formula
 * 
 * @example
 * // Set a numeric value
 * SocialCalc.SetConvertedCell(sheet, "A1", "42.5");
 * // Cell becomes: datatype="v", valuetype="n", datavalue=42.5
 * 
 * // Set a percentage
 * SocialCalc.SetConvertedCell(sheet, "B1", "25%");
 * // Cell becomes: datatype="c", valuetype="n%", datavalue=0.25, formula="25%"
 * 
 * // Set text
 * SocialCalc.SetConvertedCell(sheet, "C1", "Hello World");
 * // Cell becomes: datatype="t", valuetype="t", datavalue="Hello World"
 */
SocialCalc.SetConvertedCell = function (sheet, cr, rawvalue) {
    const cell = sheet.GetAssuredCell(cr);
    const value = SocialCalc.DetermineValueType(rawvalue);

    // Check if it's a pure number that doesn't need "constant" to remember original value
    if (value.type === 'n' && value.value == rawvalue) {
        cell.datatype = "v";
        cell.valuetype = "n";
        cell.datavalue = value.value;
    }
    // Text of some sort but left unchanged
    else if (value.type.charAt(0) === 't') {
        cell.datatype = "t";
        cell.valuetype = value.type;
        cell.datavalue = value.value;
    }
    // Special number types (%, $, dates, etc.)
    else {
        cell.datatype = "c";
        cell.valuetype = value.type;
        cell.datavalue = value.value;
        cell.formula = rawvalue;
    }
};
