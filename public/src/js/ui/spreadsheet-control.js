/**
 * @fileoverview SocialCalc SpreadsheetControl
 * The code module of the SocialCalc package that lets you embed a spreadsheet
 * control with toolbar, etc., into a web page.
 * 
 * @author Dan Bricklin
 * @copyright 2008, 2009, 2010 Socialtext, Inc. All Rights Reserved.
 * @license CPAL-1.0
 */

/**
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
 * The Original Code is SocialCalc JavaScript SpreadsheetControl.
 * 
 * The Original Developer is the Initial Developer.
 * 
 * The Initial Developer of the Original Code is Socialtext, Inc. All portions of the code written by 
 * Socialtext, Inc., are Copyright (c) Socialtext, Inc. All Rights Reserved.
 * 
 * Contributor: Dan Bricklin.
 * 
 * EXHIBIT B. Attribution Information
 * 
 * When the SpreadsheetControl is producing and/or controlling the display the Graphic Image must be
 * displayed on the screen visible to the user in a manner comparable to that in the 
 * Original Code. The Attribution Phrase must be displayed as a "tooltip" or "hover-text" for
 * that image. The image must be linked to the Attribution URL so as to access that page
 * when clicked. If the user interface includes a prominent "about" display which includes
 * factual prominent attribution in a form similar to that in the "about" display included
 * with the Original Code, including Socialtext copyright notices and URLs, then the image
 * need not be linked to the Attribution URL but the "tool-tip" is still required.
 * 
 * Attribution Copyright Notice:
 * Copyright (C) 2010 Socialtext, Inc.
 * All Rights Reserved.
 * 
 * Attribution Phrase (not exceeding 10 words): SocialCalc
 * Attribution URL: http://www.socialcalc.org/
 * Graphic Image: The contents of the sc-logo.gif file in the Original Code or
 * a suitable replacement from http://www.socialcalc.org/licenses specified as
 * being for SocialCalc.
 * 
 * Display of Attribution Information is required in Larger Works which are defined 
 * in the CPAL as a work which combines Covered Code or portions thereof with code 
 * not governed by the terms of the CPAL.
 */

/**
 * Some of the other files in the SocialCalc package are licensed under
 * different licenses. Please note the licenses of the modules you use.
 * 
 * Code History:
 * 
 * Initially coded by Dan Bricklin of Software Garden, Inc., for Socialtext, Inc.
 * Unless otherwise specified, referring to "SocialCalc" in comments refers to this
 * JavaScript version of the code, not the SocialCalc Perl code.
 * 
 * See the comments in the main SocialCalc code module file of the SocialCalc package.
 */

var SocialCalc;
if (!SocialCalc) {
    alert("Main SocialCalc code module needed");
    SocialCalc = {};
}
if (!SocialCalc.TableEditor) {
    alert("SocialCalc TableEditor code module needed");
}

/**
 * @namespace SocialCalc
 */

/**
 * SpreadsheetControl class
 * @class
 */

/**
 * Global constant: Currently active spreadsheet control object
 * Right now there can only be one active at a time
 * @type {SocialCalc.SpreadsheetControl|null}
 */
SocialCalc.CurrentSpreadsheetControlObject = null;

/**
 * SpreadsheetControl constructor
 * Creates a new spreadsheet control instance
 * @constructor
 */
SocialCalc.SpreadsheetControl = function() {
    const scc = SocialCalc.Constants;

    // Properties:
    this.parentNode = null;
    this.spreadsheetDiv = null;
    this.requestedHeight = 0;
    this.requestedWidth = 0;
    this.requestedSpaceBelow = 0;
    this.height = 0;
    this.width = 0;
    /** @type {number} calculated amount for views below toolbar, etc. */
    this.viewheight = 0;

    /**
     * Tab definitions: An array where each tab is an object of the form:
     * {
     *   name: "name",
     *   text: "text-on-tab",
     *   html: "html-to-create div",
     *     replacements:
     *       "%s.": "SocialCalc", "%id.": spreadsheet.idPrefix, "%tbt.": spreadsheet.toolbartext
     *       Other replacements from spreadsheet.tabreplacements:
     *          replacementname: {regex: regular-expression-to-match-with-g, replacement: string}
     *   view: "viewname", // view to show when selected; "sheet" or missing/null is spreadsheet
     *   oncreate: function(spreadsheet, tab-name), // called when first created to initialize
     *   onclick: function(spreadsheet, tab-name), missing/null is sheet default
     *   onclickFocus: text, // spreadsheet.idPrefix+text is given the focus if present instead of normal KeyboardFocus
     *      or if text isn't a string, that value (e.g., true) is used for SocialCalc.CmdGotFocus
     *   onunclick: function(spreadsheet, tab-name), missing/null is sheet default
     * }
     * @type {Array}
     */
    this.tabs = [];
    
    /** @type {Object} when adding tabs, add tab-name: array-index to this object */
    this.tabnums = {};
    
    /** @type {Object} see use above */
    this.tabreplacements = {};
    
    /** @type {number} currently selected tab index in this.tabs or -1 (maintained by SocialCalc.SetTab) */
    this.currentTab = -1;

    /**
     * View definitions: An object where each view is an object of the form:
     * {
     *   name: "name", // localized when first set using SocialCalc.LocalizeString
     *   element: node-in-the-dom, // filled in when initialized
     *   replacements: {}, // see below
     *   html: "html-to-create div",
     *     replacements:
     *       "%s.": "SocialCalc", "%id.": spreadsheet.idPrefix, "%tbt.": spreadsheet.toolbartext, "%img.": spreadsheet.imagePrefix,
     *       SocialCalc.LocalizeSubstring replacements ("%loc!string!" and "%ssc!constant-name!")
     *       Other replacements from viewobject.replacements:
     *          replacementname: {regex: regular-expression-to-match-with-g, replacement: string}
     *   divStyle: attributes for sheet div (SocialCalc.setStyles format)
     *   oncreate: function(spreadsheet, viewobject), // called when first created to initialize
     *   needsresize: true/false/null, // if true, do resize calc after displaying
     *   onresize: function(spreadsheet, viewobject), // called if needs resize
     *   values: {} // optional values to share with onclick handlers, etc.
     * }
     * 
     * There is always a "sheet" view.
     * @type {Object} {viewname: view-object, ...}
     */
    this.views = {};

    // Dynamic properties:
    this.sheet = null;
    this.context = null;
    this.editor = null;
    this.spreadsheetDiv = null;
    this.editorDiv = null;

    /** @type {string} remembered range for sort tab */
    this.sortrange = "";

    /** @type {string} remembered range from movefrom used by movepaste/moveinsert */
    this.moverange = "";

    // Constants:
    /** @type {string} prefix added to element ids used here, should end in "-" */
    this.idPrefix = "SocialCalc-";
    
    /** @type {string} boundary used by SpreadsheetControlCreateSpreadsheetSave */
    this.multipartBoundary = "SocialCalcSpreadsheetControlSave";
    
    /** @type {string} prefix added to img src */
    this.imagePrefix = scc.defaultImagePrefix;

    this.toolbarbackground = scc.SCToolbarbackground;
    this.tabbackground = scc.SCTabbackground;
    this.tabselectedCSS = scc.SCTabselectedCSS;
    this.tabplainCSS = scc.SCTabplainCSS;
    this.toolbartext = scc.SCToolbartext;

    /** @type {number} in pixels, will contain a text input box */
    this.formulabarheight = scc.SCFormulabarheight;

    if (scc.doWorkBook) {
        this.sheetbarheight = scc.SCSheetBarHeight;
        this.sheetbarCSS = scc.SCSheetBarCSS;
    } else {
        this.sheetbarheight = 0;
    }

    /** @type {number} in pixels */
    this.statuslineheight = scc.SCStatuslineheight;
    this.statuslineCSS = scc.SCStatuslineCSS;

    // Callbacks:
    /** 
     * A function called for Clipboard Export button: this.ExportCallback(spreadsheet_control_object)
     * @type {Function|null}
     */
    this.ExportCallback = null;

    // Initialization Code:
    this.sheet = new SocialCalc.Sheet();
    this.context = new SocialCalc.RenderContext(this.sheet);
    this.context.showGrid = true;
    this.context.showRCHeaders = true;
    this.editor = new SocialCalc.TableEditor(this.context);
    this.editor.StatusCallback.statusline = {
        func: SocialCalc.SpreadsheetControlStatuslineCallback,
        params: {
            statuslineid: `${this.idPrefix}statusline`,
            recalcid1: `${this.idPrefix}divider_recalc`,
            recalcid2: `${this.idPrefix}button_recalc`
        }
    };

    SocialCalc.CurrentSpreadsheetControlObject = this; // remember this for rendezvousing on events

    this.editor.MoveECellCallback.movefrom = (editor) => {
        const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
        spreadsheet.context.cursorsuffix = "";
        if (editor.range2.hasrange && !editor.cellhandles.noCursorSuffix) {
            if (editor.ecell.row === editor.range2.top && 
                (editor.ecell.col < editor.range2.left || editor.ecell.col > editor.range2.right + 1)) {
                spreadsheet.context.cursorsuffix = "insertleft";
            }
            if (editor.ecell.col === editor.range2.left && 
                (editor.ecell.row < editor.range2.top || editor.ecell.row > editor.range2.bottom + 1)) {
                spreadsheet.context.cursorsuffix = "insertup";
            }
        }
    };

    /**
     * Formula bar buttons configuration
     * @type {Object}
     */
    this.formulabuttons = {
        formulafunctions: {
            image: "formuladialog.gif", 
            tooltip: "Functions", // tooltips are localized when set below
            command: SocialCalc.SpreadsheetControl.DoFunctionList
        },
        multilineinput: {
            image: "multilinedialog.gif", 
            tooltip: "Multi-line Input Box",
            command: SocialCalc.SpreadsheetControl.DoMultiline
        },
        link: {
            image: "linkdialog.gif", 
            tooltip: "Link Input Box",
            command: SocialCalc.SpreadsheetControl.DoLink
        },
        sum: {
            image: "sumdialog.gif", 
            tooltip: "Auto Sum",
            command: SocialCalc.SpreadsheetControl.DoSum
        }
        // image: {image: "sumdialog.gif", tooltip: "Insert",
        //          command: SocialCalc.Images.Insert }
    };

    // Default tabs:

    /**
     * Edit tab configuration
     */
    this.tabnums.edit = this.tabs.length;
    this.tabs.push({
        name: "edit", 
        text: "Edit", 
        html: ' <div id="%id.edittools" style="padding:10px 0px 0px 0px;">' +
            '&nbsp;<img id="%id.button_undo" src="%img.undo.gif" style="vertical-align:bottom;">' +
            ' <img id="%id.button_redo" src="%img.redo.gif" style="vertical-align:bottom;">' +
            ' &nbsp;<img src="%img.divider1.gif" style="vertical-align:bottom;">&nbsp; ' +
            '<img id="%id.button_copy" src="%img.copy.gif" style="vertical-align:bottom;">' +
            ' <img id="%id.button_cut" src="%img.cut.gif" style="vertical-align:bottom;">' +
            ' <img id="%id.button_paste" src="%img.paste.gif" style="vertical-align:bottom;">' +
            ' &nbsp;<img src="%img.divider1.gif" style="vertical-align:bottom;">&nbsp; ' +
            '<img id="%id.button_delete" src="%img.delete.gif" style="vertical-align:bottom;">' +
            ' <img id="%id.button_pasteformats" src="%img.pasteformats.gif" style="vertical-align:bottom;">' +
            ' &nbsp;<img src="%img.divider1.gif" style="vertical-align:bottom;">&nbsp; ' +
            '<img id="%id.button_filldown" src="%img.filldown.gif" style="vertical-align:bottom;">' +
            ' <img id="%id.button_fillright" src="%img.fillright.gif" style="vertical-align:bottom;">' +
            ' &nbsp;<img src="%img.divider1.gif" style="vertical-align:bottom;">&nbsp; ' +
            '<img id="%id.button_movefrom" src="%img.movefromoff.gif" style="vertical-align:bottom;">' +
            ' <img id="%id.button_movepaste" src="%img.movepasteoff.gif" style="vertical-align:bottom;">' +
            ' <img id="%id.button_moveinsert" src="%img.moveinsertoff.gif" style="vertical-align:bottom;">' +
            ' &nbsp;<img src="%img.divider1.gif" style="vertical-align:bottom;">&nbsp; ' +
            '<img id="%id.button_alignleft" src="%img.alignleft.gif" style="vertical-align:bottom;">' +
            ' <img id="%id.button_aligncenter" src="%img.aligncenter.gif" style="vertical-align:bottom;">' +
            ' <img id="%id.button_alignright" src="%img.alignright.gif" style="vertical-align:bottom;">' +
            ' &nbsp;<img src="%img.divider1.gif" style="vertical-align:bottom;">&nbsp; ' +
            '<img id="%id.button_borderon" src="%img.borderson.gif" style="vertical-align:bottom;"> ' +
            ' <img id="%id.button_borderoff" src="%img.bordersoff.gif" style="vertical-align:bottom;"> ' +
            ' <img id="%id.button_swapcolors" src="%img.swapcolors.gif" style="vertical-align:bottom;"> ' +
            ' &nbsp;<img src="%img.divider1.gif" style="vertical-align:bottom;">&nbsp; ' +
            '<img id="%id.button_merge" src="%img.merge.gif" style="vertical-align:bottom;"> ' +
            ' <img id="%id.button_unmerge" src="%img.unmerge.gif" style="vertical-align:bottom;"> ' +
            ' &nbsp;<img src="%img.divider1.gif" style="vertical-align:bottom;">&nbsp; ' +
            '<img id="%id.button_insertrow" src="%img.insertrow.gif" style="vertical-align:bottom;"> ' +
            ' <img id="%id.button_insertcol" src="%img.insertcol.gif" style="vertical-align:bottom;"> ' +
            '&nbsp; <img id="%id.button_deleterow" src="%img.deleterow.gif" style="vertical-align:bottom;"> ' +
            ' <img id="%id.button_deletecol" src="%img.deletecol.gif" style="vertical-align:bottom;"> ' +
            ' &nbsp;<img id="%id.divider_recalc" src="%img.divider1.gif" style="vertical-align:bottom;">&nbsp; ' +
            '<img id="%id.button_recalc" src="%img.recalc.gif" style="vertical-align:bottom;"> ' +
            ' </div>',
        oncreate: null, // function(spreadsheet, viewobject) {SocialCalc.DoCmd(null, "fill-rowcolstuff");},
        onclick: null
    });

/**
 * Settings (Format) tab configuration
 */
this.tabnums.settings = this.tabs.length;
this.tabs.push({
    name: "settings", 
    text: "Format", 
    html: '<div id="%id.settingstools" style="display:none;">' +
        ' <div id="%id.sheetsettingstoolbar" style="display:none;">' +
        '  <table cellspacing="0" cellpadding="0"><tr><td>' +
        '   <div style="%tbt.">%loc!SHEET SETTINGS!:</div>' +
        '   </td></tr><tr><td>' +
        '   <input id="%id.settings-savesheet" type="button" value="%loc!Save!" onclick="SocialCalc.SettingsControlSave(\'sheet\');">' +
        '   <input type="button" value="%loc!Cancel!" onclick="SocialCalc.SettingsControlSave(\'cancel\');">' +
        '   <input type="button" value="%loc!Show Cell Settings!" onclick="SocialCalc.SpreadsheetControlSettingsSwitch(\'cell\');return false;">' +
        '   </td></tr></table>' +
        ' </div>' +
        ' <div id="%id.cellsettingstoolbar" style="display:none;">' +
        '  <table cellspacing="0" cellpadding="0"><tr><td>' +
        '   <div style="%tbt.">%loc!CELL SETTINGS!: <span id="%id.settingsecell">&nbsp;</span></div>' +
        '   </td></tr><tr><td>' +
        '  <input id="%id.settings-savecell" type="button" value="%loc!Save!" onclick="SocialCalc.SettingsControlSave(\'cell\');">' +
        '  <input type="button" value="%loc!Cancel!" onclick="SocialCalc.SettingsControlSave(\'cancel\');">' +
        '  <input type="button" value="%loc!Show Sheet Settings!" onclick="SocialCalc.SpreadsheetControlSettingsSwitch(\'sheet\');return false;">' +
        '  </td></tr></table>' +
        ' </div>' +
        '</div>',
    view: "settings",
    /**
     * Handle settings tab click event
     * @param {Object} s - Spreadsheet object
     * @param {string} t - Tab name
     */
    onclick: (s, t) => {
        SocialCalc.SettingsControls.idPrefix = s.idPrefix; // used to get color chooser div
        SocialCalc.SettingControlReset();
        const sheetattribs = s.sheet.EncodeSheetAttributes();
        const cellattribs = s.sheet.EncodeCellAttributes(s.editor.ecell.coord);
        SocialCalc.SettingsControlLoadPanel(s.views.settings.values.sheetspanel, sheetattribs);
        SocialCalc.SettingsControlLoadPanel(s.views.settings.values.cellspanel, cellattribs);
        document.getElementById(`${s.idPrefix}settingsecell`).innerHTML = s.editor.ecell.coord;
        SocialCalc.SpreadsheetControlSettingsSwitch("cell");
        s.views.settings.element.style.height = `${s.viewheight}px`;
        s.views.settings.element.firstChild.style.height = `${s.viewheight}px`;

        // Set save message
        let range;
        if (s.editor.range.hasrange) {
            range = `${SocialCalc.crToCoord(s.editor.range.left, s.editor.range.top)}:${SocialCalc.crToCoord(s.editor.range.right, s.editor.range.bottom)}`;
        } else {
            range = s.editor.ecell.coord;
        }
        document.getElementById(`${s.idPrefix}settings-savecell`).value = `${SocialCalc.LocalizeString("Save to")}: ${range}`;
    },
    onclickFocus: true
});

/**
 * Settings view configuration
 * @type {Object}
 */
this.views.settings = {
    name: "settings", 
    values: {},
    /**
     * Create settings view
     * @param {Object} s - Spreadsheet object
     * @param {Object} viewobj - View object
     */
    oncreate: (s, viewobj) => {
        const scc = SocialCalc.Constants;

        /**
         * Sheet settings panel configuration
         * @type {Object}
         */
        viewobj.values.sheetspanel = {
            // name: "sheet",
            colorchooser: { id: `${s.idPrefix}scolorchooser` },
            formatnumber: { 
                setting: "numberformat", 
                type: "PopupList", 
                id: `${s.idPrefix}formatnumber`,
                initialdata: scc.SCFormatNumberFormats 
            },
            formattext: { 
                setting: "textformat", 
                type: "PopupList", 
                id: `${s.idPrefix}formattext`,
                initialdata: scc.SCFormatTextFormats 
            },
            fontfamily: { 
                setting: "fontfamily", 
                type: "PopupList", 
                id: `${s.idPrefix}fontfamily`,
                initialdata: scc.SCFormatFontfamilies 
            },
            fontlook: { 
                setting: "fontlook", 
                type: "PopupList", 
                id: `${s.idPrefix}fontlook`,
                initialdata: scc.SCFormatFontlook 
            },
            fontsize: { 
                setting: "fontsize", 
                type: "PopupList", 
                id: `${s.idPrefix}fontsize`,
                initialdata: scc.SCFormatFontsizes 
            },
            textalignhoriz: { 
                setting: "textalignhoriz", 
                type: "PopupList", 
                id: `${s.idPrefix}textalignhoriz`,
                initialdata: scc.SCFormatTextAlignhoriz 
            },
            numberalignhoriz: { 
                setting: "numberalignhoriz", 
                type: "PopupList", 
                id: `${s.idPrefix}numberalignhoriz`,
                initialdata: scc.SCFormatNumberAlignhoriz 
            },
            alignvert: { 
                setting: "alignvert", 
                type: "PopupList", 
                id: `${s.idPrefix}alignvert`,
                initialdata: scc.SCFormatAlignVertical 
            },
            textcolor: { 
                setting: "textcolor", 
                type: "ColorChooser", 
                id: `${s.idPrefix}textcolor` 
            },
            bgcolor: { 
                setting: "bgcolor", 
                type: "ColorChooser", 
                id: `${s.idPrefix}bgcolor` 
            },
            padtop: { 
                setting: "padtop", 
                type: "PopupList", 
                id: `${s.idPrefix}padtop`,
                initialdata: scc.SCFormatPadsizes 
            },
            padright: { 
                setting: "padright", 
                type: "PopupList", 
                id: `${s.idPrefix}padright`,
                initialdata: scc.SCFormatPadsizes 
            },
            padbottom: { 
                setting: "padbottom", 
                type: "PopupList", 
                id: `${s.idPrefix}padbottom`,
                initialdata: scc.SCFormatPadsizes 
            },
            padleft: { 
                setting: "padleft", 
                type: "PopupList", 
                id: `${s.idPrefix}padleft`,
                initialdata: scc.SCFormatPadsizes 
            },
            colwidth: { 
                setting: "colwidth", 
                type: "PopupList", 
                id: `${s.idPrefix}colwidth`,
                initialdata: scc.SCFormatColwidth 
            },
            recalc: { 
                setting: "recalc", 
                type: "PopupList", 
                id: `${s.idPrefix}recalc`,
                initialdata: scc.SCFormatRecalc 
            }
        };

        /**
         * Cell settings panel configuration
         * @type {Object}
         */
        viewobj.values.cellspanel = {
            name: "cell",
            colorchooser: { id: `${s.idPrefix}scolorchooser` },
            cformatnumber: { 
                setting: "numberformat", 
                type: "PopupList", 
                id: `${s.idPrefix}cformatnumber`,
                initialdata: scc.SCFormatNumberFormats 
            },
            cformattext: { 
                setting: "textformat", 
                type: "PopupList", 
                id: `${s.idPrefix}cformattext`,
                initialdata: scc.SCFormatTextFormats 
            },
            cfontfamily: { 
                setting: "fontfamily", 
                type: "PopupList", 
                id: `${s.idPrefix}cfontfamily`,
                initialdata: scc.SCFormatFontfamilies 
            },
            cfontlook: { 
                setting: "fontlook", 
                type: "PopupList", 
                id: `${s.idPrefix}cfontlook`,
                initialdata: scc.SCFormatFontlook 
            },
            cfontsize: { 
                setting: "fontsize", 
                type: "PopupList", 
                id: `${s.idPrefix}cfontsize`,
                initialdata: scc.SCFormatFontsizes 
            },
            calignhoriz: { 
                setting: "alignhoriz", 
                type: "PopupList", 
                id: `${s.idPrefix}calignhoriz`,
                initialdata: scc.SCFormatTextAlignhoriz 
            },
            calignvert: { 
                setting: "alignvert", 
                type: "PopupList", 
                id: `${s.idPrefix}calignvert`,
                initialdata: scc.SCFormatAlignVertical 
            },
            ctextcolor: { 
                setting: "textcolor", 
                type: "ColorChooser", 
                id: `${s.idPrefix}ctextcolor` 
            },
            cbgcolor: { 
                setting: "bgcolor", 
                type: "ColorChooser", 
                id: `${s.idPrefix}cbgcolor` 
            },
            cbt: { 
                setting: "bt", 
                type: "BorderSide", 
                id: `${s.idPrefix}cbt` 
            },
            cbr: { 
                setting: "br", 
                type: "BorderSide", 
                id: `${s.idPrefix}cbr` 
            },
            cbb: { 
                setting: "bb", 
                type: "BorderSide", 
                id: `${s.idPrefix}cbb` 
            },
            cbl: { 
                setting: "bl", 
                type: "BorderSide", 
                id: `${s.idPrefix}cbl` 
            },
            cpadtop: { 
                setting: "padtop", 
                type: "PopupList", 
                id: `${s.idPrefix}cpadtop`,
                initialdata: scc.SCFormatPadsizes 
            },
            cpadright: { 
                setting: "padright", 
                type: "PopupList", 
                id: `${s.idPrefix}cpadright`,
                initialdata: scc.SCFormatPadsizes 
            },
            cpadbottom: { 
                setting: "padbottom", 
                type: "PopupList", 
                id: `${s.idPrefix}cpadbottom`,
                initialdata: scc.SCFormatPadsizes 
            },
            cpadleft: { 
                setting: "padleft", 
                type: "PopupList", 
                id: `${s.idPrefix}cpadleft`,
                initialdata: scc.SCFormatPadsizes 
            }
        };
                SocialCalc.SettingsControlInitializePanel(viewobj.values.sheetspanel);
        SocialCalc.SettingsControlInitializePanel(viewobj.values.cellspanel);
    },
    /**
     * Text replacement patterns for the settings view HTML
     * @type {Object}
     */
    replacements: {
        itemtitle: {
            regex: /\%itemtitle\./g, 
            replacement: 'style="padding:12px 10px 0px 10px;font-weight:bold;text-align:right;vertical-align:top;font-size:small;"'
        },
        sectiontitle: {
            regex: /\%sectiontitle\./g, 
            replacement: 'style="padding:16px 10px 0px 0px;font-weight:bold;vertical-align:top;font-size:small;color:#C00;"'
        },
        parttitle: {
            regex: /\%parttitle\./g, 
            replacement: 'style="font-weight:bold;font-size:x-small;padding:0px 0px 3px 0px;"'
        },
        itembody: {
            regex: /\%itembody\./g, 
            replacement: 'style="padding:12px 0px 0px 0px;vertical-align:top;font-size:small;"'
        },
        bodypart: {
            regex: /\%bodypart\./g, 
            replacement: 'style="padding:0px 10px 0px 0px;font-size:small;vertical-align:top;"'
        }
    },
    /** @type {string} CSS styles for the settings view container */
    divStyle: "border:1px solid black;overflow:auto;",
    /** @type {string} HTML template for the settings view interface */
    html: '<div id="%id.scolorchooser" style="display:none;position:absolute;z-index:20;"></div>' +
        '<table cellspacing="0" cellpadding="0">' +
        ' <tr><td style="vertical-align:top;">' +
        '<table id="%id.sheetsettingstable" style="display:none;" cellspacing="0" cellpadding="0">' +
        '<tr>' +
        ' <td %itemtitle.><br>%loc!Default Format!:</td>' +
        ' <td %itembody.>' +
        '   <table cellspacing="0" cellpadding="0"><tr>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Number!</div>' +
        '     <span id="%id.formatnumber"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Text!</div>' +
        '     <span id="%id.formattext"></span>' +
        '    </td>' +
        '   </tr></table>' +
        ' </td>' +
        '</tr>' +
        '<tr>' +
        ' <td %itemtitle.><br>%loc!Default Alignment!:</td>' +
        ' <td %itembody.>' +
        '   <table cellspacing="0" cellpadding="0"><tr>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Text Horizontal!</div>' +
        '     <span id="%id.textalignhoriz"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Number Horizontal!</div>' +
        '     <span id="%id.numberalignhoriz"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Vertical!</div>' +
        '     <span id="%id.alignvert"></span>' +
        '    </td>' +
        '   </tr></table>' +
        ' </td>' +
        '</tr>' +
        '<tr>' +
        ' <td %itemtitle.><br>%loc!Default Font!:</td>' +
        ' <td %itembody.>' +
        '   <table cellspacing="0" cellpadding="0"><tr>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Family!</div>' +
        '     <span id="%id.fontfamily"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Bold &amp; Italics!</div>' +
        '     <span id="%id.fontlook"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Size!</div>' +
        '     <span id="%id.fontsize"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Color!</div>' +
        '     <div id="%id.textcolor"></div>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Background!</div>' +
        '     <div id="%id.bgcolor"></div>' +
        '    </td>' +
        '   </tr></table>' +
        ' </td>' +
        '</tr>' +
        '<tr>' +
        ' <td %itemtitle.><br>%loc!Default Padding!:</td>' +
        ' <td %itembody.>' +
        '   <table cellspacing="0" cellpadding="0"><tr>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Top!</div>' +
        '     <span id="%id.padtop"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Right!</div>' +
        '     <span id="%id.padright"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Bottom!</div>' +
        '     <span id="%id.padbottom"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Left!</div>' +
        '     <span id="%id.padleft"></span>' +
        '    </td>' +
        '   </tr></table>' +
        ' </td>' +
        '</tr>' +
        '<tr>' +
        ' <td %itemtitle.><br>%loc!Default Column Width!:</td>' +
        ' <td %itembody.>' +
        '   <table cellspacing="0" cellpadding="0"><tr>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>&nbsp;</div>' +
        '     <span id="%id.colwidth"></span>' +
        '    </td>' +
        '   </tr></table>' +
        ' </td>' +
        '</tr>' +
        '<tr>' +
        ' <td %itemtitle.><br>%loc!Recalculation!:</td>' +
        ' <td %itembody.>' +
        '   <table cellspacing="0" cellpadding="0"><tr>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>&nbsp;</div>' +
        '     <span id="%id.recalc"></span>' +
        '    </td>' +
        '   </tr></table>' +
        ' </td>' +
        '</tr>' +
        '</table>' +
        '<table id="%id.cellsettingstable" cellspacing="0" cellpadding="0">' +
        '<tr>' +
        ' <td %itemtitle.><br>%loc!Format!:</td>' +
        ' <td %itembody.>' +
        '   <table cellspacing="0" cellpadding="0"><tr>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Number!</div>' +
        '     <span id="%id.cformatnumber"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Text!</div>' +
        '     <span id="%id.cformattext"></span>' +
        '    </td>' +
        '   </tr></table>' +
        ' </td>' +
        '</tr>' +
        '<tr>' +
        ' <td %itemtitle.><br>%loc!Alignment!:</td>' +
        ' <td %itembody.>' +
        '   <table cellspacing="0" cellpadding="0"><tr>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Horizontal!</div>' +
        '     <span id="%id.calignhoriz"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Vertical!</div>' +
        '     <span id="%id.calignvert"></span>' +
        '    </td>' +
        '   </tr></table>' +
        ' </td>' +
        '</tr>' +
        '<tr>' +
        ' <td %itemtitle.><br>%loc!Font!:</td>' +
        ' <td %itembody.>' +
        '   <table cellspacing="0" cellpadding="0"><tr>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Family!</div>' +
        '     <span id="%id.cfontfamily"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Bold &amp; Italics!</div>' +
        '     <span id="%id.cfontlook"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Size!</div>' +
        '     <span id="%id.cfontsize"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Color!</div>' +
        '     <div id="%id.ctextcolor"></div>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Background!</div>' +
        '     <div id="%id.cbgcolor"></div>' +
        '    </td>' +
        '   </tr></table>' +
        ' </td>' +
        '</tr>' +
        '<tr>' +
        ' <td %itemtitle.><br>%loc!Borders!:</td>' +
        ' <td %itembody.>' +
        '   <table cellspacing="0" cellpadding="0">' +
        '    <tr><td %bodypart. colspan="3"><div %parttitle.>%loc!Top Border!</div></td>' +
        '     <td %bodypart. colspan="3"><div %parttitle.>%loc!Right Border!</div></td>' +
        '     <td %bodypart. colspan="3"><div %parttitle.>%loc!Bottom Border!</div></td>' +
        '     <td %bodypart. colspan="3"><div %parttitle.>%loc!Left Border!</div></td>' +
        '    </tr><tr>' +
        '    <td %bodypart.>' +
        '     <input id="%id.cbt-onoff-bcb" onclick="SocialCalc.SettingsControlOnchangeBorder(this);" type="checkbox">' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div id="%id.cbt-color"></div>' +
        '    </td>' +
        '    <td>&nbsp;&nbsp;&nbsp;&nbsp;</td>' +
        '    <td %bodypart.>' +
        '     <input id="%id.cbr-onoff-bcb" onclick="SocialCalc.SettingsControlOnchangeBorder(this);" type="checkbox">' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div id="%id.cbr-color"></div>' +
        '    </td>' +
        '    <td>&nbsp;&nbsp;&nbsp;&nbsp;</td>' +
        '    <td %bodypart.>' +
        '     <input id="%id.cbb-onoff-bcb" onclick="SocialCalc.SettingsControlOnchangeBorder(this);" type="checkbox">' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div id="%id.cbb-color"></div>' +
        '    </td>' +
        '    <td>&nbsp;&nbsp;&nbsp;&nbsp;</td>' +
        '    <td %bodypart.>' +
        '     <input id="%id.cbl-onoff-bcb" onclick="SocialCalc.SettingsControlOnchangeBorder(this);" type="checkbox">' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div id="%id.cbl-color"></div>' +
        '    </td>' +
        '    <td>&nbsp;&nbsp;&nbsp;&nbsp;</td>' +
        '   </tr></table>' +
        ' </td>' +
        '</tr>' +
        '<tr>' +
        ' <td %itemtitle.><br>%loc!Padding!:</td>' +
        ' <td %itembody.>' +
        '   <table cellspacing="0" cellpadding="0"><tr>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Top!</div>' +
        '     <span id="%id.cpadtop"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Right!</div>' +
        '     <span id="%id.cpadright"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Bottom!</div>' +
        '     <span id="%id.cpadbottom"></span>' +
        '    </td>' +
        '    <td %bodypart.>' +
        '     <div %parttitle.>%loc!Left!</div>' +
        '     <span id="%id.cpadleft"></span>' +
        '    </td>' +
        '   </tr></table>' +
        ' </td>' +
        '</tr>' +
        '</table>' +
        ' </td><td style="vertical-align:top;padding:12px 0px 0px 12px;">' +
        '  <div style="width:100px;height:100px;overflow:hidden;border:1px solid black;background-color:#EEE;padding:6px;">' +
        '   <table cellspacing="0" cellpadding="0"><tr>' +
        '    <td id="sample-text" style="height:100px;width:100px;"><div>%loc!This is a<br>sample!</div><div>-1234.5</div></td>' +
        '   </tr></table>' +
        '  </div>' +
        ' </td></tr></table>' +
        '<br>'
};

/**
 * Sort tab configuration
 */
this.tabnums.sort = this.tabs.length;
this.tabs.push({
    name: "sort", 
    text: "Sort", 
    html: ' <div id="%id.sorttools" style="display:none;">' +
        '  <table cellspacing="0" cellpadding="0"><tr>' +
        '   <td style="vertical-align:top;padding-right:4px;width:160px;">' +
        '    <div style="%tbt.">%loc!Set Cells To Sort!</div>' +
        '    <select id="%id.sortlist" size="1" onfocus="%s.CmdGotFocus(this);"><option selected>[select range]</option></select>' +
        '    <input type="button" value="%loc!OK!" onclick="%s.DoCmd(this, \'ok-setsort\');" style="font-size:x-small;">' +
        '   </td>' +
        '   <td style="vertical-align:middle;padding-right:16px;width:100px;text-align:right;">' +
        '    <div style="%tbt.">&nbsp;</div>' +
        '    <input type="button" id="%id.sortbutton" value="%loc!Sort Cells! A1:A1" onclick="%s.DoCmd(this, \'dosort\');" style="visibility:hidden;">' +
        '   </td>' +
        '   <td style="vertical-align:top;padding-right:16px;">' +
        '    <table cellspacing="0" cellpadding="0"><tr>' +
        '     <td style="vertical-align:top;">' +
        '      <div style="%tbt.">%loc!Major Sort!</div>' +
        '      <select id="%id.majorsort" size="1" onfocus="%s.CmdGotFocus(this);"></select>' +
        '     </td><td>' +
        '      <input type="radio" name="majorsort" id="%id.majorsortup" value="up" checked><span style="font-size:x-small;color:#FFF;">%loc!Up!</span><br>' +
        '      <input type="radio" name="majorsort" id="%id.majorsortdown" value="down"><span style="font-size:x-small;color:#FFF;">%loc!Down!</span>' +
        '     </td>' +
        '    </tr></table>' +
        '   </td>' +
        '   <td style="vertical-align:top;padding-right:16px;">' +
        '    <table cellspacing="0" cellpadding="0"><tr>' +
        '     <td style="vertical-align:top;">' +
        '      <div style="%tbt.">%loc!Minor Sort!</div>' +
        '      <select id="%id.minorsort" size="1" onfocus="%s.CmdGotFocus(this);"></select>' +
        '     </td><td>' +
        '      <input type="radio" name="minorsort" id="%id.minorsortup" value="up" checked><span style="font-size:x-small;color:#FFF;">%loc!Up!</span><br>' +
        '      <input type="radio" name="minorsort" id="%id.minorsortdown" value="down"><span style="font-size:x-small;color:#FFF;">%loc!Down!</span>' +
        '     </td>' +
        '    </tr></table>' +
        '   </td>' +
        '   <td style="vertical-align:top;padding-right:16px;">' +
        '    <table cellspacing="0" cellpadding="0"><tr>' +
        '     <td style="vertical-align:top;">' +
        '      <div style="%tbt.">%loc!Last Sort!</div>' +
        '      <select id="%id.lastsort" size="1" onfocus="%s.CmdGotFocus(this);"></select>' +
        '     </td><td>' +
        '      <input type="radio" name="lastsort" id="%id.lastsortup" value="up" checked><span style="font-size:x-small;color:#FFF;">%loc!Up!</span><br>' +
        '      <input type="radio" name="lastsort" id="%id.lastsortdown" value="down"><span style="font-size:x-small;color:#FFF;">%loc!Down!</span>' +
        '     </td>' +
        '    </tr></table>' +
        '   </td>' +
        '  </tr></table>' +
        ' </div>',
    onclick: SocialCalc.SpreadsheetControlSortOnclick
});
this.editor.SettingsCallbacks.sort = {
    save: SocialCalc.SpreadsheetControlSortSave, 
    load: SocialCalc.SpreadsheetControlSortLoad
};

/**
 * Audit tab configuration
 */
this.tabnums.audit = this.tabs.length;
this.tabs.push({
    name: "audit", 
    text: "Audit", 
    html: '<div id="%id.audittools" style="display:none;">' +
        ' <div style="%tbt.">&nbsp;</div>' +
        '</div>',
    view: "audit",
    /**
     * Handle audit tab click event
     * @param {Object} s - Spreadsheet object
     * @param {string} t - Tab name
     */
    onclick: (s, t) => {
        const SCLoc = SocialCalc.LocalizeString;
        let str = `<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;"><tr><td style="font-size:small;padding:6px;"><b>${SCLoc("Audit Trail This Session")}:</b><br><br>`;
        const stack = s.sheet.changes.stack;
        const tos = s.sheet.changes.tos;
        
        for (let i = 0; i < stack.length; i++) {
            if (i === tos + 1) {
                str += `<br></td></tr><tr><td style="font-size:small;background-color:#EEE;padding:6px;">${SCLoc("UNDONE STEPS")}:<br>`;
            }
            for (let j = 0; j < stack[i].command.length; j++) {
                str += `${SocialCalc.special_chars(stack[i].command[j])}<br>`;
            }
        }
        s.views.audit.element.innerHTML = `${str}</td></tr></table>`;
        SocialCalc.CmdGotFocus(true);
    },
    onclickFocus: true
});

/**
 * Audit view configuration
 * @type {Object}
 */
this.views.audit = {
    name: "audit",
    divStyle: "border:1px solid black;overflow:auto;",
    html: 'Audit Trail'
};

/**
 * Comment tab configuration
 */
this.tabnums.comment = this.tabs.length;
this.tabs.push({
    name: "comment", 
    text: "Comment", 
    html: '<div id="%id.commenttools" style="display:none;">' +
        '<table cellspacing="0" cellpadding="0"><tr><td>' +
        '<textarea id="%id.commenttext" style="font-size:small;height:32px;width:600px;overflow:auto;" onfocus="%s.CmdGotFocus(this);"></textarea>' +
        '</td><td style="vertical-align:top;">' +
        '&nbsp;<input type="button" value="%loc!Save!" onclick="%s.SpreadsheetControlCommentSet();" style="font-size:x-small;">' +
        '</td></tr></table>' +
        '</div>',
    view: "sheet",
    onclick: SocialCalc.SpreadsheetControlCommentOnclick,
    onunclick: SocialCalc.SpreadsheetControlCommentOnunclick
});

/**
 * Names tab configuration
 */
this.tabnums.names = this.tabs.length;
this.tabs.push({
    name: "names", 
    text: "Names", 
    html: '<div id="%id.namestools" style="display:none;">' +
        '  <table cellspacing="0" cellpadding="0"><tr>' +
        '   <td style="vertical-align:top;padding-right:24px;">' +
        '    <div style="%tbt.">%loc!Existing Names!</div>' +
        '    <select id="%id.nameslist" size="1" onchange="%s.SpreadsheetControlNamesChangedName();" onfocus="%s.CmdGotFocus(this);"><option selected>[New]</option></select>' +
        '   </td>' +
        '   <td style="vertical-align:top;padding-right:6px;">' +
        '    <div style="%tbt.">%loc!Name!</div>' +
        '    <input type="text" id="%id.namesname" style="font-size:x-small;width:75px;" onfocus="%s.CmdGotFocus(this);">' +
        '   </td>' +
        '   <td style="vertical-align:top;padding-right:6px;">' +
        '    <div style="%tbt.">%loc!Description!</div>' +
        '    <input type="text" id="%id.namesdesc" style="font-size:x-small;width:150px;" onfocus="%s.CmdGotFocus(this);">' +
        '   </td>' +
        '   <td style="vertical-align:top;padding-right:6px;">' +
        '    <div style="%tbt.">%loc!Value!</div>' +
        '    <input type="text" id="%id.namesvalue" width="16" style="font-size:x-small;width:100px;" onfocus="%s.CmdGotFocus(this);">' +
        '   </td>' +
        '   <td style="vertical-align:top;padding-right:12px;width:100px;">' +
        '    <div style="%tbt.">%loc!Set Value To!</div>' +
        '    <input type="button" id="%id.namesrangeproposal" value="A1" onclick="%s.SpreadsheetControlNamesSetValue();" style="font-size:x-small;">' +
        '   </td>' +
        '   <td style="vertical-align:top;padding-right:6px;">' +
        '    <div style="%tbt.">&nbsp;</div>' +
        '    <input type="button" value="%loc!Save!" onclick="%s.SpreadsheetControlNamesSave();" style="font-size:x-small;">' +
        '    <input type="button" value="%loc!Delete!" onclick="%s.SpreadsheetControlNamesDelete()" style="font-size:x-small;">' +
        '   </td>' +
        '  </tr></table>' +
        '</div>',
    view: "sheet",
    onclick: SocialCalc.SpreadsheetControlNamesOnclick,
    onunclick: SocialCalc.SpreadsheetControlNamesOnunclick
});

/**
 * Clipboard tab configuration
 */
this.tabnums.clipboard = this.tabs.length;
this.tabs.push({
    name: "clipboard", 
    text: "Clipboard", 
    html: '<div id="%id.clipboardtools" style="display:none;">' +
        '  <table cellspacing="0" cellpadding="0"><tr>' +
        '   <td style="vertical-align:top;padding-right:24px;">' +
        '    <div style="%tbt.">' +
        '     &nbsp;' +
        '    </div>' +
        '   </td>' +
        '  </tr></table>' +
        '</div>',
    view: "clipboard",
    onclick: SocialCalc.SpreadsheetControlClipboardOnclick,
    onclickFocus: "clipboardtext"
});

/**
 * Clipboard view configuration
 * @type {Object}
 */
this.views.clipboard = {
    name: "clipboard", 
    divStyle: "overflow:auto;", 
    html: ' <div style="font-size:x-small;padding:5px 0px 10px 0px;">' +
        '  <b>%loc!Display Clipboard in!:</b>' +
        '  <input type="radio" id="%id.clipboardformat-tab" name="%id.clipboardformat" checked onclick="%s.SpreadsheetControlClipboardFormat(\'tab\');"> %loc!Tab-delimited format! &nbsp;' +
        '  <input type="radio" id="%id.clipboardformat-csv" name="%id.clipboardformat" onclick="%s.SpreadsheetControlClipboardFormat(\'csv\');"> %loc!CSV format! &nbsp;' +
        '  <input type="radio" id="%id.clipboardformat-scsave" name="%id.clipboardformat" onclick="%s.SpreadsheetControlClipboardFormat(\'scsave\');"> %loc!SocialCalc-save format!' +
        ' </div>' +
        ' <input type="button" value="%loc!Load SocialCalc Clipboard With This!" style="font-size:x-small;" onclick="%s.SpreadsheetControlClipboardLoad();">&nbsp; ' +
        ' <input type="button" value="%loc!Clear SocialCalc Clipboard!" style="font-size:x-small;" onclick="%s.SpreadsheetControlClipboardClear();">&nbsp; ' +
        ' <br>' +
        ' <textarea id="%id.clipboardtext" style="font-size:small;height:350px;width:800px;overflow:auto;" onfocus="%s.CmdGotFocus(this);"></textarea>'
};

    return;

};


/**
 * Methods - Prototype method definitions for SocialCalc.SpreadsheetControl
 */

/**
 * Initialize the spreadsheet control
 * @param {Node|string} node - DOM node or element ID to attach to
 * @param {number} height - Requested height
 * @param {number} width - Requested width  
 * @param {number} spacebelow - Space to leave below
 * @returns {*} Result from SocialCalc.InitializeSpreadsheetControl
 */
SocialCalc.SpreadsheetControl.prototype.InitializeSpreadsheetControl = function(node, height, width, spacebelow) {
    return SocialCalc.InitializeSpreadsheetControl(this, node, height, width, spacebelow);
};

/**
 * Handle resize events
 * @returns {*} Result from SocialCalc.DoOnResize
 */
SocialCalc.SpreadsheetControl.prototype.DoOnResize = function() {
    return SocialCalc.DoOnResize(this);
};

/**
 * Size the spreadsheet div element
 * @returns {*} Result from SocialCalc.SizeSSDiv
 */
SocialCalc.SpreadsheetControl.prototype.SizeSSDiv = function() {
    return SocialCalc.SizeSSDiv(this);
};

/**
 * Execute a spreadsheet command
 * @param {string} combostr - Command combination string
 * @param {string} sstr - Additional string parameter
 * @returns {*} Result from SocialCalc.SpreadsheetControlExecuteCommand
 */
SocialCalc.SpreadsheetControl.prototype.ExecuteCommand = function(combostr, sstr) {
    return SocialCalc.SpreadsheetControlExecuteCommand(this, combostr, sstr);
};

/**
 * Create HTML for the sheet
 * @returns {*} Result from SocialCalc.SpreadsheetControlCreateSheetHTML
 */
SocialCalc.SpreadsheetControl.prototype.CreateSheetHTML = function() {
    return SocialCalc.SpreadsheetControlCreateSheetHTML(this);
};

/**
 * Create spreadsheet save data
 * @param {*} otherparts - Additional parts to include in save
 * @returns {*} Result from SocialCalc.SpreadsheetControlCreateSpreadsheetSave
 */
SocialCalc.SpreadsheetControl.prototype.CreateSpreadsheetSave = function(otherparts) {
    return SocialCalc.SpreadsheetControlCreateSpreadsheetSave(this, otherparts);
};

/**
 * Decode spreadsheet save data
 * @param {string} str - Save data string to decode
 * @returns {*} Result from SocialCalc.SpreadsheetControlDecodeSpreadsheetSave
 */
SocialCalc.SpreadsheetControl.prototype.DecodeSpreadsheetSave = function(str) {
    return SocialCalc.SpreadsheetControlDecodeSpreadsheetSave(this, str);
};

/**
 * Create HTML for a specific cell
 * @param {string} coord - Cell coordinate
 * @returns {*} Result from SocialCalc.SpreadsheetControlCreateCellHTML
 */
SocialCalc.SpreadsheetControl.prototype.CreateCellHTML = function(coord) {
    return SocialCalc.SpreadsheetControlCreateCellHTML(this, coord);
};

/**
 * Create HTML save data for a cell range
 * @param {string} range - Cell range
 * @returns {*} Result from SocialCalc.SpreadsheetControlCreateCellHTMLSave
 */
SocialCalc.SpreadsheetControl.prototype.CreateCellHTMLSave = function(range) {
    return SocialCalc.SpreadsheetControlCreateCellHTMLSave(this, range);
};

/**
 * Sheet Methods to make things a little easier
 */

/**
 * Parse sheet save data
 * @param {string} str - Sheet save string to parse
 * @returns {*} Result from sheet.ParseSheetSave
 */
SocialCalc.SpreadsheetControl.prototype.ParseSheetSave = function(str) {
    return this.sheet.ParseSheetSave(str);
};

/**
 * Create sheet save data
 * @returns {*} Result from sheet.CreateSheetSave
 */
SocialCalc.SpreadsheetControl.prototype.CreateSheetSave = function() {
    return this.sheet.CreateSheetSave();
};

/**
 * Functions
 */

/**
 * InitializeSpreadsheetControl - Creates the control elements and makes them the child of node
 * 
 * @param {SocialCalc.SpreadsheetControl} spreadsheet - The spreadsheet control object
 * @param {Node|string} node - DOM node or element ID to attach to
 * @param {number} height - Requested height (0 or null for maximum that fits)
 * @param {number} width - Requested width (0 or null for maximum that fits)
 * @param {number} spacebelow - Space to leave below for other content
 * 
 * Creates the control elements and makes them the child of node (string or element).
 * If present, height and width specify size.
 * If either is 0 or null (missing), the maximum that fits on the screen
 * (taking spacebelow into account) is used.
 * 
 * Displays the tabs and creates the views (other than "sheet").
 * The first tab is set as selected, but onclick is not invoked.
 * 
 * You should do a redisplay or recalc (which redisplays) after running this.
 */
SocialCalc.InitializeSpreadsheetControl = function(spreadsheet, node, height, width, spacebelow) {
    const scc = SocialCalc.Constants;
    const SCLoc = SocialCalc.LocalizeString;
    const SCLocSS = SocialCalc.LocalizeSubstrings;

    let html, child, i, vname, v, style, button, bele;
    const tabs = spreadsheet.tabs;
    const views = spreadsheet.views;

    spreadsheet.requestedHeight = height;
    spreadsheet.requestedWidth = width;
    spreadsheet.requestedSpaceBelow = spacebelow;

    if (typeof node === "string") {
        node = document.getElementById(node);
    }

    if (node === null) {
        alert("SocialCalc.SpreadsheetControl not given parent node.");
    }

    spreadsheet.parentNode = node;

    // Create node to hold spreadsheet control
    spreadsheet.spreadsheetDiv = document.createElement("div");

    spreadsheet.SizeSSDiv(); // calculate and fill in the size values

    for (child = node.firstChild; child !== null; child = node.firstChild) {
        node.removeChild(child);
    }

    // Create the tabbed UI at the top
    html = `<div><div style="${spreadsheet.toolbarbackground}padding:12px 10px 10px 4px;height:40px;">`;

    for (i = 0; i < tabs.length; i++) {
        html += tabs[i].html;
    }

    html += `</div><div style="${spreadsheet.tabbackground}padding-bottom:4px;margin:0px 0px 8px 0px;"><table cellpadding="0" cellspacing="0"><tr>`;

    for (i = 0; i < tabs.length; i++) {
        html += `  <td id="%id.${tabs[i].name}tab" style="${i === 0 ? spreadsheet.tabselectedCSS : spreadsheet.tabplainCSS}" onclick="%s.SetTab(this);">${SCLoc(tabs[i].text)}</td>`;
    }

    html += ' </tr></table></div></div>';

    spreadsheet.currentTab = 0; // this is where we started

    for (style in spreadsheet.tabreplacements) {
        html = html.replace(spreadsheet.tabreplacements[style].regex, spreadsheet.tabreplacements[style].replacement);
    }
    html = html.replace(/\%s\./g, "SocialCalc.");
    html = html.replace(/\%id\./g, spreadsheet.idPrefix);
    html = html.replace(/\%tbt\./g, spreadsheet.toolbartext);
    html = html.replace(/\%img\./g, spreadsheet.imagePrefix);

    html = SCLocSS(html); // localize with %loc!string! and %scc!constant!

    spreadsheet.spreadsheetDiv.innerHTML = html;

    node.appendChild(spreadsheet.spreadsheetDiv);

    /**
     * Initialize SocialCalc buttons configuration
     * @type {Object}
     */
    spreadsheet.Buttons = {
        button_undo: { tooltip: "Undo", command: "undo" },
        button_redo: { tooltip: "Redo", command: "redo" },
        button_copy: { tooltip: "Copy", command: "copy" },
        button_cut: { tooltip: "Cut", command: "cut" },
        button_paste: { tooltip: "Paste", command: "paste" },
        button_pasteformats: { tooltip: "Paste Formats", command: "pasteformats" },
        button_delete: { tooltip: "Delete Contents", command: "delete" },
        button_filldown: { tooltip: "Fill Down", command: "filldown" },
        button_fillright: { tooltip: "Fill Right", command: "fillright" },
        button_movefrom: { tooltip: "Set/Clear Move From", command: "movefrom" },
        button_movepaste: { tooltip: "Move Paste", command: "movepaste" },
        button_moveinsert: { tooltip: "Move Insert", command: "moveinsert" },
        button_alignleft: { tooltip: "Align Left", command: "align-left" },
        button_aligncenter: { tooltip: "Align Center", command: "align-center" },
        button_alignright: { tooltip: "Align Right", command: "align-right" },
        button_borderon: { tooltip: "Borders On", command: "borderon" },
        button_borderoff: { tooltip: "Borders Off", command: "borderoff" },
        button_swapcolors: { tooltip: "Swap Colors", command: "swapcolors" },
        button_merge: { tooltip: "Merge Cells", command: "merge" },
        button_unmerge: { tooltip: "Unmerge Cells", command: "unmerge" },
        button_insertrow: { tooltip: "Insert Row", command: "insertrow" },
        button_insertcol: { tooltip: "Insert Column", command: "insertcol" },
        button_deleterow: { tooltip: "Delete Row", command: "deleterow" },
        button_deletecol: { tooltip: "Delete Column", command: "deletecol" },
        button_recalc: { tooltip: "Recalc", command: "recalc" }
    };

    for (button in spreadsheet.Buttons) {
        bele = document.getElementById(spreadsheet.idPrefix + button);
        if (!bele) {
            alert(`Button ${spreadsheet.idPrefix + button} missing`);
            continue;
        }
        bele.style.border = `1px solid ${scc.ISCButtonBorderNormal}`;
        SocialCalc.TooltipRegister(bele, SCLoc(spreadsheet.Buttons[button].tooltip), {});
        SocialCalc.ButtonRegister(bele, {
            normalstyle: `border:1px solid ${scc.ISCButtonBorderNormal};backgroundColor:${scc.ISCButtonBorderNormal};`,
            hoverstyle: `border:1px solid ${scc.ISCButtonBorderHover};backgroundColor:${scc.ISCButtonBorderNormal};`,
            downstyle: `border:1px solid ${scc.ISCButtonBorderDown};backgroundColor:${scc.ISCButtonDownBackground};`
        }, {
            MouseDown: SocialCalc.DoButtonCmd,
            command: spreadsheet.Buttons[button].command
        });
    }

    // Create formula bar
    spreadsheet.formulabarDiv = document.createElement("div");
    spreadsheet.formulabarDiv.style.height = `${spreadsheet.formulabarheight}px`;
    spreadsheet.formulabarDiv.innerHTML = '<input type="text" size="60" value="" disabled="true">&nbsp;';
    spreadsheet.spreadsheetDiv.appendChild(spreadsheet.formulabarDiv);
    const inputbox = new SocialCalc.InputBox(spreadsheet.formulabarDiv.firstChild, spreadsheet.editor);

    for (button in spreadsheet.formulabuttons) {
        bele = document.createElement("img");
        bele.id = spreadsheet.idPrefix + button;
        bele.src = spreadsheet.imagePrefix + spreadsheet.formulabuttons[button].image;
        bele.style.verticalAlign = "middle";
        bele.style.border = "1px solid #FFF";
        bele.style.marginLeft = "4px";
        SocialCalc.TooltipRegister(bele, SCLoc(spreadsheet.formulabuttons[button].tooltip), {});
        SocialCalc.ButtonRegister(bele, {
            normalstyle: "border:1px solid #FFF;backgroundColor:#FFF;",
            hoverstyle: "border:1px solid #CCC;backgroundColor:#FFF;",
            downstyle: "border:1px solid #000;backgroundColor:#FFF;"
        }, {
            MouseDown: spreadsheet.formulabuttons[button].command
        });
        spreadsheet.formulabarDiv.appendChild(bele);
    }
// Initialize tabs that need it
for (let i = 0; i < tabs.length; i++) { // execute any tab-specific initialization code
    if (tabs[i].oncreate) {
        tabs[i].oncreate(spreadsheet, tabs[i].name);
    }
}

// Create sheet view and others
if (!scc.doWorkBook) {
    spreadsheet.nonviewheight = spreadsheet.statuslineheight +
        spreadsheet.spreadsheetDiv.firstChild.offsetHeight +
        spreadsheet.spreadsheetDiv.lastChild.offsetHeight;
} else {
    spreadsheet.nonviewheight = spreadsheet.sheetbarheight +
        spreadsheet.spreadsheetDiv.firstChild.offsetHeight +
        spreadsheet.spreadsheetDiv.lastChild.offsetHeight;
}

spreadsheet.viewheight = spreadsheet.height - spreadsheet.nonviewheight;
spreadsheet.editorDiv = spreadsheet.editor.CreateTableEditor(spreadsheet.width, spreadsheet.viewheight);

spreadsheet.spreadsheetDiv.appendChild(spreadsheet.editorDiv);

for (vname in views) {
    html = views[vname].html;
    for (style in views[vname].replacements) {
        html = html.replace(views[vname].replacements[style].regex, views[vname].replacements[style].replacement);
    }
    html = html.replace(/\%s\./g, "SocialCalc.");
    html = html.replace(/\%id\./g, spreadsheet.idPrefix);
    html = html.replace(/\%tbt\./g, spreadsheet.toolbartext);
    html = html.replace(/\%img\./g, spreadsheet.imagePrefix);
    
    v = document.createElement("div");
    SocialCalc.setStyles(v, views[vname].divStyle);
    v.style.display = "none";
    v.style.width = `${spreadsheet.width}px`;
    v.style.height = `${spreadsheet.viewheight}px`;

    html = SCLocSS(html); // localize with %loc!string!, etc.

    v.innerHTML = html;
    spreadsheet.spreadsheetDiv.appendChild(v);
    views[vname].element = v;
    if (views[vname].oncreate) {
        views[vname].oncreate(spreadsheet, views[vname]);
    }
}

views.sheet = { name: "sheet", element: spreadsheet.editorDiv };

// Create statusline
if (!scc.doWorkBook) {
    spreadsheet.statuslineDiv = document.createElement("div");
    spreadsheet.statuslineDiv.style.cssText = spreadsheet.statuslineCSS;
    // spreadsheet.statuslineDiv.style.height = spreadsheet.statuslineheight + "px"; // didn't take padding into account!
    spreadsheet.statuslineDiv.style.height = `${spreadsheet.statuslineheight -
        (spreadsheet.statuslineDiv.style.paddingTop.slice(0, -2) - 0) -
        (spreadsheet.statuslineDiv.style.paddingBottom.slice(0, -2) - 0)}px`;
    spreadsheet.statuslineDiv.id = `${spreadsheet.idPrefix}statusline`;
    spreadsheet.spreadsheetDiv.appendChild(spreadsheet.statuslineDiv);
} else {
    SocialCalc.CreateSheetStatusBar(spreadsheet, scc);
}

// Done - refresh screen needed
return;

}; // End of SocialCalc.InitializeSpreadsheetControl

/**
 * Create sheet status bar for workbook mode
 * @param {SocialCalc.SpreadsheetControl} spreadsheet - The spreadsheet control object
 * @param {Object} scc - SocialCalc constants object
 */
SocialCalc.CreateSheetStatusBar = function(spreadsheet, scc) {
    // Create sheetbar
    if (!scc.doWorkBook) {
        return;
    }
    
    // Create a table with 1 row, containing 3 columns: 1 for sheetbar, 1 for separator, 1 for statusline
    spreadsheet.sheetstatusbarDiv = document.createElement("div");
    spreadsheet.sheetstatusbarDiv.style.height = `${spreadsheet.sheetbarheight + 3}px`;
    spreadsheet.sheetstatusbarDiv.style.backgroundColor = "#CCC";
    spreadsheet.sheetstatusbarDiv.id = `${spreadsheet.idPrefix}sheetstatusbar`;

    spreadsheet.sheetbarDiv = document.createElement("div");
    // spreadsheet.sheetbarDiv.style.cssText = spreadsheet.sheetbarCSS;
    spreadsheet.sheetbarDiv.id = `${spreadsheet.idPrefix}sheetbar`;

    spreadsheet.statuslineDiv = document.createElement("div");
    spreadsheet.statuslineDiv.style.cssText = spreadsheet.statuslineCSS;
    spreadsheet.statuslineDiv.id = `${spreadsheet.idPrefix}statusline`;

    const table = document.createElement("table");
    spreadsheet.sheetstatusbartable = table;
    table.cellSpacing = 0;
    table.cellPadding = 0;
    table.width = "100%";
    
    const tbody = document.createElement("tbody");
    table.appendChild(tbody);

    const tr = document.createElement("tr");
    tbody.appendChild(tr);
    
    let td = document.createElement("td");
    td.appendChild(spreadsheet.sheetbarDiv);
    td.width = scc.SCSheetBarWidth;
    tr.appendChild(td);

    td = document.createElement("td");
    td.innerHTML = "<span>&nbsp|&nbsp</span>";
    td.width = "1%";
    tr.appendChild(td);

    td = document.createElement("td");
    td.appendChild(spreadsheet.statuslineDiv);
    tr.appendChild(td);

    spreadsheet.sheetstatusbarDiv.appendChild(table);
    spreadsheet.spreadsheetDiv.appendChild(spreadsheet.sheetstatusbarDiv);
};

/**
 * SocialCalc function to make localization easier
 * 
 * @param {string} str - Text to localize (e.g., "Text to localize")
 * @returns {string} Localized string from SocialCalc.Constants.s_loc_text_to_localize if it exists, or original string
 * 
 * If str is "Text to localize", it returns SocialCalc.Constants.s_loc_text_to_localize if
 * it exists, or else with just "Text to localize".
 * Note that spaces are replaced with "_" and other special
 * chars with "X" in the name of the constant (e.g., "A & B"
 * would look for SocialCalc.Constants.s_loc_a_X_b.
 */
SocialCalc.LocalizeString = function(str) {
    let cstr = SocialCalc.LocalizeStringList[str]; // found already this session?
    if (!cstr) { // no - look up
        cstr = SocialCalc.Constants[`s_loc_${str.toLowerCase().replace(/\s/g, "_").replace(/\W/g, "X")}`] || str;
        SocialCalc.LocalizeStringList[str] = cstr;
    }
    return cstr;
};

/**
 * A list of strings to localize accumulated by the LocalizeString routine
 * @type {Object}
 */
SocialCalc.LocalizeStringList = {};

/**
 * SocialCalc function to make localization easier using %loc and %ssc
 * 
 * @param {string} str - String containing localization placeholders
 * @returns {string} String with localization placeholders replaced
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
 * If the constant doesn't exist, throws an alert.
 */
SocialCalc.LocalizeSubstrings = function(str) {
    const SCLoc = SocialCalc.LocalizeString;

    return str.replace(/%(loc|ssc)!(.*?)!/g, (a, t, c) => {
        if (t === "ssc") {
            return SocialCalc.Constants[c] || alert(`Missing constant: ${c}`);
        } else {
            return SCLoc(c);
        }
    });
};
/**
 * Get the current spreadsheet control object
 * @returns {SocialCalc.SpreadsheetControl|null} The current spreadsheet control object or null if none exists
 */
SocialCalc.GetSpreadsheetControlObject = function() {
    const csco = SocialCalc.CurrentSpreadsheetControlObject;
    if (csco) return csco;

    // throw ("No current SpreadsheetControl object.");
};

/**
 * Process an onResize event, setting the different views
 * @param {SocialCalc.SpreadsheetControl} spreadsheet - The spreadsheet control object
 */
SocialCalc.DoOnResize = function(spreadsheet) {
    let v;
    const views = spreadsheet.views;

    const needresize = spreadsheet.SizeSSDiv();
    if (!needresize) return;

    for (const vname in views) {
        v = views[vname].element;
        v.style.width = `${spreadsheet.width}px`;
        v.style.height = `${spreadsheet.height - spreadsheet.nonviewheight}px`;
    }

    spreadsheet.editor.ResizeTableEditor(spreadsheet.width, spreadsheet.height - spreadsheet.nonviewheight);
};

/**
 * Figure out a reasonable size for the spreadsheet, given any requested values and viewport
 * Sets ssdiv to that size
 * @param {SocialCalc.SpreadsheetControl} spreadsheet - The spreadsheet control object
 * @returns {boolean} True if different than existing values, false otherwise
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
    if (spreadsheet.height !== newval) {
        spreadsheet.height = newval;
        spreadsheet.spreadsheetDiv.style.height = `${newval}px`;
        resized = true;
    }
    
    newval = spreadsheet.requestedWidth ||
        sizes.width - (pos.left + pos.right + fudgefactorX) || 700;
    if (spreadsheet.width !== newval) {
        spreadsheet.width = newval;
        spreadsheet.spreadsheetDiv.style.width = `${newval}px`;
        resized = true;
    }

    return resized;
};

/**
 * Set the active tab in the spreadsheet control
 * @param {string|Element} obj - Either a string with the tab name or a DOM element with an ID
 */
SocialCalc.SetTab = function(obj) {
    let newtab, tname, newtabnum, newview, i, vname, ele;
    const menutabs = {};
    const tools = {};
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const tabs = spreadsheet.tabs;
    const views = spreadsheet.views;

    if (typeof obj === "string") {
        newtab = obj;
    } else {
        newtab = obj.id.slice(spreadsheet.idPrefix.length, -3);
    }

    if (spreadsheet.editor.busy && // if busy and switching from "sheet", ignore
        (!tabs[spreadsheet.currentTab].view || tabs[spreadsheet.currentTab].view === "sheet")) {
        for (i = 0; i < tabs.length; i++) {
            if (tabs[i].name === newtab && (tabs[i].view && tabs[i].view !== "sheet")) {
                return;
            }
        }
    }

    if (spreadsheet.tabs[spreadsheet.currentTab].onunclick) {
        spreadsheet.tabs[spreadsheet.currentTab].onunclick(spreadsheet, spreadsheet.tabs[spreadsheet.currentTab].name);
    }

    for (i = 0; i < tabs.length; i++) {
        tname = tabs[i].name;
        menutabs[tname] = document.getElementById(`${spreadsheet.idPrefix}${tname}tab`);
        tools[tname] = document.getElementById(`${spreadsheet.idPrefix}${tname}tools`);
        if (tname === newtab) {
            newtabnum = i;
            tools[tname].style.display = "block";
            menutabs[tname].style.cssText = spreadsheet.tabselectedCSS;
        } else {
            tools[tname].style.display = "none";
            menutabs[tname].style.cssText = spreadsheet.tabplainCSS;
        }
    }

    spreadsheet.currentTab = newtabnum;

    if (tabs[newtabnum].onclick) {
        tabs[newtabnum].onclick(spreadsheet, newtab);
    }

    for (vname in views) {
        if ((!tabs[newtabnum].view && vname === "sheet") || tabs[newtabnum].view === vname) {
            views[vname].element.style.display = "block";
            newview = vname;
        } else {
            views[vname].element.style.display = "none";
        }
    }

    if (tabs[newtabnum].onclickFocus) {
        ele = tabs[newtabnum].onclickFocus;
        if (typeof ele === "string") {
            ele = document.getElementById(`${spreadsheet.idPrefix}${ele}`);
            ele.focus();
        }
        SocialCalc.CmdGotFocus(ele);
    } else {
        SocialCalc.KeyboardFocus();
    }

    if (views[newview].needsresize && views[newview].onresize) {
        views[newview].needsresize = false;
        views[newview].onresize(spreadsheet, views[newview]);
    }

    if (newview === "sheet") {
        spreadsheet.statuslineDiv.style.display = "block";
        spreadsheet.editor.ScheduleRender();
    } else {
        spreadsheet.statuslineDiv.style.display = "none";
    }

    return;
};

/**
 * Statusline callback for spreadsheet control
 * @param {Object} editor - The editor object
 * @param {string} status - Status string
 * @param {*} arg - Additional argument
 * @param {Object} params - Parameters object containing statuslineid, recalcid1, recalcid2
 */
SocialCalc.SpreadsheetControlStatuslineCallback = function(editor, status, arg, params) {
    let rele1, rele2;

    const ele = document.getElementById(params.statuslineid);

    if (ele) {
        ele.innerHTML = editor.GetStatuslineString(status, arg, params);
    }

    switch (status) {
        case "cmdendnorender":
        case "calcfinished":
        case "doneposcalc":
            rele1 = document.getElementById(params.recalcid1);
            rele2 = document.getElementById(params.recalcid2);
            if (!rele1 || !rele2) break;
            if (editor.context.sheetobj.attribs.needsrecalc === "yes") {
                rele1.style.display = "inline";
                rele2.style.display = "inline";
            } else {
                rele1.style.display = "none";
                rele2.style.display = "none";
            }
            break;

        default:
            break;
    }
};
/**
 * Update sort range proposal in the UI
 * Updates sort range proposed in the UI in element idPrefix+sortlist
 * @param {Object} editor - The editor object
 */
SocialCalc.UpdateSortRangeProposal = function(editor) {
    const ele = document.getElementById(`${SocialCalc.GetSpreadsheetControlObject().idPrefix}sortlist`);
    if (editor.range.hasrange) {
        ele.options[0].text = `${SocialCalc.crToCoord(editor.range.left, editor.range.top)}:${SocialCalc.crToCoord(editor.range.right, editor.range.bottom)}`;
    } else {
        ele.options[0].text = SocialCalc.LocalizeString("[select range]");
    }
};

/**
 * Load column choosers for sort functionality
 * Updates list of columns for choosing which to sort for Major, Minor, and Last sort
 * @param {SocialCalc.SpreadsheetControl} spreadsheet - The spreadsheet control object
 */
SocialCalc.LoadColumnChoosers = function(spreadsheet) {
    const SCLoc = SocialCalc.LocalizeString;

    let sortrange, nrange, rparts, col, colname, sele, oldindex;

    if (spreadsheet.sortrange && spreadsheet.sortrange.indexOf(":") === -1) { // sortrange is a named range
        nrange = SocialCalc.Formula.LookupName(spreadsheet.sheet, spreadsheet.sortrange || "");
        if (nrange.type === "range") {
            rparts = nrange.value.match(/^(.*)\|(.*)\|$/);
            sortrange = `${rparts[1]}:${rparts[2]}`;
        } else {
            sortrange = "A1:A1";
        }
    } else {
        sortrange = spreadsheet.sortrange;
    }
    
    const range = SocialCalc.ParseRange(sortrange);
    
    // Major sort selector
    sele = document.getElementById(`${spreadsheet.idPrefix}majorsort`);
    oldindex = sele.selectedIndex;
    sele.options.length = 0;
    sele.options[sele.options.length] = new Option(SCLoc("[None]"), "");
    for (col = range.cr1.col; col <= range.cr2.col; col++) {
        colname = SocialCalc.rcColname(col);
        sele.options[sele.options.length] = new Option(`${SCLoc("Column ")}${colname}`, colname);
    }
    sele.selectedIndex = oldindex > 1 && oldindex <= (range.cr2.col - range.cr1.col + 1) ? oldindex : 1; // restore what was there if reasonable
    
    // Minor sort selector
    sele = document.getElementById(`${spreadsheet.idPrefix}minorsort`);
    oldindex = sele.selectedIndex;
    sele.options.length = 0;
    sele.options[sele.options.length] = new Option(SCLoc("[None]"), "");
    for (col = range.cr1.col; col <= range.cr2.col; col++) {
        colname = SocialCalc.rcColname(col);
        sele.options[sele.options.length] = new Option(colname, colname);
    }
    sele.selectedIndex = oldindex > 0 && oldindex <= (range.cr2.col - range.cr1.col + 1) ? oldindex : 0; // default to [none]
    
    // Last sort selector
    sele = document.getElementById(`${spreadsheet.idPrefix}lastsort`);
    oldindex = sele.selectedIndex;
    sele.options.length = 0;
    sele.options[sele.options.length] = new Option(SCLoc("[None]"), "");
    for (col = range.cr1.col; col <= range.cr2.col; col++) {
        colname = SocialCalc.rcColname(col);
        sele.options[sele.options.length] = new Option(colname, colname);
    }
    sele.selectedIndex = oldindex > 0 && oldindex <= (range.cr2.col - range.cr1.col + 1) ? oldindex : 0; // default to [none]
};

/**
 * Set keyboard pass-through focus
 * Sets SocialCalc.Keyboard.passThru: obj should be element with focus or "true"
 * @param {Element|boolean} obj - Element with focus or true
 */
SocialCalc.CmdGotFocus = function(obj) {
    SocialCalc.Keyboard.passThru = obj;
};

/**
 * Handle button command execution
 * @param {Event} e - Event object
 * @param {Object} buttoninfo - Button information object
 * @param {Object} bobj - Button object containing element and function info
 */
SocialCalc.DoButtonCmd = function(e, buttoninfo, bobj) {
    SocialCalc.DoCmd(bobj.element, bobj.functionobj.command);
};

/**
 * Execute a spreadsheet command
 * @param {Element} obj - DOM element that triggered the command
 * @param {string} which - Command identifier
 */
SocialCalc.DoCmd = function(obj, which) {
    let combostr, sstr, cl, i, clele, slist, slistele, str, sele, rele, lele, ele, sortrange, nrange, rparts;
    let sheet, cell, color, bgcolor, defaultcolor, defaultbgcolor;

    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const editor = spreadsheet.editor;

    switch (which) {
        case "undo":
            spreadsheet.ExecuteCommand("undo", "");
            break;

        case "redo":
            spreadsheet.ExecuteCommand("redo", "");
            break;

        case "fill-rowcolstuff":
        case "fill-text":
            cl = which.substring(5);
            clele = document.getElementById(`${spreadsheet.idPrefix}${cl}list`);
            clele.length = 0;
            for (i = 0; i < SocialCalc.SpreadsheetCmdTable[cl].length; i++) {
                clele.options[i] = new Option(SocialCalc.SpreadsheetCmdTable[cl][i].t);
            }
            which = `changed-${cl}`; // fall through to changed code

        case "changed-rowcolstuff":
        case "changed-text":
            cl = which.substring(8);
            clele = document.getElementById(`${spreadsheet.idPrefix}${cl}list`);
            slist = SocialCalc.SpreadsheetCmdTable.slists[SocialCalc.SpreadsheetCmdTable[cl][clele.selectedIndex].s]; // get sList for this command
            slistele = document.getElementById(`${spreadsheet.idPrefix}${cl}slist`);
            slistele.length = 0; // reset
            for (i = 0; i < (slist.length || 0); i++) {
                slistele.options[i] = new Option(slist[i].t, slist[i].s);
            }
            return; // nothing else to do

        case "ok-rowcolstuff":
        case "ok-text":
            cl = which.substring(3);
            clele = document.getElementById(`${spreadsheet.idPrefix}${cl}list`);
            slistele = document.getElementById(`${spreadsheet.idPrefix}${cl}slist`);
            combostr = SocialCalc.SpreadsheetCmdTable[cl][clele.selectedIndex].c;
            sstr = slistele[slistele.selectedIndex].value;
            SocialCalc.SpreadsheetControlExecuteCommand(obj, combostr, sstr);
            break;

        case "ok-setsort":
            lele = document.getElementById(`${spreadsheet.idPrefix}sortlist`);
            if (lele.selectedIndex === 0) {
                if (editor.range.hasrange) {
                    spreadsheet.sortrange = `${SocialCalc.crToCoord(editor.range.left, editor.range.top)}:${SocialCalc.crToCoord(editor.range.right, editor.range.bottom)}`;
                } else {
                    spreadsheet.sortrange = `${editor.ecell.coord}:${editor.ecell.coord}`;
                }
            } else {
                spreadsheet.sortrange = lele.options[lele.selectedIndex].value;
            }
            ele = document.getElementById(`${spreadsheet.idPrefix}sortbutton`);
            ele.value = `${SocialCalc.LocalizeString("Sort ")}${spreadsheet.sortrange}`;
            ele.style.visibility = "visible";
            SocialCalc.LoadColumnChoosers(spreadsheet);
            if (obj && obj.blur) obj.blur();
            SocialCalc.KeyboardFocus();
            return;

        case "dosort":
            if (spreadsheet.sortrange && spreadsheet.sortrange.indexOf(":") === -1) { // sortrange is a named range
                nrange = SocialCalc.Formula.LookupName(spreadsheet.sheet, spreadsheet.sortrange || "");
                if (nrange.type !== "range") return;
                rparts = nrange.value.match(/^(.*)\|(.*)\|$/);
                sortrange = `${rparts[1]}:${rparts[2]}`;
            } else {
                sortrange = spreadsheet.sortrange;
            }
            if (sortrange === "A1:A1") return;
            str = `sort ${sortrange} `;
            sele = document.getElementById(`${spreadsheet.idPrefix}majorsort`);
            rele = document.getElementById(`${spreadsheet.idPrefix}majorsortup`);
            str += `${sele.options[sele.selectedIndex].value}${rele.checked ? " up" : " down"}`;
            sele = document.getElementById(`${spreadsheet.idPrefix}minorsort`);
            if (sele.selectedIndex > 0) {
                rele = document.getElementById(`${spreadsheet.idPrefix}minorsortup`);
                str += ` ${sele.options[sele.selectedIndex].value}${rele.checked ? " up" : " down"}`;
            }
            sele = document.getElementById(`${spreadsheet.idPrefix}lastsort`);
            if (sele.selectedIndex > 0) {
                rele = document.getElementById(`${spreadsheet.idPrefix}lastsortup`);
                str += ` ${sele.options[sele.selectedIndex].value}${rele.checked ? " up" : " down"}`;
            }
            spreadsheet.ExecuteCommand(str, "");
            break;

        case "merge":
            combostr = SocialCalc.SpreadsheetCmdLookup[which] || "";
            sstr = SocialCalc.SpreadsheetCmdSLookup[which] || "";
            spreadsheet.ExecuteCommand(combostr, sstr);
            if (editor.range.hasrange) { // set ecell to upper left
                editor.MoveECell(SocialCalc.crToCoord(editor.range.left, editor.range.top));
                editor.RangeRemove();
            }
            break;
        case "movefrom":
            if (editor.range2.hasrange) { // toggle if already there
                spreadsheet.context.cursorsuffix = "";
                editor.Range2Remove();
                spreadsheet.ExecuteCommand("redisplay", "");
            } else if (editor.range.hasrange) { // set range2 to range or one cell
                editor.range2.top = editor.range.top;
                editor.range2.right = editor.range.right;
                editor.range2.bottom = editor.range.bottom;
                editor.range2.left = editor.range.left;
                editor.range2.hasrange = true;
                editor.MoveECell(SocialCalc.crToCoord(editor.range.left, editor.range.top));
            } else {
                editor.range2.top = editor.ecell.row;
                editor.range2.right = editor.ecell.col;
                editor.range2.bottom = editor.ecell.row;
                editor.range2.left = editor.ecell.col;
                editor.range2.hasrange = true;
            }
            str = editor.range2.hasrange ? "" : "off";
            ele = document.getElementById(`${spreadsheet.idPrefix}button_movefrom`);
            ele.src = `${spreadsheet.imagePrefix}movefrom${str}.gif`;
            ele = document.getElementById(`${spreadsheet.idPrefix}button_movepaste`);
            ele.src = `${spreadsheet.imagePrefix}movepaste${str}.gif`;
            ele = document.getElementById(`${spreadsheet.idPrefix}button_moveinsert`);
            ele.src = `${spreadsheet.imagePrefix}moveinsert${str}.gif`;
            if (editor.range2.hasrange) editor.RangeRemove();
            break;

        case "movepaste":
        case "moveinsert":
            if (editor.range2.hasrange) {
                spreadsheet.context.cursorsuffix = "";
                combostr = `${which} ${SocialCalc.crToCoord(editor.range2.left, editor.range2.top)}:${SocialCalc.crToCoord(editor.range2.right, editor.range2.bottom)} ${editor.ecell.coord}`;
                spreadsheet.ExecuteCommand(combostr, "");
                editor.Range2Remove();
                ele = document.getElementById(`${spreadsheet.idPrefix}button_movefrom`);
                ele.src = `${spreadsheet.imagePrefix}movefromoff.gif`;
                ele = document.getElementById(`${spreadsheet.idPrefix}button_movepaste`);
                ele.src = `${spreadsheet.imagePrefix}movepasteoff.gif`;
                ele = document.getElementById(`${spreadsheet.idPrefix}button_moveinsert`);
                ele.src = `${spreadsheet.imagePrefix}moveinsertoff.gif`;
            }
            break;

        case "swapcolors":
            sheet = spreadsheet.sheet;
            cell = sheet.GetAssuredCell(editor.ecell.coord);
            defaultcolor = sheet.attribs.defaultcolor ? sheet.colors[sheet.attribs.defaultcolor] : "rgb(0,0,0)";
            defaultbgcolor = sheet.attribs.defaultbgcolor ? sheet.colors[sheet.attribs.defaultbgcolor] : "rgb(255,255,255)";
            color = cell.color ? sheet.colors[cell.color] : defaultcolor; // get color
            if (color === defaultbgcolor) color = ""; // going to swap, so if same as background default, use default
            bgcolor = cell.bgcolor ? sheet.colors[cell.bgcolor] : defaultbgcolor;
            if (bgcolor === defaultcolor) bgcolor = ""; // going to swap, so if same as foreground default, use default
            spreadsheet.ExecuteCommand(`set %C color ${bgcolor}%Nset %C bgcolor ${color}`, "");
            break;

        default:
            combostr = SocialCalc.SpreadsheetCmdLookup[which] || "";
            sstr = SocialCalc.SpreadsheetCmdSLookup[which] || "";
            spreadsheet.ExecuteCommand(combostr, sstr);
            break;
    }

    if (obj && obj.blur) obj.blur();
    SocialCalc.KeyboardFocus();
}; // End of SocialCalc.DoCmd

/**
 * Command lookup table for spreadsheet operations
 * Maps command names to their corresponding command strings
 * @type {Object}
 */
SocialCalc.SpreadsheetCmdLookup = {
    'copy': 'copy %C all',
    'cut': 'cut %C all',
    'paste': 'paste %C all',
    'pasteformats': 'paste %C formats',
    'delete': 'erase %C formulas',
    'filldown': 'filldown %C all',
    'fillright': 'fillright %C all',
    'erase': 'erase %C all',
    'borderon': 'set %C bt %S%Nset %C br %S%Nset %C bb %S%Nset %C bl %S',
    'borderoff': 'set %C bt %S%Nset %C br %S%Nset %C bb %S%Nset %C bl %S',
    'merge': 'merge %C',
    'unmerge': 'unmerge %C',
    'align-left': 'set %C cellformat left',
    'align-center': 'set %C cellformat center',
    'align-right': 'set %C cellformat right',
    'align-default': 'set %C cellformat',
    'insertrow': 'insertrow %C',
    'insertcol': 'insertcol %C',
    'deleterow': 'deleterow %C',
    'deletecol': 'deletecol %C',
    'undo': 'undo',
    'redo': 'redo',
    'recalc': 'recalc'
};

/**
 * Secondary command lookup table for spreadsheet operations
 * Maps command names to their corresponding secondary parameter strings
 * @type {Object}
 */
SocialCalc.SpreadsheetCmdSLookup = {
    'borderon': '1px solid rgb(0,0,0)',
    'borderoff': ''
};

/**
 * NO LONGER USED
 * 
 * Legacy command table structure that was previously used for dynamic command generation.
 * This large configuration object defined various command categories, their display text,
 * parameter lists, and command strings. It has been replaced by the simpler lookup tables above.
 * 
 * The structure included:
 * - cmd: General commands array
 * - rowcolstuff: Row/column specific commands
 * - text: Text formatting commands
 * - slists: Sub-lists for various parameter types (colors, fonts, formats, etc.)
 * 
 * Preserved here for reference and potential future use.
 */

/*
SocialCalc.SpreadsheetCmdTable = {
    cmd: [
        {t:"Fill Right", s:"ffal", c:"fillright %C %S"},
        {t:"Fill Down", s:"ffal", c:"filldown %C %S"},
        {t:"Copy", s:"all", c:"copy %C %S"},
        {t:"Cut", s:"all", c:"cut %C %S"},
        {t:"Paste", s:"ffal", c:"paste %C %S"},
        {t:"Erase", s:"ffal", c:"erase %C %S"},
        {t:"Insert", s:"rowcol", c:"insert%S %C"},
        {t:"Delete", s:"rowcol", c:"delete%S %C"},
        {t:"Merge Cells", s:"none", c:"merge %C"},
        {t:"Unmerge", s:"none", c:"unmerge %C"},
        {t:"Sort", s:"sortcol", c:"sort %R %S"},
        {t:"Cell Color", s:"colors", c:"set %C color %S"},
        {t:"Cell Background", s:"colors", c:"set %C bgcolor %S"},
        {t:"Cell Number Format", s:"ntvf", c:"set %C nontextvalueformat %S"},
        {t:"Cell Font", s:"fonts", c:"set %C font %S"},
        {t:"Cell Align", s:"cellformat", c:"set %C cellformat %S"},
        {t:"Cell Borders", s:"borderOnOff", c:"set %C bt %S%Nset %C br %S%Nset %C bb %S%Nset %C bl %S"},
        {t:"Column Width", s:"colWidths", c:"set %W width %S"},
        {t:"Default Color", s:"colors", c:"set sheet defaultcolor %S"},
        {t:"Default Background", s:"colors", c:"set sheet defaultbgcolor %S"},
        {t:"Default Number Format", s:"ntvf", c:"set sheet defaultnontextvalueformat %S"},
        {t:"Default Font", s:"fonts", c:"set sheet defaultfont %S"},
        {t:"Default Text Align", s:"cellformat", c:"set sheet defaulttextformat %S"},
        {t:"Default Number Align", s:"cellformat", c:"set sheet defaultnontextformat %S"},
        {t:"Default Column Width", s:"colWidths", c:"set sheet defaultcolwidth %S"}
    ],
    // ... (rest of the legacy structure would continue here)
};
*/
/**
 * Execute a spreadsheet command with parameter substitution
 * @param {Element} obj - DOM element that triggered the command
 * @param {string} combostr - Command string with placeholders (%C, %R, %N, %S, %W, %P)
 * @param {string} sstr - Secondary string parameter
 */
SocialCalc.SpreadsheetControlExecuteCommand = function(obj, combostr, sstr) {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const eobj = spreadsheet.editor;

    const str = {
        P: "%",
        N: "\n"
    };
    
    if (eobj.range.hasrange) {
        str.R = `${SocialCalc.crToCoord(eobj.range.left, eobj.range.top)}:${SocialCalc.crToCoord(eobj.range.right, eobj.range.bottom)}`;
        str.C = str.R;
        str.W = `${SocialCalc.rcColname(eobj.range.left)}:${SocialCalc.rcColname(eobj.range.right)}`;
    } else {
        str.C = eobj.ecell.coord;
        str.R = `${eobj.ecell.coord}:${eobj.ecell.coord}`;
        str.W = SocialCalc.rcColname(SocialCalc.coordToCr(eobj.ecell.coord).col);
    }
    str.S = sstr;
    
    combostr = combostr.replace(/%C/g, str.C);
    combostr = combostr.replace(/%R/g, str.R);
    combostr = combostr.replace(/%N/g, str.N);
    combostr = combostr.replace(/%S/g, str.S);
    combostr = combostr.replace(/%W/g, str.W);
    combostr = combostr.replace(/%P/g, str.P);

    eobj.EditorScheduleSheetCommands(combostr, true, false);
};

/**
 * Create HTML representation of the whole spreadsheet
 * @param {SocialCalc.SpreadsheetControl} spreadsheet - The spreadsheet control object
 * @returns {string} HTML representation of the spreadsheet
 */
SocialCalc.SpreadsheetControlCreateSheetHTML = function(spreadsheet) {
    const context = new SocialCalc.RenderContext(spreadsheet.sheet);
    const div = document.createElement("div");
    const ele = context.RenderSheet(null, { type: "html" });
    div.appendChild(ele);
    
    const result = div.innerHTML;
    
    // Cleanup
    delete context;
    delete ele;
    delete div;
    
    return result;
};

/**
 * Create HTML representation of a cell
 * @param {SocialCalc.SpreadsheetControl} spreadsheet - The spreadsheet control object
 * @param {string} coord - Cell coordinate
 * @param {string} [linkstyle] - Link style to use (optional)
 * @returns {string} HTML representation of the cell. Blank is "", not "&nbsp;"
 */
SocialCalc.SpreadsheetControlCreateCellHTML = function(spreadsheet, coord, linkstyle) {
    let result = "";
    const cell = spreadsheet.sheet.cells[coord];

    if (!cell) return "";

    if (cell.displaystring === undefined) {
        result = SocialCalc.FormatValueForDisplay(
            spreadsheet.sheet, 
            cell.datavalue, 
            coord, 
            (linkstyle || spreadsheet.context.defaultHTMLlinkstyle)
        );
    } else {
        result = cell.displaystring;
    }

    if (result === "&nbsp;") result = "";

    return result;
};

/**
 * Create HTML representation of a range of cells or the whole sheet
 * @param {SocialCalc.SpreadsheetControl} spreadsheet - The spreadsheet control object
 * @param {string|null} range - Range string or null for whole sheet
 * @param {string} [linkstyle] - Link style to use (optional)
 * @returns {string} HTML representation in save format with version header and encoded cell data
 * 
 * Returns the HTML representation of a range of cells, or the whole sheet if range is null.
 * The form is:
 *    version:1.0
 *    coord:cell-HTML
 *    coord:cell-HTML
 *    ...
 *
 * Empty cells are skipped. The cell-HTML is encoded with ":"=>"\c", newline=>"\n", and "\"=>"\b".
 */
SocialCalc.SpreadsheetControlCreateCellHTMLSave = function(spreadsheet, range, linkstyle) {
    let cr1, cr2, row, col, coord, cell, cellHTML;
    const result = [];
    let prange;

    if (range) {
        prange = SocialCalc.ParseRange(range);
    } else {
        prange = {
            cr1: { row: 1, col: 1 },
            cr2: { 
                row: spreadsheet.sheet.attribs.lastrow, 
                col: spreadsheet.sheet.attribs.lastcol 
            }
        };
    }
    
    cr1 = prange.cr1;
    cr2 = prange.cr2;

    result.push("version:1.0");

    for (row = cr1.row; row <= cr2.row; row++) {
        for (col = cr1.col; col <= cr2.col; col++) {
            coord = SocialCalc.crToCoord(col, row);
            cell = spreadsheet.sheet.cells[coord];
            if (!cell) continue;
            
            if (cell.displaystring === undefined) {
                cellHTML = SocialCalc.FormatValueForDisplay(
                    spreadsheet.sheet, 
                    cell.datavalue, 
                    coord, 
                    (linkstyle || spreadsheet.context.defaultHTMLlinkstyle)
                );
            } else {
                cellHTML = cell.displaystring;
            }
            
            if (cellHTML === "&nbsp;") continue;
            result.push(`${coord}:${SocialCalc.encodeForSave(cellHTML)}`);
        }
    }

    result.push(""); // one extra to get extra \n
    return result.join("\n");
};

/**
 * Formula Bar Button Routines
 */

/**
 * Display the function list dialog
 * Shows a dialog with function categories, function names, and descriptions
 */
SocialCalc.SpreadsheetControl.DoFunctionList = function() {
    const scf = SocialCalc.Formula;
    const scc = SocialCalc.Constants;
    const fcl = scc.function_classlist;

    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const idp = `${spreadsheet.idPrefix}function`;

    let ele = document.getElementById(`${idp}dialog`);
    if (ele) return; // already have one

    scf.FillFunctionInfo();

    let str = `<table><tr><td><span style="font-size:x-small;font-weight:bold">%loc!Category!</span><br>` +
        `<select id="${idp}class" size="${fcl.length}" style="width:120px;" onchange="SocialCalc.SpreadsheetControl.FunctionClassChosen(this.options[this.selectedIndex].value);">`;
    
    for (let i = 0; i < fcl.length; i++) {
        str += `<option value="${fcl[i]}"${i === 0 ? ' selected>' : '>'}${SocialCalc.special_chars(scf.FunctionClasses[fcl[i]].name)}</option>`;
    }
    
    str += `</select></td><td>&nbsp;&nbsp;</td><td id="${idp}list"><span style="font-size:x-small;font-weight:bold">%loc!Functions!</span><br>` +
        `<select id="${idp}name" size="${fcl.length}" style="width:240px;" ` +
        `onchange="SocialCalc.SpreadsheetControl.FunctionChosen(this.options[this.selectedIndex].value);" ondblclick="SocialCalc.SpreadsheetControl.DoFunctionPaste();">`;
    
    str += SocialCalc.SpreadsheetControl.GetFunctionNamesStr("all");
    str += `</td></tr><tr><td colspan="3">` +
        `<div id="${idp}desc" style="width:380px;height:80px;overflow:auto;font-size:x-small;">${SocialCalc.SpreadsheetControl.GetFunctionInfoStr(scf.FunctionClasses[fcl[0]].items[0])}</div>` +
        `<div style="width:380px;text-align:right;padding-top:6px;font-size:small;">` +
        `<input type="button" value="%loc!Paste!" style="font-size:smaller;" onclick="SocialCalc.SpreadsheetControl.DoFunctionPaste();">&nbsp;` +
        `<input type="button" value="%loc!Cancel!" style="font-size:smaller;" onclick="SocialCalc.SpreadsheetControl.HideFunctions();"></div>` +
        `</td></tr></table>`;

    const main = document.createElement("div");
    main.id = `${idp}dialog`;
    main.style.position = "absolute";

    const vp = SocialCalc.GetViewportInfo();
    main.style.top = `${vp.height / 3}px`;
    main.style.left = `${vp.width / 3}px`;
    main.style.zIndex = 100;
    main.style.backgroundColor = "#FFF";
    main.style.border = "1px solid black";
    main.style.width = "400px";

    str = `<table cellspacing="0" cellpadding="0" style="border-bottom:1px solid black;"><tr>` +
        `<td style="font-size:10px;cursor:default;width:100%;background-color:#999;color:#FFF;">&nbsp;%loc!Function List!</td>` +
        `<td style="font-size:10px;cursor:default;color:#666;" onclick="SocialCalc.SpreadsheetControl.HideFunctions();">&nbsp;X&nbsp;</td></tr></table>` +
        `<div style="background-color:#DDD;">${str}</div>`;

    str = SocialCalc.LocalizeSubstrings(str);
    main.innerHTML = str;

    SocialCalc.DragRegister(main.firstChild.firstChild.firstChild.firstChild, true, true, {
        MouseDown: SocialCalc.DragFunctionStart, 
        MouseMove: SocialCalc.DragFunctionPosition,
        MouseUp: SocialCalc.DragFunctionPosition,
        Disabled: null, 
        positionobj: main
    });

    spreadsheet.spreadsheetDiv.appendChild(main);

    ele = document.getElementById(`${idp}name`);
    ele.focus();
    SocialCalc.CmdGotFocus(ele);
    // TODO: need to do keyboard handling: if esc, hide; if All, letter scrolls to there
};

/**
 * Get HTML option string for function names in a category
 * @param {string} cname - Category name
 * @returns {string} HTML options string for select element
 */
SocialCalc.SpreadsheetControl.GetFunctionNamesStr = function(cname) {
    const scf = SocialCalc.Formula;
    let str = "";

    const f = scf.FunctionClasses[cname];
    for (let i = 0; i < f.items.length; i++) {
        str += `<option value="${f.items[i]}"${i === 0 ? ' selected>' : '>'}${f.items[i]}</option>`;
    }

    return str;
};

/**
 * Fill function names into a select element
 * @param {string} cname - Category name
 * @param {HTMLSelectElement} ele - Select element to populate
 */
SocialCalc.SpreadsheetControl.FillFunctionNames = function(cname, ele) {
    const scf = SocialCalc.Formula;

    ele.length = 0;
    const f = scf.FunctionClasses[cname];
    for (let i = 0; i < f.items.length; i++) {
        ele.options[i] = new Option(f.items[i], f.items[i]);
        if (i === 0) {
            ele.options[i].selected = true;
        }
    }
};
/**
 * Get function information string for display
 * @param {string} fname - Function name
 * @returns {string} HTML formatted function information with signature and description
 */
SocialCalc.SpreadsheetControl.GetFunctionInfoStr = function(fname) {
    const scf = SocialCalc.Formula;
    const f = scf.FunctionList[fname];
    const scsc = SocialCalc.special_chars;

    let str = `<b>${fname}(${scsc(scf.FunctionArgString(fname))})</b><br>`;
    str += scsc(f[3]);

    return str;
};

/**
 * Handle function category selection
 * @param {string} cname - Category name that was chosen
 */
SocialCalc.SpreadsheetControl.FunctionClassChosen = function(cname) {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const idp = `${spreadsheet.idPrefix}function`;
    const scf = SocialCalc.Formula;

    SocialCalc.SpreadsheetControl.FillFunctionNames(cname, document.getElementById(`${idp}name`));
    SocialCalc.SpreadsheetControl.FunctionChosen(scf.FunctionClasses[cname].items[0]);
};

/**
 * Handle function selection and update description
 * @param {string} fname - Function name that was chosen
 */
SocialCalc.SpreadsheetControl.FunctionChosen = function(fname) {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const idp = `${spreadsheet.idPrefix}function`;

    document.getElementById(`${idp}desc`).innerHTML = SocialCalc.SpreadsheetControl.GetFunctionInfoStr(fname);
};

/**
 * Hide the function list dialog
 */
SocialCalc.SpreadsheetControl.HideFunctions = function() {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();

    const ele = document.getElementById(`${spreadsheet.idPrefix}functiondialog`);
    ele.innerHTML = "";

    SocialCalc.DragUnregister(ele);
    SocialCalc.KeyboardFocus();

    if (ele.parentNode) {
        ele.parentNode.removeChild(ele);
    }
};

/**
 * Paste selected function into input with opening parenthesis
 */
SocialCalc.SpreadsheetControl.DoFunctionPaste = function() {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const editor = spreadsheet.editor;
    const ele = document.getElementById(`${spreadsheet.idPrefix}functionname`);
    const mele = document.getElementById(`${spreadsheet.idPrefix}multilinetextarea`);

    const text = `${ele.value}(`;

    SocialCalc.SpreadsheetControl.HideFunctions();

    if (mele) { // multi-line editing is in progress
        mele.value += text;
        mele.focus();
        SocialCalc.CmdGotFocus(mele);
    } else {
        editor.EditorAddToInput(text, "=");
    }
};

/**
 * Display multi-line input dialog
 * Shows a textarea for editing cell contents with multiple lines
 */
SocialCalc.SpreadsheetControl.DoMultiline = function() {
    const SCLocSS = SocialCalc.LocalizeSubstrings;

    let str, ele, text;

    const scc = SocialCalc.Constants;
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const editor = spreadsheet.editor;
    const wval = editor.workingvalues;

    const idp = `${spreadsheet.idPrefix}multiline`;

    ele = document.getElementById(`${idp}dialog`);
    if (ele) return; // already have one

    switch (editor.state) {
        case "start":
            wval.ecoord = editor.ecell.coord;
            wval.erow = editor.ecell.row;
            wval.ecol = editor.ecell.col;
            editor.RangeRemove();
            text = SocialCalc.GetCellContents(editor.context.sheetobj, wval.ecoord);
            break;

        case "input":
        case "inputboxdirect":
            text = editor.inputBox.GetText();
            break;
    }

    editor.inputBox.element.disabled = true;

    text = SocialCalc.special_chars(text);

    str = `<textarea id="${idp}textarea" style="width:380px;height:120px;margin:10px 0px 0px 6px;">${text}</textarea>` +
        `<div style="width:380px;text-align:right;padding:6px 0px 4px 6px;font-size:small;">` +
        SCLocSS('<input type="button" value="%loc!Set Cell Contents!" style="font-size:smaller;" onclick="SocialCalc.SpreadsheetControl.DoMultilinePaste();">&nbsp;' +
        '<input type="button" value="%loc!Clear!" style="font-size:smaller;" onclick="SocialCalc.SpreadsheetControl.DoMultilineClear();">&nbsp;' +
        '<input type="button" value="%loc!Cancel!" style="font-size:smaller;" onclick="SocialCalc.SpreadsheetControl.HideMultiline();"></div>' +
        '</div>');

    const main = document.createElement("div");
    main.id = `${idp}dialog`;
    main.style.position = "absolute";

    const vp = SocialCalc.GetViewportInfo();
    main.style.top = `${vp.height / 3}px`;
    main.style.left = `${vp.width / 3}px`;
    main.style.zIndex = 100;
    main.style.backgroundColor = "#FFF";
    main.style.border = "1px solid black";
    main.style.width = "400px";

    main.innerHTML = '<table cellspacing="0" cellpadding="0" style="border-bottom:1px solid black;"><tr>' +
        '<td style="font-size:10px;cursor:default;width:100%;background-color:#999;color:#FFF;">' +
        SCLocSS("&nbsp;%loc!Multi-line Input Box!") + '</td>' +
        '<td style="font-size:10px;cursor:default;color:#666;" onclick="SocialCalc.SpreadsheetControl.HideMultiline();">&nbsp;X&nbsp;</td></tr></table>' +
        `<div style="background-color:#DDD;">${str}</div>`;

    SocialCalc.DragRegister(main.firstChild.firstChild.firstChild.firstChild, true, true, {
        MouseDown: SocialCalc.DragFunctionStart, 
        MouseMove: SocialCalc.DragFunctionPosition,
        MouseUp: SocialCalc.DragFunctionPosition,
        Disabled: null, 
        positionobj: main
    });

    spreadsheet.spreadsheetDiv.appendChild(main);

    ele = document.getElementById(`${idp}textarea`);
    ele.focus();
    SocialCalc.CmdGotFocus(ele);
    // TODO: need to do keyboard handling: if esc, hide?
};

/**
 * Hide the multi-line input dialog and restore input state
 */
SocialCalc.SpreadsheetControl.HideMultiline = function() {
    const scc = SocialCalc.Constants;
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const editor = spreadsheet.editor;

    const ele = document.getElementById(`${spreadsheet.idPrefix}multilinedialog`);
    ele.innerHTML = "";

    SocialCalc.DragUnregister(ele);
    SocialCalc.KeyboardFocus();

    if (ele.parentNode) {
        ele.parentNode.removeChild(ele);
    }

    switch (editor.state) {
        case "start":
            editor.inputBox.DisplayCellContents(null);
            break;

        case "input":
        case "inputboxdirect":
            editor.inputBox.element.disabled = false;
            editor.inputBox.Focus();
            break;
    }
};

/**
 * Clear the multi-line textarea content
 */
SocialCalc.SpreadsheetControl.DoMultilineClear = function() {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();

    const ele = document.getElementById(`${spreadsheet.idPrefix}multilinetextarea`);

    ele.value = "";
    ele.focus();
};
/**
 * Paste content from multi-line textarea into the cell
 * Handles the actual setting of cell content from the multi-line input dialog
 */
SocialCalc.SpreadsheetControl.DoMultilinePaste = function() {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const editor = spreadsheet.editor;
    const wval = editor.workingvalues;

    const ele = document.getElementById(`${spreadsheet.idPrefix}multilinetextarea`);
    const text = ele.value;

    SocialCalc.SpreadsheetControl.HideMultiline();

    switch (editor.state) {
        case "start":
            wval.partialexpr = "";
            wval.ecoord = editor.ecell.coord;
            wval.erow = editor.ecell.row;
            wval.ecol = editor.ecell.col;
            break;
        case "input":
        case "inputboxdirect":
            editor.inputBox.Blur();
            editor.inputBox.ShowInputBox(false);
            editor.state = "start";
            break;
    }

    editor.EditorSaveEdit(text);
};

/**
 * Display link input dialog
 * Shows a dialog for creating and editing hyperlinks in cells
 */
SocialCalc.SpreadsheetControl.DoLink = function() {
    const SCLoc = SocialCalc.LocalizeString;

    let str, ele, text, cell, setformat, popup;

    const scc = SocialCalc.Constants;
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const editor = spreadsheet.editor;
    const wval = editor.workingvalues;

    const idp = `${spreadsheet.idPrefix}link`;

    ele = document.getElementById(`${idp}dialog`);
    if (ele) return; // already have one

    switch (editor.state) {
        case "start":
            wval.ecoord = editor.ecell.coord;
            wval.erow = editor.ecell.row;
            wval.ecol = editor.ecell.col;
            editor.RangeRemove();
            text = SocialCalc.GetCellContents(editor.context.sheetobj, wval.ecoord);
            break;

        case "input":
        case "inputboxdirect":
            text = editor.inputBox.GetText();
            break;
    }

    editor.inputBox.element.disabled = true;

    if (text.charAt(0) === "'") {
        text = text.slice(1);
    }

    const parts = SocialCalc.ParseCellLinkText(text);

    text = SocialCalc.special_chars(text);

    cell = spreadsheet.sheet.cells[editor.ecell.coord];
    if (!cell || !cell.textvalueformat) { // set to link format, but don't override
        setformat = " checked";
    } else {
        setformat = "";
    }

    popup = parts.newwin ? " checked" : "";

    str = `<div style="padding:6px 0px 4px 6px;">` +
        `<span style="font-size:smaller;">${SCLoc("Description")}</span><br>` +
        `<input type="text" id="${idp}desc" style="width:380px;" value="${SocialCalc.special_chars(parts.desc)}"><br>` +
        `<span style="font-size:smaller;">${SCLoc("URL")}</span><br>` +
        `<input type="text" id="${idp}url" style="width:380px;" value="${SocialCalc.special_chars(parts.url)}"><br>`;
    
    if (SocialCalc.Callbacks.MakePageLink) { // only show if handling pagenames here
        str += `<span style="font-size:smaller;">${SCLoc("Page Name")}</span><br>` +
            `<input type="text" id="${idp}pagename" style="width:380px;" value="${SocialCalc.special_chars(parts.pagename)}"><br>` +
            `<span style="font-size:smaller;">${SCLoc("Workspace")}</span><br>` +
            `<input type="text" id="${idp}workspace" style="width:380px;" value="${SocialCalc.special_chars(parts.workspace)}"><br>`;
    }
    
    str += SocialCalc.LocalizeSubstrings(`<input type="checkbox" id="${idp}format"${setformat}>&nbsp;` +
        `<span style="font-size:smaller;">%loc!Set to Link format!</span><br>` +
        `<input type="checkbox" id="${idp}popup"${popup}>&nbsp;` +
        `<span style="font-size:smaller;">%loc!Show in new browser window!</span>` +
        `</div>` +
        `<div style="width:380px;text-align:right;padding:6px 0px 4px 6px;font-size:small;">` +
        `<input type="button" value="%loc!Set Cell Contents!" style="font-size:smaller;" onclick="SocialCalc.SpreadsheetControl.DoLinkPaste();">&nbsp;` +
        `<input type="button" value="%loc!Clear!" style="font-size:smaller;" onclick="SocialCalc.SpreadsheetControl.DoLinkClear();">&nbsp;` +
        `<input type="button" value="%loc!Cancel!" style="font-size:smaller;" onclick="SocialCalc.SpreadsheetControl.HideLink();"></div>` +
        `</div>`);

    const main = document.createElement("div");
    main.id = `${idp}dialog`;
    main.style.position = "absolute";

    const vp = SocialCalc.GetViewportInfo();
    main.style.top = `${vp.height / 3}px`;
    main.style.left = `${vp.width / 3}px`;
    main.style.zIndex = 100;
    main.style.backgroundColor = "#FFF";
    main.style.border = "1px solid black";
    main.style.width = "400px";

    main.innerHTML = '<table cellspacing="0" cellpadding="0" style="border-bottom:1px solid black;"><tr>' +
        '<td style="font-size:10px;cursor:default;width:100%;background-color:#999;color:#FFF;">' +
        `&nbsp;${SCLoc("Link Input Box")}</td>` +
        '<td style="font-size:10px;cursor:default;color:#666;" onclick="SocialCalc.SpreadsheetControl.HideLink();">&nbsp;X&nbsp;</td></tr></table>' +
        `<div style="background-color:#DDD;">${str}</div>`;

    SocialCalc.DragRegister(main.firstChild.firstChild.firstChild.firstChild, true, true, {
        MouseDown: SocialCalc.DragFunctionStart, 
        MouseMove: SocialCalc.DragFunctionPosition,
        MouseUp: SocialCalc.DragFunctionPosition,
        Disabled: null, 
        positionobj: main
    });

    spreadsheet.spreadsheetDiv.appendChild(main);

    ele = document.getElementById(`${idp}url`);
    ele.focus();
    SocialCalc.CmdGotFocus(ele);
    // TODO: need to do keyboard handling: if esc, hide?
};

/**
 * Hide the link input dialog and restore editor state
 */
SocialCalc.SpreadsheetControl.HideLink = function() {
    const scc = SocialCalc.Constants;
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const editor = spreadsheet.editor;

    const ele = document.getElementById(`${spreadsheet.idPrefix}linkdialog`);
    ele.innerHTML = "";

    SocialCalc.DragUnregister(ele);
    SocialCalc.KeyboardFocus();

    if (ele.parentNode) {
        ele.parentNode.removeChild(ele);
    }

    switch (editor.state) {
        case "start":
            editor.inputBox.DisplayCellContents(null);
            break;

        case "input":
        case "inputboxdirect":
            editor.inputBox.element.disabled = false;
            editor.inputBox.Focus();
            break;
    }
};

/**
 * Clear all link input fields and focus on URL field
 */
SocialCalc.SpreadsheetControl.DoLinkClear = function() {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();

    document.getElementById(`${spreadsheet.idPrefix}linkdesc`).value = "";
    document.getElementById(`${spreadsheet.idPrefix}linkpagename`).value = "";
    document.getElementById(`${spreadsheet.idPrefix}linkworkspace`).value = "";

    const ele = document.getElementById(`${spreadsheet.idPrefix}linkurl`);
    ele.value = "";
    ele.focus();
};
/**
 * Paste link data from the link dialog into the cell
 * Constructs the appropriate link text format and saves it to the cell
 */
SocialCalc.SpreadsheetControl.DoLinkPaste = function() {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const editor = spreadsheet.editor;
    const wval = editor.workingvalues;

    const descele = document.getElementById(`${spreadsheet.idPrefix}linkdesc`);
    const urlele = document.getElementById(`${spreadsheet.idPrefix}linkurl`);
    const pagenameele = document.getElementById(`${spreadsheet.idPrefix}linkpagename`);
    const workspaceele = document.getElementById(`${spreadsheet.idPrefix}linkworkspace`);
    const formatele = document.getElementById(`${spreadsheet.idPrefix}linkformat`);
    const popupele = document.getElementById(`${spreadsheet.idPrefix}linkpopup`);

    let text = "";
    let ltsym, gtsym, obsym, cbsym;

    if (popupele.checked) {
        ltsym = "<<"; 
        gtsym = ">>"; 
        obsym = "[["; 
        cbsym = "]]";
    } else {
        ltsym = "<"; 
        gtsym = ">"; 
        obsym = "["; 
        cbsym = "]";
    }

    if (pagenameele && pagenameele.value) {
        if (workspaceele.value) {
            text = `${descele.value}{${workspaceele.value}${obsym}${pagenameele.value}${cbsym}}`;
        } else {
            text = `${descele.value}${obsym}${pagenameele.value}${cbsym}`;
        }
    } else {
        text = `${descele.value}${ltsym}${urlele.value}${gtsym}`;
    }

    SocialCalc.SpreadsheetControl.HideLink();

    switch (editor.state) {
        case "start":
            wval.partialexpr = "";
            wval.ecoord = editor.ecell.coord;
            wval.erow = editor.ecell.row;
            wval.ecol = editor.ecell.col;
            break;
        case "input":
        case "inputboxdirect":
            editor.inputBox.Blur();
            editor.inputBox.ShowInputBox(false);
            editor.state = "start";
            break;
    }

    if (formatele.checked) {
        SocialCalc.SpreadsheetControlExecuteCommand(null, "set %C textvalueformat text-link", "");
    }

    editor.EditorSaveEdit(text);
};

/**
 * Automatically create a SUM formula
 * Creates a SUM formula based on the current selection or automatically detects
 * the range above the current cell
 */
SocialCalc.SpreadsheetControl.DoSum = function() {
    let cmd, cell, row, col, sel, cr, foundvalue;

    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const editor = spreadsheet.editor;
    const sheet = editor.context.sheetobj;

    if (editor.range.hasrange) {
        sel = `${SocialCalc.crToCoord(editor.range.left, editor.range.top)}:${SocialCalc.crToCoord(editor.range.right, editor.range.bottom)}`;
        cmd = `set ${SocialCalc.crToCoord(editor.range.right, editor.range.bottom + 1)} formula sum(${sel})`;
    } else {
        row = editor.ecell.row - 1;
        col = editor.ecell.col;
        if (row <= 1) {
            cmd = `set ${editor.ecell.coord} constant e#REF! 0 #REF!`;
        } else {
            foundvalue = false;
            while (row > 0) {
                cr = SocialCalc.crToCoord(col, row);
                cell = sheet.GetAssuredCell(cr);
                if (!cell.datatype || cell.datatype === "t") {
                    if (foundvalue) {
                        row++;
                        break;
                    }
                } else {
                    foundvalue = true;
                }
                row--;
            }
            cmd = `set ${editor.ecell.coord} formula sum(${SocialCalc.crToCoord(col, row)}:${SocialCalc.crToCoord(col, editor.ecell.row - 1)})`;
        }
    }

    editor.EditorScheduleSheetCommands(cmd, true, false);
};

/**
 * TAB Routines
 */

/**
 * Sort tab click handler
 * Initializes the sort interface with available ranges and named ranges
 * @param {SocialCalc.SpreadsheetControl} s - Spreadsheet control object
 * @param {string} t - Tab name
 */
SocialCalc.SpreadsheetControlSortOnclick = function(s, t) {
    let name, i;
    const namelist = [];
    const nl = document.getElementById(`${s.idPrefix}sortlist`);
    
    SocialCalc.LoadColumnChoosers(s);
    s.editor.RangeChangeCallback.sort = SocialCalc.UpdateSortRangeProposal;

    for (name in s.sheet.names) {
        namelist.push(name);
    }
    namelist.sort();
    
    nl.length = 0;
    nl.options[0] = new Option(SocialCalc.LocalizeString("[select range]"));
    
    for (i = 0; i < namelist.length; i++) {
        name = namelist[i];
        nl.options[i + 1] = new Option(name, name);
        if (name === s.sortrange) {
            nl.options[i + 1].selected = true;
        }
    }
    
    if (s.sortrange === "") {
        nl.options[0].selected = true;
    }

    SocialCalc.UpdateSortRangeProposal(s.editor);
    SocialCalc.KeyboardFocus();
    return;
};

/**
 * Save sort settings to string format
 * @param {Object} editor - Editor object
 * @param {string} setting - Setting name
 * @returns {string} Formatted sort settings string
 * 
 * Format is: sort:sortrange:major:up/down:minor:up/down:last:up/down
 */
SocialCalc.SpreadsheetControlSortSave = function(editor, setting) {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    let str, sele, rele;

    str = `sort:${SocialCalc.encodeForSave(spreadsheet.sortrange)}:`;
    
    sele = document.getElementById(`${spreadsheet.idPrefix}majorsort`);
    rele = document.getElementById(`${spreadsheet.idPrefix}majorsortup`);
    str += `${sele.selectedIndex}${rele.checked ? ":up" : ":down"}`;
    
    sele = document.getElementById(`${spreadsheet.idPrefix}minorsort`);
    if (sele.selectedIndex > 0) {
        rele = document.getElementById(`${spreadsheet.idPrefix}minorsortup`);
        str += `:${sele.selectedIndex}${rele.checked ? ":up" : ":down"}`;
    } else {
        str += "::";
    }
    
    sele = document.getElementById(`${spreadsheet.idPrefix}lastsort`);
    if (sele.selectedIndex > 0) {
        rele = document.getElementById(`${spreadsheet.idPrefix}lastsortup`);
        str += `:${sele.selectedIndex}${rele.checked ? ":up" : ":down"}`;
    } else {
        str += "::";
    }
    
    return `${str}\n`;
};

/**
 * Load sort settings from string format
 * @param {Object} editor - Editor object
 * @param {string} setting - Setting name
 * @param {string} line - Settings line to parse
 * @param {Object} flags - Loading flags
 * @returns {boolean} True if successfully loaded
 */
SocialCalc.SpreadsheetControlSortLoad = function(editor, setting, line, flags) {
    let parts, ele, sele;

    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();

    parts = line.split(":");
    spreadsheet.sortrange = SocialCalc.decodeFromSave(parts[1]);
    
    ele = document.getElementById(`${spreadsheet.idPrefix}sortbutton`);
    if (spreadsheet.sortrange) {
        ele.value = `${SocialCalc.LocalizeString("Sort ")}${spreadsheet.sortrange}`;
        ele.style.visibility = "visible";
    } else {
        ele.style.visibility = "hidden";
    }
    
    SocialCalc.LoadColumnChoosers(spreadsheet);

    // Major sort settings
    sele = document.getElementById(`${spreadsheet.idPrefix}majorsort`);
    sele.selectedIndex = parts[2] - 0;
    document.getElementById(`${spreadsheet.idPrefix}majorsort${parts[3]}`).checked = true;
    
    // Minor sort settings
    sele = document.getElementById(`${spreadsheet.idPrefix}minorsort`);
    if (parts[4]) {
        sele.selectedIndex = parts[4] - 0;
        document.getElementById(`${spreadsheet.idPrefix}minorsort${parts[5]}`).checked = true;
    } else {
        sele.selectedIndex = 0;
        document.getElementById(`${spreadsheet.idPrefix}minorsortup`).checked = true;
    }
    
    // Last sort settings
    sele = document.getElementById(`${spreadsheet.idPrefix}lastsort`);
    if (parts[6]) {
        sele.selectedIndex = parts[6] - 0;
        document.getElementById(`${spreadsheet.idPrefix}lastsort${parts[7]}`).checked = true;
    } else {
        sele.selectedIndex = 0;
        document.getElementById(`${spreadsheet.idPrefix}lastsortup`).checked = true;
    }

    return true;
};

/**
 * Comment functionality
 */

/**
 * Handle comment tab click event
 * Sets up callback for cell movement and displays current cell's comment
 * @param {SocialCalc.SpreadsheetControl} s - Spreadsheet control object
 * @param {string} t - Tab name
 */
SocialCalc.SpreadsheetControlCommentOnclick = function(s, t) {
    s.editor.MoveECellCallback.comment = SocialCalc.SpreadsheetControlCommentMoveECell;
    SocialCalc.SpreadsheetControlCommentDisplay(s, t);
    SocialCalc.KeyboardFocus();
    return;
};

/**
 * Display the comment for the current cell in the comment textarea
 * @param {SocialCalc.SpreadsheetControl} s - Spreadsheet control object
 * @param {string} t - Tab name
 */
SocialCalc.SpreadsheetControlCommentDisplay = function(s, t) {
    let c = "";
    if (s.editor.ecell && s.editor.ecell.coord && s.sheet.cells[s.editor.ecell.coord]) {
        c = s.sheet.cells[s.editor.ecell.coord].comment || "";
    }
    document.getElementById(`${s.idPrefix}commenttext`).value = c;
};

/**
 * Callback for when the active cell changes in comment mode
 * Updates the displayed comment for the new cell
 * @param {Object} editor - Editor object
 */
SocialCalc.SpreadsheetControlCommentMoveECell = function(editor) {
    SocialCalc.SpreadsheetControlCommentDisplay(SocialCalc.GetSpreadsheetControlObject(), "comment");
};

/**
 * Set the comment for the current cell from the textarea value
 * Updates the cell's comment and refreshes the cell's CSS display
 */
SocialCalc.SpreadsheetControlCommentSet = function() {
    const s = SocialCalc.GetSpreadsheetControlObject();
    s.ExecuteCommand(`set %C comment ${SocialCalc.encodeForSave(document.getElementById(`${s.idPrefix}commenttext`).value)}`);
    const cell = SocialCalc.GetEditorCellElement(s.editor, s.editor.ecell.row, s.editor.ecell.col);
    s.editor.UpdateCellCSS(cell, s.editor.ecell.row, s.editor.ecell.col);
    SocialCalc.KeyboardFocus();
};

/**
 * Handle comment tab unclick event
 * Removes the cell movement callback
 * @param {SocialCalc.SpreadsheetControl} s - Spreadsheet control object
 * @param {string} t - Tab name
 */
SocialCalc.SpreadsheetControlCommentOnunclick = function(s, t) {
    delete s.editor.MoveECellCallback.comment;
};

/**
 * Names functionality
 */

/**
 * Handle names tab click event
 * Initializes the names interface and sets up callbacks
 * @param {SocialCalc.SpreadsheetControl} s - Spreadsheet control object
 * @param {string} t - Tab name
 */
SocialCalc.SpreadsheetControlNamesOnclick = function(s, t) {
    document.getElementById(`${s.idPrefix}namesname`).value = "";
    document.getElementById(`${s.idPrefix}namesdesc`).value = "";
    document.getElementById(`${s.idPrefix}namesvalue`).value = "";
    s.editor.RangeChangeCallback.names = SocialCalc.SpreadsheetControlNamesRangeChange;
    s.editor.MoveECellCallback.names = SocialCalc.SpreadsheetControlNamesRangeChange;
    SocialCalc.SpreadsheetControlNamesRangeChange(s.editor);
    SocialCalc.SpreadsheetControlNamesFillNameList();
    SocialCalc.SpreadsheetControlNamesChangedName();
};

/**
 * Fill the names list dropdown with existing named ranges
 * Populates the select element with all defined names, sorted alphabetically
 */
SocialCalc.SpreadsheetControlNamesFillNameList = function() {
    const SCLoc = SocialCalc.LocalizeString;
    let name, i;
    const namelist = [];
    const s = SocialCalc.GetSpreadsheetControlObject();
    const nl = document.getElementById(`${s.idPrefix}nameslist`);
    const currentname = document.getElementById(`${s.idPrefix}namesname`).value.toUpperCase().replace(/[^A-Z0-9_\.]/g, "");
    
    for (name in s.sheet.names) {
        namelist.push(name);
    }
    namelist.sort();
    
    nl.length = 0;
    if (namelist.length > 0) {
        nl.options[0] = new Option(SCLoc("[New]"));
    } else {
        nl.options[0] = new Option(SCLoc("[None]"));
    }
    
    for (i = 0; i < namelist.length; i++) {
        name = namelist[i];
        nl.options[i + 1] = new Option(name, name);
        if (name === currentname) {
            nl.options[i + 1].selected = true;
        }
    }
    
    if (currentname === "") {
        nl.options[0].selected = true;
    }
};

/**
 * Handle selection change in the names list
 * Updates the name input fields with the selected name's properties
 */
SocialCalc.SpreadsheetControlNamesChangedName = function() {
    const s = SocialCalc.GetSpreadsheetControlObject();
    const nl = document.getElementById(`${s.idPrefix}nameslist`);
    const name = nl.options[nl.selectedIndex].value;
    
    if (s.sheet.names[name]) {
        document.getElementById(`${s.idPrefix}namesname`).value = name;
        document.getElementById(`${s.idPrefix}namesdesc`).value = s.sheet.names[name].desc || "";
        document.getElementById(`${s.idPrefix}namesvalue`).value = s.sheet.names[name].definition || "";
    } else {
        document.getElementById(`${s.idPrefix}namesname`).value = "";
        document.getElementById(`${s.idPrefix}namesdesc`).value = "";
        document.getElementById(`${s.idPrefix}namesvalue`).value = "";
    }
};

/**
 * Update the range proposal button when selection or active cell changes
 * @param {Object} editor - Editor object
 */
SocialCalc.SpreadsheetControlNamesRangeChange = function(editor) {
    const s = SocialCalc.GetSpreadsheetControlObject();
    const ele = document.getElementById(`${s.idPrefix}namesrangeproposal`);
    
    if (editor.range.hasrange) {
        ele.value = `${SocialCalc.crToCoord(editor.range.left, editor.range.top)}:${SocialCalc.crToCoord(editor.range.right, editor.range.bottom)}`;
    } else {
        ele.value = editor.ecell.coord;
    }
};

/**
 * Handle names tab unclick event
 * Removes the range and cell movement callbacks
 * @param {SocialCalc.SpreadsheetControl} s - Spreadsheet control object
 * @param {string} t - Tab name
 */
SocialCalc.SpreadsheetControlNamesOnunclick = function(s, t) {
    delete s.editor.RangeChangeCallback.names;
    delete s.editor.MoveECellCallback.names;
};

/**
 * Set the value field to the current range proposal
 * Copies the range proposal button value to the names value input
 */
SocialCalc.SpreadsheetControlNamesSetValue = function() {
    const s = SocialCalc.GetSpreadsheetControlObject();
    document.getElementById(`${s.idPrefix}namesvalue`).value = document.getElementById(`${s.idPrefix}namesrangeproposal`).value;
    SocialCalc.KeyboardFocus();
};

/**
 * Save the current name definition
 * Creates or updates a named range with the current input values
 */
SocialCalc.SpreadsheetControlNamesSave = function() {
    const s = SocialCalc.GetSpreadsheetControlObject();
    const name = document.getElementById(`${s.idPrefix}namesname`).value;
    
    SocialCalc.SetTab(s.tabs[0].name); // return to first tab
    SocialCalc.KeyboardFocus();
    
    if (name !== "") {
        s.ExecuteCommand(`name define ${name} ${document.getElementById(`${s.idPrefix}namesvalue`).value}\nname desc ${name} ${document.getElementById(`${s.idPrefix}namesdesc`).value}`);
    }
};

/**
 * Delete the current name definition
 * Removes the selected named range from the spreadsheet
 */
SocialCalc.SpreadsheetControlNamesDelete = function() {
    const s = SocialCalc.GetSpreadsheetControlObject();
    const name = document.getElementById(`${s.idPrefix}namesname`).value;
    
    SocialCalc.SetTab(s.tabs[0].name); // return to first tab
    SocialCalc.KeyboardFocus();
    
    if (name !== "") {
        s.ExecuteCommand(`name delete ${name}`);
        // Commented out immediate UI updates - let the system handle refresh
        // document.getElementById(s.idPrefix+"namesname").value = "";
        // document.getElementById(s.idPrefix+"namesvalue").value = "";
        // document.getElementById(s.idPrefix+"namesdesc").value = "";
        // SocialCalc.SpreadsheetControlNamesFillNameList();
    }
    SocialCalc.KeyboardFocus();
};

/**
 * Clipboard functionality
 */

/**
 * Handle clipboard tab click event
 * Initializes the clipboard display with tab-delimited format
 * @param {SocialCalc.SpreadsheetControl} s - Spreadsheet control object
 * @param {string} t - Tab name
 */
SocialCalc.SpreadsheetControlClipboardOnclick = function(s, t) {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const clipele = document.getElementById(`${spreadsheet.idPrefix}clipboardtext`);
    document.getElementById(`${spreadsheet.idPrefix}clipboardformat-tab`).checked = true;
    clipele.value = SocialCalc.ConvertSaveToOtherFormat(SocialCalc.Clipboard.clipboard, "tab");
    return;
};

/**
 * Change clipboard display format
 * @param {string} which - Format type ("tab", "csv", or "scsave")
 */
SocialCalc.SpreadsheetControlClipboardFormat = function(which) {
    const s = SocialCalc.GetSpreadsheetControlObject();
    const clipele = document.getElementById(`${s.idPrefix}clipboardtext`);
    clipele.value = SocialCalc.ConvertSaveToOtherFormat(SocialCalc.Clipboard.clipboard, which);
};

/**
 * Load clipboard content into SocialCalc clipboard
 * Converts the textarea content to the appropriate format and loads it
 */
SocialCalc.SpreadsheetControlClipboardLoad = function() {
    const s = SocialCalc.GetSpreadsheetControlObject();
    let savetype = "tab";
    
    SocialCalc.SetTab(s.tabs[0].name); // return to first tab
    SocialCalc.KeyboardFocus();
    
    if (document.getElementById(`${s.idPrefix}clipboardformat-csv`).checked) {
        savetype = "csv";
    } else if (document.getElementById(`${s.idPrefix}clipboardformat-scsave`).checked) {
        savetype = "scsave";
    }
    
    s.editor.EditorScheduleSheetCommands(
        `loadclipboard ${SocialCalc.encodeForSave(
            SocialCalc.ConvertOtherFormatToSave(
                document.getElementById(`${s.idPrefix}clipboardtext`).value, 
                savetype
            )
        )}`, 
        true, 
        false
    );
};

/**
 * Clear the clipboard content and SocialCalc clipboard
 * Empties both the textarea and the internal clipboard
 */
SocialCalc.SpreadsheetControlClipboardClear = function() {
    const s = SocialCalc.GetSpreadsheetControlObject();
    const clipele = document.getElementById(`${s.idPrefix}clipboardtext`);
    
    clipele.value = "";
    s.editor.EditorScheduleSheetCommands("clearclipboard", true, false);
    clipele.focus();
};

/**
 * Export clipboard content using callback
 * Calls the export callback if one is defined
 */
SocialCalc.SpreadsheetControlClipboardExport = function() {
    const s = SocialCalc.GetSpreadsheetControlObject();
    
    if (s.ExportCallback) {
        s.ExportCallback(s);
    }
    
    SocialCalc.SetTab(s.tabs[0].name); // return to first tab
    SocialCalc.KeyboardFocus();
};
/**
 * Settings functionality
 */

/**
 * Switch between sheet and cell settings panels
 * @param {string} target - Either "sheet" or "cell" to indicate which panel to show
 */
SocialCalc.SpreadsheetControlSettingsSwitch = function(target) {
    SocialCalc.SettingControlReset();
    const s = SocialCalc.GetSpreadsheetControlObject();
    const sheettable = document.getElementById(`${s.idPrefix}sheetsettingstable`);
    const celltable = document.getElementById(`${s.idPrefix}cellsettingstable`);
    const sheettoolbar = document.getElementById(`${s.idPrefix}sheetsettingstoolbar`);
    const celltoolbar = document.getElementById(`${s.idPrefix}cellsettingstoolbar`);
    
    if (target === "sheet") {
        sheettable.style.display = "block";
        celltable.style.display = "none";
        sheettoolbar.style.display = "block";
        celltoolbar.style.display = "none";
        SocialCalc.SettingsControlSetCurrentPanel(s.views.settings.values.sheetspanel);
    } else {
        sheettable.style.display = "none";
        celltable.style.display = "block";
        sheettoolbar.style.display = "none";
        celltoolbar.style.display = "block";
        SocialCalc.SettingsControlSetCurrentPanel(s.views.settings.values.cellspanel);
    }
};

/**
 * Save settings from the current panel
 * @param {string} target - Either "sheet", "cell", or "cancel"
 */
SocialCalc.SettingsControlSave = function(target) {
    let range, cmdstr;
    const s = SocialCalc.GetSpreadsheetControlObject();
    const sc = SocialCalc.SettingsControls;
    const panelobj = sc.CurrentPanel;
    const attribs = SocialCalc.SettingsControlUnloadPanel(panelobj);

    SocialCalc.SetTab(s.tabs[0].name); // return to first tab
    SocialCalc.KeyboardFocus();

    if (target === "sheet") {
        cmdstr = s.sheet.DecodeSheetAttributes(attribs);
    } else if (target === "cell") {
        if (s.editor.range.hasrange) {
            range = `${SocialCalc.crToCoord(s.editor.range.left, s.editor.range.top)}:${SocialCalc.crToCoord(s.editor.range.right, s.editor.range.bottom)}`;
        }
        cmdstr = s.sheet.DecodeCellAttributes(s.editor.ecell.coord, attribs, range);
    } else { // Cancel
        // No action needed for cancel
    }
    
    if (cmdstr) {
        s.editor.EditorScheduleSheetCommands(cmdstr, true, false);
    }
};

/**
 * SAVE / LOAD ROUTINES
 */

/**
 * Create a complete spreadsheet save string in multi-part MIME format
 * 
 * @param {SocialCalc.SpreadsheetControl} spreadsheet - The spreadsheet control object
 * @param {Object|null} otherparts - Optional additional parts to include
 * @returns {string} Complete save string in MIME format
 * 
 * Saves the spreadsheet's sheet data, editor settings, and audit trail (redo stack).
 * The serialized data strings are concatenated together in multi-part MIME format.
 * The first part lists the types of the subsequent parts (e.g., "sheet", "editor", and "audit")
 * in this format:
 *   # comments
 *   version:1.0
 *   part:type1
 *   part:type2
 *   ...
 * 
 * If otherparts is non-null, it is an object with:
 *   partname1: "part contents - should end with \n",
 *   partname2: "part contents - should end with \n"
 */
SocialCalc.SpreadsheetControlCreateSpreadsheetSave = function(spreadsheet, otherparts) {
    let otherpartsstr = "";
    let otherpartsnames = "";
    let partname, extranl;

    if (otherparts) {
        for (partname in otherparts) {
            if (otherparts[partname].charAt(otherparts[partname].length - 1) !== "\n") {
                extranl = "\n";
            } else {
                extranl = "";
            }
            otherpartsstr += `--${spreadsheet.multipartBoundary}\nContent-type: text/plain; charset=UTF-8\n\n${otherparts[partname]}${extranl}`;
            otherpartsnames += `part:${partname}\n`;
        }
    }

    const result = `socialcalc:version:1.0\nMIME-Version: 1.0\nContent-Type: multipart/mixed; boundary=${spreadsheet.multipartBoundary}\n--${spreadsheet.multipartBoundary}\nContent-type: text/plain; charset=UTF-8\n\n# SocialCalc Spreadsheet Control Save\nversion:1.0\npart:sheet\npart:edit\npart:audit\n${otherpartsnames}--${spreadsheet.multipartBoundary}\nContent-type: text/plain; charset=UTF-8\n\n${spreadsheet.CreateSheetSave()}--${spreadsheet.multipartBoundary}\nContent-type: text/plain; charset=UTF-8\n\n${spreadsheet.editor.SaveEditorSettings()}--${spreadsheet.multipartBoundary}\nContent-type: text/plain; charset=UTF-8\n\n${spreadsheet.sheet.CreateAuditString()}${otherpartsstr}--${spreadsheet.multipartBoundary}--\n`;

    return result;
};

/**
 * Decode a spreadsheet save string into its component parts
 * 
 * @param {SocialCalc.SpreadsheetControl} spreadsheet - The spreadsheet control object
 * @param {string} str - The save string to decode
 * @returns {Object} Object with part information: {type1: {start: startpos, end: endpos}, type2:...}
 * 
 * Separates the parts from a spreadsheet save string, returning an object with the sub-strings.
 */
SocialCalc.SpreadsheetControlDecodeSpreadsheetSave = function(spreadsheet, str) {
    let pos1, mpregex, searchinfo, boundary, boundaryregex, blanklineregex, start, ending, lines, i, p, pnum;
    const parts = {};
    const partlist = [];

    pos1 = str.search(/^MIME-Version:\s1\.0/mi);
    if (pos1 < 0) return parts;

    mpregex = /^Content-Type:\s*multipart\/mixed;\s*boundary=(\S+)/mig;
    mpregex.lastIndex = pos1;

    searchinfo = mpregex.exec(str);
    if (mpregex.lastIndex <= 0) return parts;
    boundary = searchinfo[1];

    boundaryregex = new RegExp(`^--${boundary}(?:\r\n|\n)`, "mg");
    boundaryregex.lastIndex = mpregex.lastIndex;

    searchinfo = boundaryregex.exec(str); // find header top boundary
    blanklineregex = /(?:\r\n|\n)(?:\r\n|\n)/gm;
    blanklineregex.lastIndex = boundaryregex.lastIndex;
    searchinfo = blanklineregex.exec(str); // skip to after blank line
    if (!searchinfo) return parts;
    start = blanklineregex.lastIndex;
    boundaryregex.lastIndex = start;
    searchinfo = boundaryregex.exec(str); // find end of header
    if (!searchinfo) return parts;
    ending = searchinfo.index;

    lines = str.substring(start, ending).split(/\r\n|\n/); // get header as lines
    for (i = 0; i < lines.length; i++) {
        const line = lines[i];
        p = line.split(":");
        switch (p[0]) {
            case "version":
                break;
            case "part":
                partlist.push(p[1]);
                break;
        }
    }

    for (pnum = 0; pnum < partlist.length; pnum++) { // get each part
        blanklineregex.lastIndex = ending;
        searchinfo = blanklineregex.exec(str); // find blank line ending mime-part header
        if (!searchinfo) return parts;
        start = blanklineregex.lastIndex;
        if (pnum === partlist.length - 1) { // last one has different boundary
            boundaryregex = new RegExp(`^--${boundary}--$`, "mg");
        }
        boundaryregex.lastIndex = start;
        searchinfo = boundaryregex.exec(str); // find ending boundary
        if (!searchinfo) return parts;
        ending = searchinfo.index;
        parts[partlist[pnum]] = { start: start, end: ending }; // return position within full string
    }

    return parts;
};

/**
 * SettingsControls
 * 
 * Each settings panel has an object in the following form:
 * 
 *    {ctrl-name1: {setting: setting-nameA, type: ctrl-type, id: id-component},
 *     ctrl-name2: {setting: setting-nameB, type: ctrl-type, id: id-component, initialdata: optional-initialdata-override},
 *     ...}
 * 
 * The ctrl-types are names that correspond to:
 * 
 *    SocialCalc.SettingsControls.Controls = {
 *       ctrl-type1: {
 *          SetValue: function(panel-obj, ctrl-name, {def: true/false, val: value}) {...;},
 *          ColorValues: if true, Onchanged converts between hex and RGB
 *          GetValue: function(panel-obj, ctrl-name) {...return {def: true/false, val: value};},
 *          Initialize: function(panel-obj, ctrl-name) {...;}, // used to fill dropdowns, etc.
 *          InitialData: control-dependent, // used by Initialize (if no panel ctrlname.initialdata)
 *          OnReset: function(ctrl-name) {...;}, // called to put down popups, etc.
 *          ChangedCallback: function(ctrl-name) {...;} // if not null, called by control when user changes value
 *       }
 * 
 * @type {Object}
 */
SocialCalc.SettingsControls = {
    Controls: {},
    CurrentPanel: null // panel object to search on events
};

/**
 * Set the current settings panel
 * @param {Object} panelobj - Panel object to set as current
 */
SocialCalc.SettingsControlSetCurrentPanel = function(panelobj) {
    SocialCalc.SettingsControls.CurrentPanel = panelobj;
    SocialCalc.SettingsControls.PopupChangeCallback({ panelobj: panelobj }, "", null);
};

/**
 * Initialize all controls in a settings panel
 * @param {Object} panelobj - Panel object containing control definitions
 */
SocialCalc.SettingsControlInitializePanel = function(panelobj) {
    const sc = SocialCalc.SettingsControls;

    for (const ctrlname in panelobj) {
        if (ctrlname === "name") continue;
        const ctrl = sc.Controls[panelobj[ctrlname].type];
        if (ctrl && ctrl.Initialize) {
            ctrl.Initialize(panelobj, ctrlname);
        }
    }
};
/**
 * Load attributes into a settings panel
 * @param {Object} panelobj - Panel object containing control definitions
 * @param {Object} attribs - Attributes object with values to load into the panel
 */
SocialCalc.SettingsControlLoadPanel = function(panelobj, attribs) {
    const sc = SocialCalc.SettingsControls;

    for (const ctrlname in panelobj) {
        if (ctrlname === "name") continue;
        const ctrl = sc.Controls[panelobj[ctrlname].type];
        if (ctrl && ctrl.SetValue) {
            ctrl.SetValue(panelobj, ctrlname, attribs[panelobj[ctrlname].setting]);
        }
    }
};

/**
 * Extract attributes from a settings panel
 * @param {Object} panelobj - Panel object containing control definitions
 * @returns {Object} Attributes object with values from the panel controls
 */
SocialCalc.SettingsControlUnloadPanel = function(panelobj) {
    const sc = SocialCalc.SettingsControls;
    const attribs = {};

    for (const ctrlname in panelobj) {
        if (ctrlname === "name") continue;
        const ctrl = sc.Controls[panelobj[ctrlname].type];
        if (ctrl && ctrl.GetValue) {
            attribs[panelobj[ctrlname].setting] = ctrl.GetValue(panelobj, ctrlname);
        }
    }

    return attribs;
};

/**
 * Callback for popup control changes - updates the sample text display
 * @param {Object} attribs - Attributes object containing panelobj and other data
 * @param {string} id - Control ID that changed
 * @param {*} value - New value
 */
SocialCalc.SettingsControls.PopupChangeCallback = function(attribs, id, value) {
    const sc = SocialCalc.Constants;

    const ele = document.getElementById("sample-text");

    if (!ele || !attribs || !attribs.panelobj) return;

    const idPrefix = SocialCalc.CurrentSpreadsheetControlObject.idPrefix;
    const c = attribs.panelobj.name === "cell" ? "c" : "";

    let v, a, parts, str1, str2;

    parts = sc.defaultCellLayout.match(/^padding.(\S+) (\S+) (\S+) (\S+).vertical.align.(\S+);$/) || [];

    const cv = {
        color: ["textcolor"], 
        backgroundColor: ["bgcolor", "#FFF"],
        fontSize: ["fontsize", sc.defaultCellFontSize], 
        fontFamily: ["fontfamily"],
        paddingTop: ["padtop", parts[1]], 
        paddingRight: ["padright", parts[2]],
        paddingBottom: ["padbottom", parts[3]], 
        paddingLeft: ["padleft", parts[4]],
        verticalAlign: ["alignvert", parts[5]]
    };

    for (a in cv) {
        v = SocialCalc.Popup.GetValue(`${idPrefix}${c}${cv[a][0]}`) || cv[a][1] || "";
        ele.style[a] = v;
    }

    if (c === "c") {
        const borderControls = {
            borderTop: "cbt", 
            borderRight: "cbr", 
            borderBottom: "cbb", 
            borderLeft: "cbl"
        };
        for (a in borderControls) {
            v = SocialCalc.SettingsControls.BorderSideGetValue(attribs.panelobj, borderControls[a]);
            ele.style[a] = v ? (v.val || "") : "";
        }
        v = SocialCalc.Popup.GetValue(`${idPrefix}calignhoriz`);
        ele.style.textAlign = v || "left";
        ele.childNodes[1].style.textAlign = v || "right";
    } else {
        ele.style.border = "";
        v = SocialCalc.Popup.GetValue(`${idPrefix}textalignhoriz`);
        ele.style.textAlign = v || "left";
        v = SocialCalc.Popup.GetValue(`${idPrefix}numberalignhoriz`);
        ele.childNodes[1].style.textAlign = v || "right";
    }

    v = SocialCalc.Popup.GetValue(`${idPrefix}${c}fontlook`);
    parts = v ? (v.match(/^(\S+) (\S+)$/) || []) : [];
    ele.style.fontStyle = parts[1] || "";
    ele.style.fontWeight = parts[2] || "";

    v = SocialCalc.Popup.GetValue(`${idPrefix}${c}formatnumber`) || "General";
    str1 = SocialCalc.FormatNumber.formatNumberWithFormat(9.8765, v, "");
    str2 = SocialCalc.FormatNumber.formatNumberWithFormat(-1234.5, v, "");
    if (str2 !== "??-???-??&nbsp;??:??:??") { // not bad date from negative number
        str1 += `<br>${str2}`;
    }
    
    ele.childNodes[1].innerHTML = str1;
};

/**
 * PopupList Control Implementation
 */

/**
 * Set value for a PopupList control
 * @param {Object} panelobj - Panel object
 * @param {string} ctrlname - Control name
 * @param {Object} value - Value object with def and val properties
 */
SocialCalc.SettingsControls.PopupListSetValue = function(panelobj, ctrlname, value) {
    if (!value) {
        alert(`${ctrlname} no value`);
        return;
    }

    const sp = SocialCalc.Popup;

    if (!value.def) {
        sp.SetValue(panelobj[ctrlname].id, value.val);
    } else {
        sp.SetValue(panelobj[ctrlname].id, "");
    }
};

/**
 * Get value from a PopupList control
 * @param {Object} panelobj - Panel object
 * @param {string} ctrlname - Control name
 * @returns {Object|null} Value object with def and val properties, or null if control not found
 */
SocialCalc.SettingsControls.PopupListGetValue = function(panelobj, ctrlname) {
    const ctl = panelobj[ctrlname];
    if (!ctl) return null;

    const value = SocialCalc.Popup.GetValue(ctl.id);
    if (value) {
        return { def: false, val: value };
    } else {
        return { def: true, val: 0 };
    }
};

/**
 * Initialize a PopupList control
 * @param {Object} panelobj - Panel object
 * @param {string} ctrlname - Control name
 */
SocialCalc.SettingsControls.PopupListInitialize = function(panelobj, ctrlname) {
    const sc = SocialCalc.SettingsControls;
    let initialdata = panelobj[ctrlname].initialdata || 
                      sc.Controls[panelobj[ctrlname].type].InitialData || "";
    initialdata = SocialCalc.LocalizeSubstrings(initialdata);
    const optionvals = initialdata.split(/\|/);

    const options = [];

    for (let i = 0; i < (optionvals.length || 0); i++) {
        const val = optionvals[i];
        const pos = val.indexOf(":");
        let otext = val.substring(0, pos);
        
        if (otext.indexOf("\\") !== -1) { // escape any colons
            otext = otext.replace(/\\c/g, ":");
            otext = otext.replace(/\\b/g, "\\");
        }
        
        otext = SocialCalc.special_chars(otext);
        
        if (otext === "[custom]") {
            options[i] = {
                o: SocialCalc.Constants.s_PopupListCustom, 
                v: val.substring(pos + 1), 
                a: { custom: true }
            };
        } else if (otext === "[cancel]") {
            options[i] = {
                o: SocialCalc.Constants.s_PopupListCancel, 
                v: "", 
                a: { cancel: true }
            };
        } else if (otext === "[break]") {
            options[i] = {
                o: "-----", 
                v: "", 
                a: { skip: true }
            };
        } else if (otext === "[newcol]") {
            options[i] = {
                o: "", 
                v: "", 
                a: { newcol: true }
            };
        } else {
            options[i] = {
                o: otext, 
                v: val.substring(pos + 1)
            };
        }
    }

    SocialCalc.Popup.Create("List", panelobj[ctrlname].id, {});
    SocialCalc.Popup.Initialize(panelobj[ctrlname].id, {
        options: options,
        attribs: {
            changedcallback: SocialCalc.SettingsControls.PopupChangeCallback,
            panelobj: panelobj
        }
    });
};

/**
 * Reset PopupList controls
 * @param {string} ctrlname - Control name (unused but part of interface)
 */
SocialCalc.SettingsControls.PopupListReset = function(ctrlname) {
    SocialCalc.Popup.Reset("List");
};

/**
 * PopupList control definition
 * @type {Object}
 */
SocialCalc.SettingsControls.Controls.PopupList = {
    SetValue: SocialCalc.SettingsControls.PopupListSetValue,
    GetValue: SocialCalc.SettingsControls.PopupListGetValue,
    Initialize: SocialCalc.SettingsControls.PopupListInitialize,
    OnReset: SocialCalc.SettingsControls.PopupListReset,
    ChangedCallback: null
};
/**
 * ColorChooser Control Implementation
 */

/**
 * Set value for a ColorChooser control
 * @param {Object} panelobj - Panel object
 * @param {string} ctrlname - Control name
 * @param {Object} value - Value object with def and val properties
 */
SocialCalc.SettingsControls.ColorChooserSetValue = function(panelobj, ctrlname, value) {
    if (!value) {
        alert(`${ctrlname} no value`);
        return;
    }

    const sp = SocialCalc.Popup;

    if (!value.def) {
        sp.SetValue(panelobj[ctrlname].id, value.val);
    } else {
        sp.SetValue(panelobj[ctrlname].id, "");
    }
};

/**
 * Get value from a ColorChooser control
 * @param {Object} panelobj - Panel object
 * @param {string} ctrlname - Control name
 * @returns {Object} Value object with def and val properties
 */
SocialCalc.SettingsControls.ColorChooserGetValue = function(panelobj, ctrlname) {
    const value = SocialCalc.Popup.GetValue(panelobj[ctrlname].id);
    if (value) {
        return { def: false, val: value };
    } else {
        return { def: true, val: 0 };
    }
};

/**
 * Initialize a ColorChooser control
 * @param {Object} panelobj - Panel object
 * @param {string} ctrlname - Control name
 */
SocialCalc.SettingsControls.ColorChooserInitialize = function(panelobj, ctrlname) {
    SocialCalc.Popup.Create("ColorChooser", panelobj[ctrlname].id, {});
    SocialCalc.Popup.Initialize(panelobj[ctrlname].id, {
        attribs: {
            title: "&nbsp;",
            moveable: true,
            width: "106px",
            changedcallback: SocialCalc.SettingsControls.PopupChangeCallback,
            panelobj: panelobj
        }
    });
};

/**
 * Reset ColorChooser controls
 * @param {string} ctrlname - Control name (unused but part of interface)
 */
SocialCalc.SettingsControls.ColorChooserReset = function(ctrlname) {
    SocialCalc.Popup.Reset("ColorChooser");
};

/**
 * ColorChooser control definition
 * @type {Object}
 */
SocialCalc.SettingsControls.Controls.ColorChooser = {
    SetValue: SocialCalc.SettingsControls.ColorChooserSetValue,
    GetValue: SocialCalc.SettingsControls.ColorChooserGetValue,
    Initialize: SocialCalc.SettingsControls.ColorChooserInitialize,
    OnReset: SocialCalc.SettingsControls.ColorChooserReset,
    ChangedCallback: null
};

/**
 * BorderSide Control Implementation
 */

/**
 * Set value for a BorderSide control
 * @param {Object} panelobj - Panel object
 * @param {string} ctrlname - Control name
 * @param {Object} value - Value object with def and val properties
 */
SocialCalc.SettingsControls.BorderSideSetValue = function(panelobj, ctrlname, value) {
    let ele, idname, parts;
    const idstart = panelobj[ctrlname].id;

    if (!value) {
        alert(`${ctrlname} no value`);
        return;
    }

    ele = document.getElementById(`${idstart}-onoff-bcb`); // border checkbox
    if (!ele) return;

    if (value.val) { // border does not use default: it looks only to the value currently
        ele.checked = true;
        ele.value = value.val;
        parts = value.val.match(/(\S+)\s+(\S+)\s+(\S.+)/);
        idname = `${idstart}-color`;
        SocialCalc.Popup.SetValue(idname, parts[3]);
        SocialCalc.Popup.SetDisabled(idname, false);
    } else {
        ele.checked = false;
        ele.value = value.val;
        idname = `${idstart}-color`;
        SocialCalc.Popup.SetValue(idname, "");
        SocialCalc.Popup.SetDisabled(idname, true);
    }
};

/**
 * Get value from a BorderSide control
 * @param {Object} panelobj - Panel object
 * @param {string} ctrlname - Control name
 * @returns {Object} Value object with def and val properties
 */
SocialCalc.SettingsControls.BorderSideGetValue = function(panelobj, ctrlname) {
    let ele, value;
    const idstart = panelobj[ctrlname].id;

    ele = document.getElementById(`${idstart}-onoff-bcb`); // border checkbox
    if (!ele) return;

    if (ele.checked) { // on
        value = SocialCalc.Popup.GetValue(`${idstart}-color`);
        value = `1px solid ${value || "rgb(0,0,0)"}`;
        return { def: false, val: value };
    } else { // off
        return { def: false, val: "" };
    }
};

/**
 * Initialize a BorderSide control
 * @param {Object} panelobj - Panel object
 * @param {string} ctrlname - Control name
 */
SocialCalc.SettingsControls.BorderSideInitialize = function(panelobj, ctrlname) {
    const idstart = panelobj[ctrlname].id;

    SocialCalc.Popup.Create("ColorChooser", `${idstart}-color`, {});
    SocialCalc.Popup.Initialize(`${idstart}-color`, {
        attribs: {
            title: "&nbsp;",
            width: "106px",
            moveable: true,
            changedcallback: SocialCalc.SettingsControls.PopupChangeCallback,
            panelobj: panelobj
        }
    });
};

/**
 * Handle border control changes
 * @param {Element} ele - The element that triggered the change event
 */
SocialCalc.SettingsControlOnchangeBorder = function(ele) {
    const sc = SocialCalc.SettingsControls;
    const panelobj = sc.CurrentPanel;

    const nameparts = ele.id.match(/(^.*\-)(\w+)\-(\w+)\-(\w+)$/);
    if (!nameparts) return;
    
    const prefix = nameparts[1];
    const ctrlname = nameparts[2];
    const ctrlsubid = nameparts[3];
    const ctrlidsuffix = nameparts[4];
    const ctrltype = panelobj[ctrlname].type;

    switch (ctrlidsuffix) {
        case "bcb": // border checkbox
            if (ele.checked) {
                sc.Controls[ctrltype].SetValue(sc.CurrentPanel, ctrlname, {
                    def: false,
                    val: ele.value || "1px solid rgb(0,0,0)"
                });
            } else {
                sc.Controls[ctrltype].SetValue(sc.CurrentPanel, ctrlname, {
                    def: false,
                    val: ""
                });
            }
            break;
    }
};

/**
 * BorderSide control definition
 * @type {Object}
 */
SocialCalc.SettingsControls.Controls.BorderSide = {
    SetValue: SocialCalc.SettingsControls.BorderSideSetValue,
    GetValue: SocialCalc.SettingsControls.BorderSideGetValue,
    OnClick: SocialCalc.SettingsControls.ColorComboOnClick,
    Initialize: SocialCalc.SettingsControls.BorderSideInitialize,
    InitialData: { thickness: "1 pixel:1px", style: "Solid:solid" },
    ChangedCallback: null
};

/**
 * Reset all settings controls
 * Calls OnReset for all registered control types
 */
SocialCalc.SettingControlReset = function() {
    const sc = SocialCalc.SettingsControls;

    for (const ctrlname in sc.Controls) {
        if (sc.Controls[ctrlname].OnReset) {
            sc.Controls[ctrlname].OnReset(ctrlname);
        }
    }
};

/**
 * CtrlSEditor implementation for editing SocialCalc.OtherSaveParts
 */

/**
 * Holds other parts to save - must be set when loaded if you want to keep
 * @type {Object}
 */
SocialCalc.OtherSaveParts = {};

/**
 * Create and display an editor for SocialCalc save parts
 * @param {string} whichpart - The part name to edit, or empty string to list all parts
 */
SocialCalc.CtrlSEditor = function(whichpart) {
    let strtoedit, partname;
    
    if (whichpart.length > 0) {
        strtoedit = SocialCalc.special_chars(SocialCalc.OtherSaveParts[whichpart] || "");
    } else {
        strtoedit = "Listing of Parts\n";
        for (partname in SocialCalc.OtherSaveParts) {
            strtoedit += SocialCalc.special_chars(`\nPart: ${partname}\n=====\n${SocialCalc.OtherSaveParts[partname]}\n`);
        }
    }
    
    const editbox = document.createElement("div");
    editbox.style.cssText = "position:absolute;z-index:500;width:300px;height:300px;left:100px;top:200px;border:1px solid black;background-color:#EEE;text-align:center;";
    editbox.id = "socialcalc-editbox";
    editbox.innerHTML = `${whichpart}<br><br><textarea id="socialcalc-editbox-textarea" style="width:250px;height:200px;">${strtoedit}</textarea><br><br><input type="button" onclick="SocialCalc.CtrlSEditorDone('socialcalc-editbox', '${whichpart}');" value="OK">`;
    document.body.appendChild(editbox);

    const ebta = document.getElementById("socialcalc-editbox-textarea");
    ebta.focus();
    SocialCalc.CmdGotFocus(ebta);
};

/**
 * Close the CtrlSEditor and save changes
 * @param {string} idprefix - The ID prefix of the editor elements
 * @param {string} whichpart - The part name being edited
 */
SocialCalc.CtrlSEditorDone = function(idprefix, whichpart) {
    const edittextarea = document.getElementById(`${idprefix}-textarea`);
    const text = edittextarea.value;
    
    if (whichpart.length > 0) {
        if (text.length > 0) {
            SocialCalc.OtherSaveParts[whichpart] = text;
        } else {
            delete SocialCalc.OtherSaveParts[whichpart];
        }
    }

    const editbox = document.getElementById(idprefix);
    SocialCalc.KeyboardFocus();
    editbox.parentNode.removeChild(editbox);
};



