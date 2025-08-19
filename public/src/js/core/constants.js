/**
 * @fileoverview The module of the SocialCalc package with customizable constants, strings, etc.
 * This is where most of the common localizations are done.
 * 
 * @copyright (c) Copyright 2008, 2009, 2010 Socialtext, Inc. All Rights Reserved.
 * 
 * @license The contents of this file are subject to the Artistic License 2.0; you may not
 * use this file except in compliance with the License. You may obtain a copy of 
 * the License at http://socialcalc.org/licenses/al-20/.
 * 
 * Some of the other files in the SocialCalc package are licensed under
 * different licenses. Please note the licenses of the modules you use.
 * 
 * @history
 * Initially coded by Dan Bricklin of Software Garden, Inc., for Socialtext, Inc.
 * Based in part on the SocialCalc 1.1.0 code written in Perl.
 * The SocialCalc 1.1.0 code was:
 *    Portions (c) Copyright 2005, 2006, 2007 Software Garden, Inc.
 *    All Rights Reserved.
 *    Portions (c) Copyright 2007 Socialtext, Inc.
 *    All Rights Reserved.
 * The Perl SocialCalc started as modifications to the wikiCalc(R) program, version 1.0.
 * wikiCalc 1.0 was written by Software Garden, Inc.
 * Unless otherwise specified, referring to "SocialCalc" in comments refers to this
 * JavaScript version of the code, not the SocialCalc Perl code.
 */

/**
 * Initialize SocialCalc namespace if not exists
 * @namespace SocialCalc
 */
var SocialCalc;
if (!SocialCalc) SocialCalc = {};

/**
 * @description TO LEARN HOW TO LOCALIZE OR CUSTOMIZE SOCIALCALC, PLEASE READ THIS:
 * 
 * The constants are all properties of the SocialCalc.Constants object.
 * They are grouped here by what they are for, which module uses them, etc.
 * 
 * Properties whose names start with "s_" are strings, or arrays of strings,
 * that are good candidates for translation from the English.
 * 
 * Other properties relate to visual settings, localization parameters, etc.
 * 
 * These values are not used when SocialCalc modules are first loaded.
 * They may be modified before the first use of the routines that use them,
 * e.g., before creating SocialCalc objects.
 * 
 * The exceptions are:
 *    TooltipOffsetX and TooltipOffsetY, as described with their definitions.
 * 
 * SocialCalc IS NOT DESIGNED FOR USE WITH A TRANSLATION FUNCTION each time a string
 * is used. Instead, language translations may be done by modifying this object.
 * 
 * To customize SocialCalc, you may either replace this file with a modified version
 * or you can overwrite the values before use. An example would be to
 * iterate over all the properties looking for names that start with "s_" and
 * use some other mechanism to obtain a localized string and replace the values
 * here with those translated values.
 * 
 * There is also a function, SocialCalc.ConstantsSetClasses, that may be used
 * to easily switch SocialCalc from using explicit CSS styles for many things
 * to using CSS classes. See the function, below, for more information.
 */

/**
 * SocialCalc Constants Configuration
 * Contains all customizable constants, strings, and localization settings
 * @namespace SocialCalc.Constants
 */
SocialCalc.Constants = {

   /**
    * Main SocialCalc module, socialcalc-3.js
    */

   /**
    * Common Constants
    */

   /** 
    * Sets the default type for text on reading source file
    * @type {string}
    * @default "t"
    */
   textdatadefaulttype: "t",

   /**
    * Common error messages
    */

   /** 
    * Error thrown if browser can't handle events like IE or Firefox
    * @type {string}
    */
   s_BrowserNotSupported: "Browser not supported.",

   /** 
    * Internal error message - hopefully unlikely, but a test failed
    * @type {string}
    */
   s_InternalError: "Internal SocialCalc error (probably an internal bug): ",

   /**
    * SocialCalc.ParseSheetSave - Errors thrown on unexpected value in save file
    */

   /** 
    * Unknown column type item error
    * @type {string}
    */
   s_pssUnknownColType: "Unknown col type item",

   /** 
    * Unknown row type item error
    * @type {string}
    */
   s_pssUnknownRowType: "Unknown row type item",

   /** 
    * Unknown line type error
    * @type {string}
    */
   s_pssUnknownLineType: "Unknown line type",

   /**
    * SocialCalc.CellFromStringParts - Error thrown on unexpected value in save file
    */

   /** 
    * Unknown cell type item error
    * @type {string}
    */
   s_cfspUnknownCellType: "Unknown cell type item",

   /**
    * SocialCalc.CanonicalizeSheet
    */

   /** 
    * If true, do the canonicalization calculations
    * @type {boolean}
    */
   doCanonicalizeSheet: true,

   /**
    * ExecuteSheetCommand
    */

   /** 
    * Unknown sheet command error
    * @type {string}
    */
   s_escUnknownSheetCmd: "Unknown sheet command: ",

   /** 
    * Unknown set coord command error
    * @type {string}
    */
   s_escUnknownSetCoordCmd: "Unknown set coord command: ",

   /** 
    * Unknown command error
    * @type {string}
    */
   s_escUnknownCmd: "Unknown command: ",

   /**
    * SocialCalc.CheckAndCalcCell
    */

   /** 
    * Circular reference found during recalc
    * @type {string}
    */
   s_caccCircRef: "Circular reference to ",

   /**
    * SocialCalc.RenderContext
    */

   /** 
    * Used to set minimum width of the row header column - a string in pixels
    * @type {string}
    */
   defaultRowNameWidth: "30",

   /** 
    * Used when guessing row heights - number
    * @type {number}
    */
   defaultAssumedRowHeight: 15,

   /** 
    * If non-null, each cell will render with an ID starting with this
    * @type {string}
    */
   defaultCellIDPrefix: "cell_",

   /**
    * Default sheet display values
    */

   /** 
    * Default cell layout CSS
    * @type {string}
    */
   defaultCellLayout: "padding:2px 2px 1px 2px;vertical-align:top;",

   /** 
    * Default cell font style
    * @type {string}
    */
   defaultCellFontStyle: "normal normal",

   /** 
    * Default cell font size
    * @type {string}
    */
   defaultCellFontSize: "small",

   /** 
    * Default cell font family
    * @type {string}
    */
   defaultCellFontFamily: "Verdana,Arial,Helvetica,sans-serif",

   /** 
    * Default pane divider width as string
    * @type {string}
    */
   defaultPaneDividerWidth: "2",

   /** 
    * Default pane divider height as string
    * @type {string}
    */
   defaultPaneDividerHeight: "3",

   /** 
    * Used as style to set each border when grid enabled
    * @type {string}
    */
   defaultGridCSS: "1px solid #C0C0C0;",

   /** 
    * Class added to cells with non-null comments when grid enabled
    * @type {string}
    */
   defaultCommentClass: "",

   /** 
    * Style added to cells with non-null comments when grid enabled
    * @type {string}
    */
   defaultCommentStyle: "background-repeat:no-repeat;background-position:top right;background-image:url(/src/images/sc-commentbg.gif);",

   /** 
    * Class added to cells with non-null comments when grid not enabled
    * @type {string}
    */
   defaultCommentNoGridClass: "",

   /** 
    * Style added to cells with non-null comments when grid not enabled
    * @type {string}
    */
   defaultCommentNoGridStyle: "",

   /** 
    * Default column width as text
    * @type {string}
    */
   defaultColWidth: "80",

   /** 
    * Minimum column width as numeric value
    * @type {number}
    */
   defaultMinimumColWidth: 10,

   /**
    * Default sheet display values - at least one of class and/or style are needed for each
    */

   /** 
    * Default highlight type cursor class
    * @type {string}
    */
   defaultHighlightTypeCursorClass: "",

   /** 
    * Default highlight type cursor style
    * @type {string}
    */
   defaultHighlightTypeCursorStyle: "color:#FFF;backgroundColor:#A6A6A6;",

   /** 
    * Default highlight type range class
    * @type {string}
    */
   defaultHighlightTypeRangeClass: "",

   /** 
    * Default highlight type range style
    * @type {string}
    */
   defaultHighlightTypeRangeStyle: "color:#000;backgroundColor:#E5E5E5;",

   /** 
    * Regular column heading letters class - needs a cursor property
    * @type {string}
    */
   defaultColnameClass: "",

   /** 
    * Regular column heading letters style - needs a cursor property
    * @type {string}
    */
   defaultColnameStyle: "font-size:small;text-align:center;color:#FFFFFF;background-color:#808080;cursor:e-resize;",

   /** 
    * Column with selected cell class - needs a cursor property
    * @type {string}
    */
   defaultSelectedColnameClass: "",

   /** 
    * Column with selected cell style - needs a cursor property
    * @type {string}
    */
   defaultSelectedColnameStyle: "font-size:small;text-align:center;color:#FFFFFF;background-color:#404040;cursor:e-resize;",

   /** 
    * Regular row heading numbers class
    * @type {string}
    */
   defaultRownameClass: "",

   /** 
    * Regular row heading numbers style
    * @type {string}
    */
   defaultRownameStyle: "font-size:small;text-align:right;color:#FFFFFF;background-color:#808080;",

   /** 
    * Row with selected cell class - needs a cursor property
    * @type {string}
    */
   defaultSelectedRownameClass: "",

   /** 
    * Row with selected cell style - needs a cursor property
    * @type {string}
    */
   defaultSelectedRownameStyle: "font-size:small;text-align:right;color:#FFFFFF;background-color:#404040;",

   /** 
    * Corner cell in upper left class
    * @type {string}
    */
   defaultUpperLeftClass: "",

   /** 
    * Corner cell in upper left style
    * @type {string}
    */
   defaultUpperLeftStyle: "font-size:small;",

   /** 
    * Used if present for spanned cells peeking into a pane - at least one of class/style needed
    * @type {string}
    */
   defaultSkippedCellClass: "",

   /** 
    * Used if present for spanned cells peeking into a pane
    * @type {string}
    */
   defaultSkippedCellStyle: "font-size:small;background-color:#CCC",

   /** 
    * Used if present for the look of the space between panes - at least one of class/style needed
    * @type {string}
    */
   defaultPaneDividerClass: "",

   /** 
    * Used if present for the look of the space between panes
    * @type {string}
    */
   defaultPaneDividerStyle: "font-size:small;background-color:#C0C0C0;padding:0px;",

   /** 
    * Unlikely thrown error
    * @type {string}
    */
   s_rcMissingSheet: "Render Context must have a sheet object",

   /**
    * SocialCalc.format_text_for_display
    */

   /** 
    * Used for format "text-link"; you could make this an img tag if desired
    * @type {string}
    */
   defaultLinkFormatString: '<span style="font-size:smaller;text-decoration:none !important;background-color:#66B;color:#FFF;">Link</span>',

   /** 
    * Used for format "text-link"; you could make this an img tag if desired
    * @type {string}
    */
   defaultPageLinkFormatString: '<span style="font-size:smaller;text-decoration:none !important;background-color:#66B;color:#FFF;">Page</span>',

   /**
    * SocialCalc.format_number_for_display
    */

   /** 
    * Default date-time format
    * @type {string}
    */
   defaultFormatdt: 'd-mmm-yyyy h:mm:ss',

   /** 
    * Default date format
    * @type {string}
    */
   defaultFormatd: 'd-mmm-yyyy',

   /** 
    * Default time format
    * @type {string}
    */
   defaultFormatt: '[h]:mm:ss',

   /** 
    * How TRUE shows when rendered
    * @type {string}
    */
   defaultDisplayTRUE: 'TRUE',

   /** 
    * How FALSE shows when rendered
    * @type {string}
    */
   defaultDisplayFALSE: 'FALSE',

   /**
    * SocialCalc Table Editor module, socialcalctableeditor.js
    */

   /**
    * SocialCalc.TableEditor
    */

   /** 
    * URL prefix for images relative to HTML files
    * @type {string}
    */
   defaultImagePrefix: "/src/images/sc-",

   /** 
    * If present, many TableEditor elements are assigned IDs with this prefix
    * @type {string}
    */
   defaultTableEditorIDPrefix: "te_",

   /** 
    * Number of rows to move cursor on PgUp/PgDn keys
    * @type {number}
    */
   defaultPageUpDnAmount: 15,

   /** 
    * Turns on Ctrl-S trapdoor for setting custom numeric formats and commands if true
    * @type {boolean}
    */
   AllowCtrlS: true,

   /**
    * SocialCalc.CreateTableEditor
    */

   /** 
    * The short size for the scrollbars, etc. (numeric in pixels)
    * @type {number}
    */
   defaultTableControlThickness: 20,

   /** 
    * If present, the class for the TableEditor griddiv element
    * @type {string}
    */
   cteGriddivClass: "",

   /**
    * SocialCalc.EditorGetStatuslineString - strings shown on status line
    */

   /** @type {string} */
   s_statusline_executing: "Executing...",

   /** @type {string} */
   s_statusline_displaying: "Displaying...",

   /** @type {string} */
   s_statusline_ordering: "Ordering...",

   /** @type {string} */
   s_statusline_calculating: "Calculating...",

   /** @type {string} */
   s_statusline_calculatingls: "Calculating... Loading Sheet...",

   /** @type {string} */
   s_statusline_doingserverfunc: "doing server function ",

   /** @type {string} */
   s_statusline_incell: " in cell ",

   /** @type {string} */
   s_statusline_calcstart: "Calculation start...",

   /** @type {string} */
   s_statusline_sum: "SUM",

   /** @type {string} */
   s_statusline_recalcneeded: '<span style="color:#999;">(Recalc needed)</span>',

   /** @type {string} */
   s_statusline_circref: '<span style="color:red;">Circular reference: ',

   /**
    * SocialCalc.InputBoxDisplayCellContents
    */

   /** @type {string} */
   s_inputboxdisplaymultilinetext: "[Multi-line text: Click icon on right to edit]",

   /**
    * SocialCalc.InputEcho
    */

   /** 
    * If present, the class of the popup inputEcho div
    * @type {string}
    */
   defaultInputEchoClass: "",

   /** 
    * If present, pseudo style for inputEcho
    * @type {string}
    */
   defaultInputEchoStyle: "filter:alpha(opacity=90);opacity:.9;backgroundColor:#FFD;border:1px solid #884;" +
      "fontSize:small;padding:2px 10px 1px 2px;cursor:default;",

   /** 
    * If present, the class of the popup inputEcho div
    * @type {string}
    */
   defaultInputEchoPromptClass: "",

   /** 
    * If present, pseudo style for inputEcho prompt
    * @type {string}
    */
   defaultInputEchoPromptStyle: "filter:alpha(opacity=90);opacity:.9;backgroundColor:#FFD;" +
      "borderLeft:1px solid #884;borderRight:1px solid #884;borderBottom:1px solid #884;" +
      "fontSize:small;fontStyle:italic;padding:2px 10px 1px 2px;cursor:default;",

   /**
    * SocialCalc.InputEchoText
    */

   /** 
    * Displayed when typing "=unknown("
    * @type {string}
    */
   ietUnknownFunction: "Unknown function ",

   /**
    * SocialCalc.CellHandles
    */

   /** 
    * Extent of inner circle within 90px image
    * @type {number}
    */
   CH_radius1: 29.0,

   /** 
    * Extent of outer circle within 90px image
    * @type {number}
    */
   CH_radius2: 41.0,

   /** 
    * Tooltip for fill all handle
    * @type {string}
    */
   s_CHfillAllTooltip: "Fill Contents and Formats Down/Right",

   /** 
    * Tooltip for fill formulas handle
    * @type {string}
    */
   s_CHfillContentsTooltip: "Fill Contents Only Down/Right",

   /** 
    * Tooltip for move paste all
    * @type {string}
    */
   s_CHmovePasteAllTooltip: "Move Contents and Formats",

   /** 
    * Tooltip for move paste contents
    * @type {string}
    */
   s_CHmovePasteContentsTooltip: "Move Contents Only",

   /** 
    * Tooltip for move insert all
    * @type {string}
    */
   s_CHmoveInsertAllTooltip: "Slide Contents and Formats within Row/Col",

   /** 
    * Tooltip for move insert contents
    * @type {string}
    */
   s_CHmoveInsertContentsTooltip: "Slide Contents within Row/Col",

   /** 
    * Short form of operation to follow drag
    * @type {Object<string, string>}
    */
   s_CHindicatorOperationLookup: {
      "Fill": "Fill", 
      "FillC": "Fill Contents",
      "Move": "Move", 
      "MoveI": "Slide", 
      "MoveC": "Move Contents", 
      "MoveIC": "Slide Contents"
   },

   /** 
    * Direction that modifies operation during drag
    * @type {Object<string, string>}
    */
   s_CHindicatorDirectionLookup: {
      "Down": " Down", 
      "Right": " Right",
      "Horizontal": " Horizontal", 
      "Vertical": " Vertical"
   },

/**
 * SocialCalc.TableControl
 */

/** 
 * Length of pane slider (numeric in pixels)
 * @type {number}
 */
defaultTCSliderThickness: 9,

/** 
 * Length of scroll +/- buttons (numeric in pixels)
 * @type {number}
 */
defaultTCButtonThickness: 20,

/** 
 * Length of thumb (numeric in pixels)
 * @type {number}
 */
defaultTCThumbThickness: 15,

/**
 * SocialCalc.CreateTableControl
 */

/** 
 * If present, pseudo style (text-align is textAlign) for main div of a table control
 * @type {string}
 */
TCmainStyle: "backgroundColor:#EEE;",

/** 
 * If present, the CSS class of the main div for a table control
 * @type {string}
 */
TCmainClass: "",

/** 
 * backgroundColor may be used while waiting for image that may not come
 * @type {string}
 */
TCendcapStyle: "backgroundColor:#FFF;",

/** 
 * CSS class for table control endcap
 * @type {string}
 */
TCendcapClass: "",

/** 
 * Style for table control pane slider
 * @type {string}
 */
TCpanesliderStyle: "backgroundColor:#CCC;",

/** 
 * CSS class for table control pane slider
 * @type {string}
 */
TCpanesliderClass: "",

/** 
 * Tooltip for horizontal table control pane slider
 * @type {string}
 */
s_panesliderTooltiph: "Drag to lock pane vertically",

/** 
 * Tooltip for vertical table control pane slider
 * @type {string}
 */
s_panesliderTooltipv: "Drag to lock pane horizontally",

/** 
 * Style for table control less button
 * @type {string}
 */
TClessbuttonStyle: "backgroundColor:#AAA;",

/** 
 * CSS class for table control less button
 * @type {string}
 */
TClessbuttonClass: "",

/** 
 * Repeat wait time for less button in milliseconds
 * @type {number}
 */
TClessbuttonRepeatWait: 300,

/** 
 * Repeat interval for less button in milliseconds
 * @type {number}
 */
TClessbuttonRepeatInterval: 20,

/** 
 * Style for table control more button
 * @type {string}
 */
TCmorebuttonStyle: "backgroundColor:#AAA;",

/** 
 * CSS class for table control more button
 * @type {string}
 */
TCmorebuttonClass: "",

/** 
 * Repeat wait time for more button in milliseconds
 * @type {number}
 */
TCmorebuttonRepeatWait: 300,

/** 
 * Repeat interval for more button in milliseconds
 * @type {number}
 */
TCmorebuttonRepeatInterval: 20,

/** 
 * Style for table control scroll area
 * @type {string}
 */
TCscrollareaStyle: "backgroundColor:#DDD;",

/** 
 * CSS class for table control scroll area
 * @type {string}
 */
TCscrollareaClass: "",

/** 
 * Repeat wait time for scroll area in milliseconds
 * @type {number}
 */
TCscrollareaRepeatWait: 500,

/** 
 * Repeat interval for scroll area in milliseconds
 * @type {number}
 */
TCscrollareaRepeatInterval: 100,

/** 
 * CSS class for table control thumb
 * @type {string}
 */
TCthumbClass: "",

/** 
 * Style for table control thumb
 * @type {string}
 */
TCthumbStyle: "backgroundColor:#CCC;",

/**
 * SocialCalc.TCPSDragFunctionStart
 */

/** 
 * At least one of class/style for pane slider tracking line display in table control
 * @type {string}
 */
TCPStrackinglineClass: "",

/** 
 * If present, pseudo style (text-align is textAlign) for tracking line
 * @type {string}
 */
TCPStrackinglineStyle: "overflow:hidden;position:absolute;zIndex:100;",

/** 
 * Narrow dimension of tracking line (string with units)
 * @type {string}
 */
TCPStrackinglineThickness: "2px",

/**
 * SocialCalc.TCTDragFunctionStart
 */

/** 
 * At least one of class/style for vertical thumb dragging status display in table control
 * @type {string}
 */
TCTDFSthumbstatusvClass: "",

/** 
 * If present, pseudo style (text-align is textAlign) for vertical thumb status
 * @type {string}
 */
TCTDFSthumbstatusvStyle: "height:20px;width:auto;border:3px solid #808080;overflow:hidden;" +
                         "backgroundColor:#FFF;fontSize:small;position:absolute;zIndex:100;",

/** 
 * At least one of class/style for horizontal thumb dragging status display in table control
 * @type {string}
 */
TCTDFSthumbstatushClass: "",

/** 
 * If present, pseudo style (text-align is textAlign) for horizontal thumb status
 * @type {string}
 */
TCTDFSthumbstatushStyle: "height:20px;width:auto;border:1px solid black;padding:2px;" +
                         "backgroundColor:#FFF;fontSize:small;position:absolute;zIndex:100;",

/** 
 * At least one of class/style for thumb dragging status display in table control
 * @type {string}
 */
TCTDFSthumbstatusrownumClass: "",

/** 
 * If present, real style for thumb status row number
 * @type {string}
 */
TCTDFSthumbstatusrownumStyle: "color:#FFF;background-color:#808080;font-size:small;white-space:nowrap;padding:3px;",

/** 
 * Top offset for thumbstatus display while dragging vertically
 * @type {number}
 */
TCTDFStopOffsetv: 0,

/** 
 * Left offset for thumbstatus display while dragging vertically
 * @type {number}
 */
TCTDFSleftOffsetv: -80,

/** 
 * Text Control Drag Function text before row number
 * @type {string}
 */
s_TCTDFthumbstatusPrefixv: "Row ",

/** 
 * Top offset for thumbstatus display while dragging horizontally
 * @type {number}
 */
TCTDFStopOffseth: -30,

/** 
 * Left offset for thumbstatus display while dragging horizontally
 * @type {number}
 */
TCTDFSleftOffseth: 0,

/** 
 * Text Control Drag Function text before col number
 * @type {string}
 */
s_TCTDFthumbstatusPrefixh: "Col ",

/**
 * SocialCalc.TooltipInfo
 * 
 * Note: These two values are used to set the TooltipInfo initial values when the code is first read in.
 * Modifying them here after loading has no effect -- you need to modify SocialCalc.TooltipInfo directly
 * to dynamically set them. This is different than most other constants which may be modified until use.
 */

/** 
 * Offset in pixels from mouse position (to right on left side of screen, to left on right)
 * @type {number}
 */
TooltipOffsetX: 2,

/** 
 * Offset in pixels above mouse position for lower edge
 * @type {number}
 */
TooltipOffsetY: 10,

/**
 * SocialCalc.TooltipDisplay
 */

/** 
 * At least one of class/style for tooltip display
 * @type {string}
 */
TDpopupElementClass: "",

/** 
 * If present, pseudo style (text-align is textAlign) for tooltip popup
 * @type {string}
 */
TDpopupElementStyle: "border:1px solid black;padding:1px 2px 2px 2px;textAlign:center;backgroundColor:#FFF;" +
                     "fontSize:7pt;fontFamily:Verdana,Arial,Helvetica,sans-serif;" +
                     "position:absolute;width:auto;zIndex:110;",

/**
 * SocialCalc Spreadsheet Control module, socialcalcspreadsheetcontrol.js
 */

/**
 * SocialCalc.SpreadsheetControl
 */

/** 
 * Background style for spreadsheet control toolbar
 * @type {string}
 */
SCToolbarbackground: "background-color:#404040;",

/** 
 * Background style for spreadsheet control tabs
 * @type {string}
 */
SCTabbackground: "background-color:#CCC;",

/** 
 * CSS for selected spreadsheet control tab
 * @type {string}
 */
SCTabselectedCSS: "font-size:small;padding:6px 30px 6px 8px;color:#FFF;background-color:#404040;cursor:default;border-right:1px solid #CCC;",

/** 
 * CSS for plain spreadsheet control tab
 * @type {string}
 */
SCTabplainCSS: "font-size:small;padding:6px 30px 6px 8px;color:#FFF;background-color:#808080;cursor:default;border-right:1px solid #CCC;",

/** 
 * Text style for spreadsheet control toolbar
 * @type {string}
 */
SCToolbartext: "font-size:x-small;font-weight:bold;color:#FFF;padding-bottom:4px;",

/** 
 * Height of formula bar in pixels, will contain a text input box
 * @type {number}
 */
SCFormulabarheight: 30,

/** 
 * Height of status line in pixels
 * @type {number}
 */
SCStatuslineheight: 20,

/** 
 * CSS for status line
 * @type {string}
 */
SCStatuslineCSS: "font-size:10px;padding:3px 0px;",

/**
 * Workbook settings
 */

/** 
 * Enable workbook functionality
 * @type {boolean}
 */
doWorkBook: true,

/** 
 * Height of sheet bar in pixels
 * @type {number}
 */
SCSheetBarHeight: 25,

/** 
 * Background style for sheet bar
 * @type {string}
 */
SCSheetBarBackground: "background-color:#CCC;",

/** 
 * CSS for sheet bar
 * @type {string}
 */
SCSheetBarCSS: "background-color:#CCC;",

/** 
 * Width of sheet bar as percentage
 * @type {string}
 */
SCSheetBarWidth: "70%",

/**
 * Constants for default Format tab (settings)
 * 
 * *** EVEN THOUGH THESE DON'T START WITH s_: ***
 * 
 * These should be carefully checked for localization. Make sure you understand what they do and how they work!
 * The first part of "first:second|first:second|..." is what is displayed and the second is the value to be used.
 * The value is normally not translated -- only the displayed part. The [cancel], [break], etc., are not translated --
 * they are commands to SocialCalc.SettingsControls.PopupListInitialize
 */

/** 
 * Format options for number formats
 * @type {string}
 */
SCFormatNumberFormats: "[cancel]:|[break]:|%loc!Default!:|[custom]:|%loc!Automatic!:general|%loc!Auto w/ commas!:[,]General|[break]:|" +
        "00:00|000:000|0000:0000|00000:00000|[break]:|%loc!Formula!:formula|%loc!Hidden!:hidden|[newcol]:" +
        "1234:0|1,234:#,##0|1,234.5:#,##0.0|1,234.56:#,##0.00|1,234.567:#,##0.000|1,234.5678:#,##0.0000|" +
        "[break]:|1,234%:#,##0%|1,234.5%:#,##0.0%|1,234.56%:#,##0.00%|" +
        "[newcol]:|$1,234:$#,##0|$1,234.5:$#,##0.0|$1,234.56:$#,##0.00|[break]:|" +
        "(1,234):#,##0_);(#,##0)|(1,234.5):#,##0.0_);(#,##0.0)|(1,234.56):#,##0.00_);(#,##0.00)|[break]:|" +
        "($1,234):#,##0_);($#,##0)|($1,234.5):$#,##0.0_);($#,##0.0)|($1,234.56):$#,##0.00_);($#,##0.00)|" +
        "[newcol]:|1/4/06:m/d/yy|01/04/2006:mm/dd/yyyy|2006-01-04:yyyy-mm-dd|4-Jan-06:d-mmm-yy|04-Jan-2006:dd-mmm-yyyy|January 4, 2006:mmmm d, yyyy|" +
        "[break]:|1\\c23:h:mm|1\\c23 PM:h:mm AM/PM|1\\c23\\c45:h:mm:ss|01\\c23\\c45:hh:mm:ss|26\\c23 (h\\cm):[hh]:mm|69\\c45 (m\\cs):[mm]:ss|69 (s):[ss]|" +
        "[newcol]:|2006-01-04 01\\c23\\c45:yyyy-mm-dd hh:mm:ss|January 4, 2006:mmmm d, yyyy hh:mm:ss|Wed:ddd|Wednesday:dddd|",

/** 
 * Format options for text formats
 * @type {string}
 */
SCFormatTextFormats: "[cancel]:|[break]:|%loc!Default!:|[custom]:|%loc!Automatic!:general|%loc!Plain Text!:text-plain|" +
        "HTML:text-html|%loc!Wikitext!:text-wiki|%loc!Link!:text-link|%loc!Formula!:formula|%loc!Hidden!:hidden|",

/** 
 * Format options for padding sizes
 * @type {string}
 */
SCFormatPadsizes: "[cancel]:|[break]:|%loc!Default!:|[custom]:|%loc!No padding!:0px|" +
        "[newcol]:|1 pixel:1px|2 pixels:2px|3 pixels:3px|4 pixels:4px|5 pixels:5px|" +
        "6 pixels:6px|7 pixels:7px|8 pixels:8px|[newcol]:|9 pixels:9px|10 pixels:10px|11 pixels:11px|" +
        "12 pixels:12px|13 pixels:13px|14 pixels:14px|16 pixels:16px|" +
        "18 pixels:18px|[newcol]:|20 pixels:20px|22 pixels:22px|24 pixels:24px|28 pixels:28px|36 pixels:36px|",

/** 
 * Format options for font sizes
 * @type {string}
 */
SCFormatFontsizes: "[cancel]:|[break]:|%loc!Default!:|[custom]:|X-Small:x-small|Small:small|Medium:medium|Large:large|X-Large:x-large|" +
              "[newcol]:|6pt:6pt|7pt:7pt|8pt:8pt|9pt:9pt|10pt:10pt|11pt:11pt|12pt:12pt|14pt:14pt|16pt:16pt|" +
              "[newcol]:|18pt:18pt|20pt:20pt|22pt:22pt|24pt:24pt|28pt:28pt|36pt:36pt|48pt:48pt|72pt:72pt|" +
              "[newcol]:|8 pixels:8px|9 pixels:9px|10 pixels:10px|11 pixels:11px|" +
              "12 pixels:12px|13 pixels:13px|14 pixels:14px|[newcol]:|16 pixels:16px|" +
              "18 pixels:18px|20 pixels:20px|22 pixels:22px|24 pixels:24px|28 pixels:28px|36 pixels:36px|",

/** 
 * Format options for font families
 * @type {string}
 */
SCFormatFontfamilies: "[cancel]:|[break]:|%loc!Default!:|[custom]:|Verdana:Verdana,Arial,Helvetica,sans-serif|" +
              "Arial:arial,helvetica,sans-serif|Courier:'Courier New',Courier,monospace|",

/** 
 * Format options for font appearance
 * @type {string}
 */
SCFormatFontlook: "[cancel]:|[break]:|%loc!Default!:|%loc!Normal!:normal normal|%loc!Bold!:normal bold|%loc!Italic!:italic normal|" +
              "%loc!Bold Italic!:italic bold",

/** 
 * Format options for text horizontal alignment
 * @type {string}
 */
SCFormatTextAlignhoriz: "[cancel]:|[break]:|%loc!Default!:|%loc!Left!:left|%loc!Center!:center|%loc!Right!:right|",

/** 
 * Format options for number horizontal alignment
 * @type {string}
 */
SCFormatNumberAlignhoriz: "[cancel]:|[break]:|%loc!Default!:|%loc!Left!:left|%loc!Center!:center|%loc!Right!:right|",

/** 
 * Format options for vertical alignment
 * @type {string}
 */
SCFormatAlignVertical: "[cancel]:|[break]:|%loc!Default!:|%loc!Top!:top|%loc!Middle!:middle|%loc!Bottom!:bottom|",

/** 
 * Format options for column width
 * @type {string}
 */
SCFormatColwidth: "[cancel]:|[break]:|%loc!Default!:|[custom]:|[newcol]:|" +
              "20 pixels:20|40:40|60:60|80:80|100:100|120:120|140:140|160:160|" +
              "[newcol]:|180 pixels:180|200:200|220:220|240:240|260:260|280:280|300:300|",

/** 
 * Format options for recalculation settings
 * @type {string}
 */
SCFormatRecalc: "[cancel]:|[break]:|%loc!Auto!:|%loc!Manual!:off|",

/**
 * SocialCalc.InitializeSpreadsheetControl
 */

/** 
 * Normal border color for ISC buttons
 * @type {string}
 */
ISCButtonBorderNormal: "#404040",

/** 
 * Hover border color for ISC buttons
 * @type {string}
 */
ISCButtonBorderHover: "#999",

/** 
 * Down border color for ISC buttons
 * @type {string}
 */
ISCButtonBorderDown: "#FFF",

/** 
 * Background color for ISC buttons when pressed down
 * @type {string}
 */
ISCButtonDownBackground: "#888",

/**
 * SocialCalc.SettingsControls.PopupListInitialize
 */

/** 
 * Cancel text for popup lists
 * @type {string}
 */
s_PopupListCancel: "[Cancel]",

/** 
 * Custom text for popup lists
 * @type {string}
 */
s_PopupListCustom: "Custom",

/**
 * s_loc_ constants accessed by SocialCalc.LocalizeString and SocialCalc.LocalizeSubstrings
 * Used extensively by socialcalcspreadsheetcontrol.js
 */

/** @type {string} */ s_loc_align_center: "Align Center",
/** @type {string} */ s_loc_align_left: "Align Left",
/** @type {string} */ s_loc_align_right: "Align Right",
/** @type {string} */ s_loc_alignment: "Alignment",
/** @type {string} */ s_loc_audit: "Audit",
/** @type {string} */ s_loc_audit_trail_this_session: "Audit Trail This Session",
/** @type {string} */ s_loc_auto: "Auto",
/** @type {string} */ s_loc_auto_sum: "Auto Sum",
/** @type {string} */ s_loc_auto_wX_commas: "Auto w/ commas",
/** @type {string} */ s_loc_automatic: "Automatic",
/** @type {string} */ s_loc_background: "Background",
/** @type {string} */ s_loc_bold: "Bold",
/** @type {string} */ s_loc_bold_XampX_italics: "Bold &amp; Italics",
/** @type {string} */ s_loc_bold_italic: "Bold Italic",
/** @type {string} */ s_loc_borders: "Borders",
/** @type {string} */ s_loc_borders_off: "Borders Off",
/** @type {string} */ s_loc_borders_on: "Borders On",
/** @type {string} */ s_loc_bottom: "Bottom",
/** @type {string} */ s_loc_bottom_border: "Bottom Border",
/** @type {string} */ s_loc_cell_settings: "CELL SETTINGS",
/** @type {string} */ s_loc_csv_format: "CSV format",
/** @type {string} */ s_loc_cancel: "Cancel",
/** @type {string} */ s_loc_category: "Category",
/** @type {string} */ s_loc_center: "Center",
/** @type {string} */ s_loc_clear: "Clear",
/** @type {string} */ s_loc_clear_socialcalc_clipboard: "Clear SocialCalc Clipboard",
/** @type {string} */ s_loc_clipboard: "Clipboard",
/** @type {string} */ s_loc_color: "Color",
/** @type {string} */ s_loc_column_: "Column ",
/** @type {string} */ s_loc_comment: "Comment",
/** @type {string} */ s_loc_copy: "Copy",
/** @type {string} */ s_loc_custom: "Custom",
/** @type {string} */ s_loc_cut: "Cut",
/** @type {string} */ s_loc_default: "Default",
/** @type {string} */ s_loc_default_alignment: "Default Alignment",
/** @type {string} */ s_loc_default_column_width: "Default Column Width",
/** @type {string} */ s_loc_default_font: "Default Font",
/** @type {string} */ s_loc_default_format: "Default Format",
/** @type {string} */ s_loc_default_padding: "Default Padding",
/** @type {string} */ s_loc_delete: "Delete",
/** @type {string} */ s_loc_delete_column: "Delete Column",
/** @type {string} */ s_loc_delete_contents: "Delete Contents",
/** @type {string} */ s_loc_delete_row: "Delete Row",
/** @type {string} */ s_loc_description: "Description",
/** @type {string} */ s_loc_display_clipboard_in: "Display Clipboard in",
/** @type {string} */ s_loc_down: "Down",
/** @type {string} */ s_loc_edit: "Edit",
/** @type {string} */ s_loc_existing_names: "Existing Names",
/** @type {string} */ s_loc_family: "Family",
/** @type {string} */ s_loc_fill_down: "Fill Down",
/** @type {string} */ s_loc_fill_right: "Fill Right",
/** @type {string} */ s_loc_font: "Font",
/** @type {string} */ s_loc_format: "Format",
/** @type {string} */ s_loc_formula: "Formula",
/** @type {string} */ s_loc_function_list: "Function List",
/** @type {string} */ s_loc_functions: "Functions",
/** @type {string} */ s_loc_grid: "Grid",
/** @type {string} */ s_loc_hidden: "Hidden",
/** @type {string} */ s_loc_horizontal: "Horizontal",
/** @type {string} */ s_loc_insert_column: "Insert Column",
/** @type {string} */ s_loc_insert_row: "Insert Row",
/** @type {string} */ s_loc_italic: "Italic",
/** @type {string} */ s_loc_last_sort: "Last Sort",
/** @type {string} */ s_loc_left: "Left",
/** @type {string} */ s_loc_left_border: "Left Border",
/** @type {string} */ s_loc_link: "Link",
/** @type {string} */ s_loc_link_input_box: "Link Input Box",
/** @type {string} */ s_loc_list: "List",
/** @type {string} */ s_loc_load_socialcalc_clipboard_with_this: "Load SocialCalc Clipboard With This",
/** @type {string} */ s_loc_major_sort: "Major Sort",
/** @type {string} */ s_loc_manual: "Manual",
/** @type {string} */ s_loc_merge_cells: "Merge Cells",
/** @type {string} */ s_loc_middle: "Middle",
/** @type {string} */ s_loc_minor_sort: "Minor Sort",
/** @type {string} */ s_loc_move_insert: "Move Insert",
/** @type {string} */ s_loc_move_paste: "Move Paste",
/** @type {string} */ s_loc_multiXline_input_box: "Multi-line Input Box",
/** @type {string} */ s_loc_name: "Name",
/** @type {string} */ s_loc_names: "Names",
/** @type {string} */ s_loc_no_padding: "No padding",
/** @type {string} */ s_loc_normal: "Normal",
/** @type {string} */ s_loc_number: "Number",
/** @type {string} */ s_loc_number_horizontal: "Number Horizontal",
/** @type {string} */ s_loc_ok: "OK",
/** @type {string} */ s_loc_padding: "Padding",
/** @type {string} */ s_loc_page_name: "Page Name",
/** @type {string} */ s_loc_paste: "Paste",
/** @type {string} */ s_loc_paste_formats: "Paste Formats",
/** @type {string} */ s_loc_plain_text: "Plain Text",
/** @type {string} */ s_loc_recalc: "Recalc",
/** @type {string} */ s_loc_recalculation: "Recalculation",
/** @type {string} */ s_loc_redo: "Redo",
/** @type {string} */ s_loc_right: "Right",
/** @type {string} */ s_loc_right_border: "Right Border",
/** @type {string} */ s_loc_sheet_settings: "SHEET SETTINGS",
/** @type {string} */ s_loc_save: "Save",
/** @type {string} */ s_loc_save_to: "Save to",
/** @type {string} */ s_loc_set_cell_contents: "Set Cell Contents",
/** @type {string} */ s_loc_set_cells_to_sort: "Set Cells To Sort",
/** @type {string} */ s_loc_set_value_to: "Set Value To",
/** @type {string} */ s_loc_set_to_link_format: "Set to Link format",
/** @type {string} */ s_loc_setXclear_move_from: "Set/Clear Move From",
/** @type {string} */ s_loc_show_cell_settings: "Show Cell Settings",
/** @type {string} */ s_loc_show_sheet_settings: "Show Sheet Settings",
/** @type {string} */ s_loc_show_in_new_browser_window: "Show in new browser window",
/** @type {string} */ s_loc_size: "Size",
/** @type {string} */ s_loc_socialcalcXsave_format: "SocialCalc-save format",
/** @type {string} */ s_loc_sort: "Sort",
/** @type {string} */ s_loc_sort_: "Sort ",
/** @type {string} */ s_loc_sort_cells: "Sort Cells",
/** @type {string} */ s_loc_swap_colors: "Swap Colors",
/** @type {string} */ s_loc_tabXdelimited_format: "Tab-delimited format",
/** @type {string} */ s_loc_text: "Text",
/** @type {string} */ s_loc_text_horizontal: "Text Horizontal",
/** @type {string} */ s_loc_this_is_aXbrXsample: "This is a<br>sample",
/** @type {string} */ s_loc_top: "Top",
/** @type {string} */ s_loc_top_border: "Top Border",
/** @type {string} */ s_loc_undone_steps: "UNDONE STEPS",
/** @type {string} */ s_loc_url: "URL",
/** @type {string} */ s_loc_undo: "Undo",
/** @type {string} */ s_loc_unmerge_cells: "Unmerge Cells",
/** @type {string} */ s_loc_up: "Up",
/** @type {string} */ s_loc_value: "Value",
/** @type {string} */ s_loc_vertical: "Vertical",
/** @type {string} */ s_loc_wikitext: "Wikitext",
/** @type {string} */ s_loc_workspace: "Workspace",
/** @type {string} */ s_loc_XnewX: "[New]",
/** @type {string} */ s_loc_XnoneX: "[None]",
/** @type {string} */ s_loc_Xselect_rangeX: "[select range]",

/**
 * SocialCalc Spreadsheet Viewer module, socialcalcviewer.js
 */

/**
 * SocialCalc.SpreadsheetViewer
 */

/** 
 * Height of status line in pixels
 * @type {number}
 */
SVStatuslineheight: 20,

/** 
 * CSS styling for status line
 * @type {string}
 */
SVStatuslineCSS: "font-size:10px;padding:3px 0px;",

/**
 * SocialCalc Format Number module, formatnumber2.js
 */

/** 
 * The thousands separator character when formatting numbers for display
 * @type {string}
 */
FormatNumber_separatorchar: ",",

/** 
 * The decimal separator character when formatting numbers for display
 * @type {string}
 */
FormatNumber_decimalchar: ".",

/** 
 * The currency string used if none specified
 * @type {string}
 */
FormatNumber_defaultCurrency: "$",

/**
 * The following constants are arrays of strings with the short (3 character) and full names of days and months
 */

/** 
 * Full day names for date formatting
 * @type {string[]}
 */
s_FormatNumber_daynames: ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],

/** 
 * Short (3-character) day names for date formatting
 * @type {string[]}
 */
s_FormatNumber_daynames3: ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"],

/** 
 * Full month names for date formatting
 * @type {string[]}
 */
s_FormatNumber_monthnames: ["January", "February", "March", "April", "May", "June", "July", "August", "September",
                           "October", "November", "December"],

/** 
 * Short (3-character) month names for date formatting
 * @type {string[]}
 */
s_FormatNumber_monthnames3: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],

/** 
 * AM indicator for time formatting
 * @type {string}
 */
s_FormatNumber_am: "AM",

/** 
 * Short AM indicator for time formatting
 * @type {string}
 */
s_FormatNumber_am1: "A",

/** 
 * PM indicator for time formatting
 * @type {string}
 */
s_FormatNumber_pm: "PM",

/** 
 * Short PM indicator for time formatting
 * @type {string}
 */
s_FormatNumber_pm1: "P",

/**
 * SocialCalc Spreadsheet Formula module, formula1.js
 */

/** 
 * Parse error: Improperly formed number exponent
 * @type {string}
 */
s_parseerrexponent: "Improperly formed number exponent",

/** 
 * Parse error: Unexpected character in formula
 * @type {string}
 */
s_parseerrchar: "Unexpected character in formula",

/** 
 * Parse error: Improperly formed string
 * @type {string}
 */
s_parseerrstring: "Improperly formed string",

/** 
 * Parse error: Improperly formed special value
 * @type {string}
 */
s_parseerrspecialvalue: "Improperly formed special value",

/** 
 * Parse error: Two operators inappropriately in a row
 * @type {string}
 */
s_parseerrtwoops: "Error in formula (two operators inappropriately in a row)",

/** 
 * Parse error: Missing open parenthesis in list with comma(s)
 * @type {string}
 */
s_parseerrmissingopenparen: "Missing open parenthesis in list with comma(s). ",

/** 
 * Parse error: Closing parenthesis without open parenthesis
 * @type {string}
 */
s_parseerrcloseparennoopen: "Closing parenthesis without open parenthesis. ",

/** 
 * Parse error: Missing close parenthesis
 * @type {string}
 */
s_parseerrmissingcloseparen: "Missing close parenthesis. ",

/** 
 * Parse error: Missing operand
 * @type {string}
 */
s_parseerrmissingoperand: "Missing operand. ",

/** 
 * Parse error: General error in formula
 * @type {string}
 */
s_parseerrerrorinformula: "Error in formula.",

/** 
 * Calculation error: Error value in formula
 * @type {string}
 */
s_calcerrerrorvalueinformula: "Error value in formula",

/** 
 * Parse error: Error in formula resulting in bad value
 * @type {string}
 */
s_parseerrerrorinformulabadval: "Error in formula resulting in bad value",

/** 
 * Formula results in range value message
 * @type {string}
 */
s_formularangeresult: "Formula results in range value:",

/** 
 * Calculation error: Formula results in bad numeric value
 * @type {string}
 */
s_calcerrnumericnan: "Formula results in an bad numeric value",

/** 
 * Calculation error: Numeric overflow
 * @type {string}
 */
s_calcerrnumericoverflow: "Numeric overflow",

/** 
 * Error message when FindSheetInCache returns null
 * @type {string}
 */
s_sheetunavailable: "Sheet unavailable:",

/** 
 * Calculation error: Cell reference missing when expected
 * @type {string}
 */
s_calcerrcellrefmissing: "Cell reference missing when expected.",

/** 
 * Calculation error: Sheet name missing when expected
 * @type {string}
 */
s_calcerrsheetnamemissing: "Sheet name missing when expected.",

/** 
 * Error message for circular name reference
 * @type {string}
 */
s_circularnameref: "Circular name reference to name",

/** 
 * Calculation error: Unknown name
 * @type {string}
 */
s_calcerrunknownname: "Unknown name",

/** 
 * Calculation error: Incorrect arguments to function
 * @type {string}
 */
s_calcerrincorrectargstofunction: "Incorrect arguments to function",

/** 
 * Sheet function error: Unknown function
 * @type {string}
 */
s_sheetfuncunknownfunction: "Unknown function",

/** 
 * Sheet function error: LN argument must be greater than 0
 * @type {string}
 */
s_sheetfunclnarg: "LN argument must be greater than 0",

/** 
 * Sheet function error: LOG10 argument must be greater than 0
 * @type {string}
 */
s_sheetfunclog10arg: "LOG10 argument must be greater than 0",

/** 
 * Sheet function error: LOG second argument must be numeric greater than 0
 * @type {string}
 */
s_sheetfunclogsecondarg: "LOG second argument must be numeric greater than 0",

/** 
 * Sheet function error: LOG first argument must be greater than 0
 * @type {string}
 */
s_sheetfunclogfirstarg: "LOG first argument must be greater than 0",

/** 
 * Sheet function error: ROUND second argument must be numeric
 * @type {string}
 */
s_sheetfuncroundsecondarg: "ROUND second argument must be numeric",

/** 
 * Sheet function error: DDB life must be greater than 1
 * @type {string}
 */
s_sheetfuncddblife: "DDB life must be greater than 1",

/** 
 * Sheet function error: SLN life must be greater than 1
 * @type {string}
 */
s_sheetfuncslnlife: "SLN life must be greater than 1",

/**
 * Function definition text - descriptions for spreadsheet functions
 */

/** @type {string} */ s_fdef_ABS: 'Absolute value function. ',
/** @type {string} */ s_fdef_ACOS: 'Trigonometric arccosine function. ',
/** @type {string} */ s_fdef_AND: 'True if all arguments are true. ',
/** @type {string} */ s_fdef_ASIN: 'Trigonometric arcsine function. ',
/** @type {string} */ s_fdef_ATAN: 'Trigonometric arctan function. ',
/** @type {string} */ s_fdef_ATAN2: 'Trigonometric arc tangent function (result is in radians). ',
/** @type {string} */ s_fdef_AVERAGE: 'Averages the values. ',
/** @type {string} */ s_fdef_CHOOSE: 'Returns the value specified by the index. The values may be ranges of cells. ',
/** @type {string} */ s_fdef_COLUMNS: 'Returns the number of columns in the range. ',
/** @type {string} */ s_fdef_COS: 'Trigonometric cosine function (value is in radians). ',
/** @type {string} */ s_fdef_COUNT: 'Counts the number of numeric values, not blank, text, or error. ',
/** @type {string} */ s_fdef_COUNTA: 'Counts the number of non-blank values. ',
/** @type {string} */ s_fdef_COUNTBLANK: 'Counts the number of blank values. (Note: "" is not blank.) ',
/** @type {string} */ s_fdef_COUNTIF: 'Counts the number of number of cells in the range that meet the criteria. The criteria may be a value ("x", 15, 1+3) or a test (>25). ',
/** @type {string} */ s_fdef_DATE: 'Returns the appropriate date value given numbers for year, month, and day. For example: DATE(2006,2,1) for February 1, 2006. Note: In this program, day "1" is December 31, 1899 and the year 1900 is not a leap year. Some programs use January 1, 1900, as day "1" and treat 1900 as a leap year. In both cases, though, dates on or after March 1, 1900, are the same. ',
/** @type {string} */ s_fdef_DAVERAGE: 'Averages the values in the specified field in records that meet the criteria. ',
/** @type {string} */ s_fdef_DAY: 'Returns the day of month for a date value. ',
/** @type {string} */ s_fdef_DCOUNT: 'Counts the number of numeric values, not blank, text, or error, in the specified field in records that meet the criteria. ',
/** @type {string} */ s_fdef_DCOUNTA: 'Counts the number of non-blank values in the specified field in records that meet the criteria. ',
/** @type {string} */ s_fdef_DDB: 'Returns the amount of depreciation at the given period of time (the default factor is 2 for double-declining balance).   ',
/** @type {string} */ s_fdef_DEGREES: 'Converts value in radians into degrees. ',
/** @type {string} */ s_fdef_DGET: 'Returns the value of the specified field in the single record that meets the criteria. ',
/** @type {string} */ s_fdef_DMAX: 'Returns the maximum of the numeric values in the specified field in records that meet the criteria. ',
/** @type {string} */ s_fdef_DMIN: 'Returns the maximum of the numeric values in the specified field in records that meet the criteria. ',
/** @type {string} */ s_fdef_DPRODUCT: 'Returns the result of multiplying the numeric values in the specified field in records that meet the criteria. ',
/** @type {string} */ s_fdef_DSTDEV: 'Returns the sample standard deviation of the numeric values in the specified field in records that meet the criteria. ',
/** @type {string} */ s_fdef_DSTDEVP: 'Returns the standard deviation of the numeric values in the specified field in records that meet the criteria. ',
/** @type {string} */ s_fdef_DSUM: 'Returns the sum of the numeric values in the specified field in records that meet the criteria. ',
/** @type {string} */ s_fdef_DVAR: 'Returns the sample variance of the numeric values in the specified field in records that meet the criteria. ',
/** @type {string} */ s_fdef_DVARP: 'Returns the variance of the numeric values in the specified field in records that meet the criteria. ',
/** @type {string} */ s_fdef_EVEN: 'Rounds the value up in magnitude to the nearest even integer. ',
/** @type {string} */ s_fdef_EXACT: 'Returns "true" if the values are exactly the same, including case, type, etc. ',
/** @type {string} */ s_fdef_EXP: 'Returns e raised to the value power. ',
/** @type {string} */ s_fdef_FACT: 'Returns factorial of the value. ',
/** @type {string} */ s_fdef_FALSE: 'Returns the logical value "false". ',
/** @type {string} */ s_fdef_FIND: 'Returns the starting position within string2 of the first occurrence of string1 at or after "start". If start is omitted, 1 is assumed. ',
/** @type {string} */ s_fdef_FV: 'Returns the future value of repeated payments of money invested at the given rate for the specified number of periods, with optional present value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period). ',
/** @type {string} */ s_fdef_HLOOKUP: 'Look for the matching value for the given value in the range and return the corresponding value in the cell specified by the row offset. If rangelookup is 1 (the default) and not 0, match if within numeric brackets (match<=value) instead of exact match. ',
/** @type {string} */ s_fdef_HOUR: 'Returns the hour portion of a time or date/time value. ',
/** @type {string} */ s_fdef_IF: 'Results in true-value if logical-expression is TRUE or non-zero, otherwise results in false-value. ',
/** @type {string} */ s_fdef_INDEX: 'Returns a cell or range reference for the specified row and column in the range. If range is 1-dimensional, then only one of rownum or colnum are needed. If range is 2-dimensional and rownum or colnum are zero, a reference to the range of just the specified column or row is returned. You can use the returned reference value in a range, e.g., sum(A1:INDEX(A2:A10,4)). ',
/** @type {string} */ s_fdef_INT: 'Returns the value rounded down to the nearest integer (towards -infinity). ',
/** @type {string} */ s_fdef_IRR: 'Returns the interest rate at which the cash flows in the range have a net present value of zero. Uses an iterative process that will return #NUM! error if it does not converge. There may be more than one possible solution. Providing the optional guess value may help in certain situations where it does not converge or finds an inappropriate solution (the default guess is 10%). ',
/** @type {string} */ s_fdef_ISBLANK: 'Returns "true" if the value is a reference to a blank cell. ',
/** @type {string} */ s_fdef_ISERR: 'Returns "true" if the value is of type "Error" but not "NA". ',
/** @type {string} */ s_fdef_ISERROR: 'Returns "true" if the value is of type "Error". ',
/** @type {string} */ s_fdef_ISLOGICAL: 'Returns "true" if the value is of type "Logical" (true/false). ',
/** @type {string} */ s_fdef_ISNA: 'Returns "true" if the value is the error type "NA". ',
/** @type {string} */ s_fdef_ISNONTEXT: 'Returns "true" if the value is not of type "Text". ',
/** @type {string} */ s_fdef_ISNUMBER: 'Returns "true" if the value is of type "Number" (including logical values). ',
/** @type {string} */ s_fdef_ISTEXT: 'Returns "true" if the value is of type "Text". ',
/** @type {string} */ s_fdef_LEFT: 'Returns the specified number of characters from the text value. If count is omitted, 1 is assumed. ',
/** @type {string} */ s_fdef_LEN: 'Returns the number of characters in the text value. ',
/** @type {string} */ s_fdef_LN: 'Returns the natural logarithm of the value. ',
/** @type {string} */ s_fdef_LOG: 'Returns the logarithm of the value using the specified base. ',
/** @type {string} */ s_fdef_LOG10: 'Returns the base 10 logarithm of the value. ',
/** @type {string} */ s_fdef_LOWER: 'Returns the text value with all uppercase characters converted to lowercase. ',
/** @type {string} */ s_fdef_MATCH: 'Look for the matching value for the given value in the range and return position (the first is 1) in that range. If rangelookup is 1 (the default) and not 0, match if within numeric brackets (match<=value) instead of exact match. If rangelookup is -1, act like 1 but the bracket is match>=value. ',
/** @type {string} */ s_fdef_MAX: 'Returns the maximum of the numeric values. ',
/** @type {string} */ s_fdef_MID: 'Returns the specified number of characters from the text value starting from the specified position. ',
/** @type {string} */ s_fdef_MIN: 'Returns the minimum of the numeric values. ',
/** @type {string} */ s_fdef_MINUTE: 'Returns the minute portion of a time or date/time value. ',
/** @type {string} */ s_fdef_MOD: 'Returns the remainder of the first value divided by the second. ',
/** @type {string} */ s_fdef_MONTH: 'Returns the month part of a date value. ',
/** @type {string} */ s_fdef_N: 'Returns the value if it is a numeric value otherwise an error. ',
/** @type {string} */ s_fdef_NA: 'Returns the #N/A error value which propagates through most operations. ',
/** @type {string} */ s_fdef_NOT: 'Returns FALSE if value is true, and TRUE if it is false. ',
/** @type {string} */ s_fdef_NOW: 'Returns the current date/time. ',
/** @type {string} */ s_fdef_NPER: 'Returns the number of periods at which payments invested each period at the given rate with optional future value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period) has the given present value. ',
/** @type {string} */ s_fdef_NPV: 'Returns the net present value of cash flows (which may be individual values and/or ranges) at the given rate. The flows are positive if income, negative if paid out, and are assumed at the end of each period. ',
/** @type {string} */ s_fdef_ODD: 'Rounds the value up in magnitude to the nearest odd integer. ',
/** @type {string} */ s_fdef_OR: 'True if any argument is true ',
/** @type {string} */ s_fdef_PI: 'The value 3.1415926... ',
/** @type {string} */ s_fdef_PMT: 'Returns the amount of each payment that must be invested at the given rate for the specified number of periods to have the specified present value, with optional future value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period). ',
/** @type {string} */ s_fdef_POWER: 'Returns the first value raised to the second value power. ',
/** @type {string} */ s_fdef_PRODUCT: 'Returns the result of multiplying the numeric values. ',
/** @type {string} */ s_fdef_PROPER: 'Returns the text value with the first letter of each word converted to uppercase and the others to lowercase. ',
/** @type {string} */ s_fdef_PV: 'Returns the present value of the given number of payments each invested at the given rate, with optional future value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period). ',
/** @type {string} */ s_fdef_RADIANS: 'Converts value in degrees into radians. ',
/** @type {string} */ s_fdef_RATE: 'Returns the rate at which the given number of payments each invested at the given rate has the specified present value, with optional future value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period). Uses an iterative process that will return #NUM! error if it does not converge. There may be more than one possible solution. Providing the optional guess value may help in certain situations where it does not converge or finds an inappropriate solution (the default guess is 10%). ',
/** @type {string} */ s_fdef_REPLACE: 'Returns text1 with the specified number of characters starting from the specified position replaced by text2. ',
/** @type {string} */ s_fdef_REPT: 'Returns the text repeated the specified number of times. ',
/** @type {string} */ s_fdef_RIGHT: 'Returns the specified number of characters from the text value starting from the end. If count is omitted, 1 is assumed. ',
/** @type {string} */ s_fdef_ROUND: 'Rounds the value to the specified number of decimal places. If precision is negative, then round to powers of 10. The default precision is 0 (round to integer). ',
/** @type {string} */ s_fdef_ROWS: 'Returns the number of rows in the range. ',
/** @type {string} */ s_fdef_SECOND: 'Returns the second portion of a time or date/time value (truncated to an integer). ',
/** @type {string} */ s_fdef_SIN: 'Trigonometric sine function (value is in radians) ',
/** @type {string} */ s_fdef_SLN: 'Returns the amount of depreciation at each period of time using the straight-line method. ',
/** @type {string} */ s_fdef_SQRT: 'Square root of the value ',
/** @type {string} */ s_fdef_STDEV: 'Returns the sample standard deviation of the numeric values. ',
/** @type {string} */ s_fdef_STDEVP: 'Returns the standard deviation of the numeric values. ',
/** @type {string} */ s_fdef_SUBSTITUTE: 'Returns text1 with the all occurrences of oldtext replaced by newtext. If "occurrence" is present, then only that occurrence is replaced. ',
/** @type {string} */ s_fdef_SUM: 'Adds the numeric values. The values to the sum function may be ranges in the form similar to A1:B5. ',
/** @type {string} */ s_fdef_SUMIF: 'Sums the numeric values of cells in the range that meet the criteria. The criteria may be a value ("x", 15, 1+3) or a test (>25). If range2 is present, then range1 is tested and the corresponding range2 value is summed. ',
/** @type {string} */ s_fdef_SYD: 'Depreciation by Sum of Year\'s Digits method. ',
/** @type {string} */ s_fdef_T: 'Returns the text value or else a null string. ',
/** @type {string} */ s_fdef_TAN: 'Trigonometric tangent function (value is in radians) ',
/** @type {string} */ s_fdef_TIME: 'Returns the time value given the specified hour, minute, and second. ',
/** @type {string} */ s_fdef_TODAY: 'Returns the current date (an integer). Note: In this program, day "1" is December 31, 1899 and the year 1900 is not a leap year. Some programs use January 1, 1900, as day "1" and treat 1900 as a leap year. In both cases, though, dates on or after March 1, 1900, are the same. ',
/** @type {string} */ s_fdef_TRIM: 'Returns the text value with leading, trailing, and repeated spaces removed. ',
/** @type {string} */ s_fdef_TRUE: 'Returns the logical value "true". ',
/** @type {string} */ s_fdef_TRUNC: 'Truncates the value to the specified number of decimal places. If precision is negative, truncate to powers of 10. ',
/** @type {string} */ s_fdef_UPPER: 'Returns the text value with all lowercase characters converted to uppercase. ',
/** @type {string} */ s_fdef_VALUE: 'Converts the specified text value into a numeric value. Various forms that look like numbers (including digits followed by %, forms that look like dates, etc.) are handled. This may not handle all of the forms accepted by other spreadsheets and may be locale dependent. ',
/** @type {string} */ s_fdef_VAR: 'Returns the sample variance of the numeric values. ',
/** @type {string} */ s_fdef_VARP: 'Returns the variance of the numeric values. ',
/** @type {string} */ s_fdef_VLOOKUP: 'Look for the matching value for the given value in the range and return the corresponding value in the cell specified by the column offset. If rangelookup is 1 (the default) and not 0, match if within numeric brackets (match>=value) instead of exact match. ',
/** @type {string} */ s_fdef_WEEKDAY: 'Returns the day of week specified by the date value. If type is 1 (the default), Sunday is day and Saturday is day 7. If type is 2, Monday is day 1 and Sunday is day 7. If type is 3, Monday is day 0 and Sunday is day 6. ',
/** @type {string} */ s_fdef_YEAR: 'Returns the year part of a date value. ',

/**
 * Function argument patterns for function help
 */

/** @type {string} */ s_farg_v: "value",
/** @type {string} */ s_farg_vn: "value1, value2, ...",
/** @type {string} */ s_farg_xy: "valueX, valueY",
/** @type {string} */ s_farg_choose: "index, value1, value2, ...",
/** @type {string} */ s_farg_range: "range",
/** @type {string} */ s_farg_rangec: "range, criteria",
/** @type {string} */ s_farg_date: "year, month, day",
/** @type {string} */ s_farg_dfunc: "databaserange, fieldname, criteriarange",
/** @type {string} */ s_farg_ddb: "cost, salvage, lifetime, period [, factor]",
/** @type {string} */ s_farg_find: "string1, string2 [, start]",
/** @type {string} */ s_farg_fv: "rate, n, payment, [pv, [paytype]]",
/** @type {string} */ s_farg_hlookup: "value, range, row, [rangelookup]",
/** @type {string} */ s_farg_iffunc: "logical-expression, true-value, false-value",
/** @type {string} */ s_farg_index: "range, rownum, colnum",
/** @type {string} */ s_farg_irr: "range, [guess]",
/** @type {string} */ s_farg_tc: "text, count",
/** @type {string} */ s_farg_log: "value, base",
/** @type {string} */ s_farg_match: "value, range, [rangelookup]",
/** @type {string} */ s_farg_mid: "text, start, length",
/** @type {string} */ s_farg_nper: "rate, payment, pv, [fv, [paytype]]",
/** @type {string} */ s_farg_npv: "rate, value1, value2, ...",
/** @type {string} */ s_farg_pmt: "rate, n, pv, [fv, [paytype]]",
/** @type {string} */ s_farg_pv: "rate, n, payment, [fv, [paytype]]",
/** @type {string} */ s_farg_rate: "n, payment, pv, [fv, [paytype, [guess]]]",
/** @type {string} */ s_farg_replace: "text1, start, length, text2",
/** @type {string} */ s_farg_vp: "value, [precision]",
/** @type {string} */ s_farg_valpre: "value, precision",
/** @type {string} */ s_farg_csl: "cost, salvage, lifetime",
/** @type {string} */ s_farg_cslp: "cost, salvage, lifetime, period",
/** @type {string} */ s_farg_subs: "text1, oldtext, newtext [, occurrence]",
/** @type {string} */ s_farg_sumif: "range1, criteria [, range2]",
/** @type {string} */ s_farg_hms: "hour, minute, second",
/** @type {string} */ s_farg_txt: "text",
/** @type {string} */ s_farg_vlookup: "value, range, col, [rangelookup]",
/** @type {string} */ s_farg_weekday: "date, [type]",
/** @type {string} */ s_farg_dt: "date",

/**
 * Order of function classes for categorization
 * @type {string[]}
 */
function_classlist: ["all", "stat", "lookup", "datetime", "financial", "test", "math", "text"],

/**
 * Function class display names
 */

/** @type {string} */ s_fclass_all: "All",
/** @type {string} */ s_fclass_stat: "Statistics",
/** @type {string} */ s_fclass_lookup: "Lookup",
/** @type {string} */ s_fclass_datetime: "Date & Time",
/** @type {string} */ s_fclass_financial: "Financial",
/** @type {string} */ s_fclass_test: "Test",
/** @type {string} */ s_fclass_math: "Math",
/** @type {string} */ s_fclass_text: "Text",

/** 
 * Marker for end of constants object
 * @type {null}
 */
lastone: null

};

/**
 * Default CSS classnames for use with SocialCalc.ConstantsSetClasses
 * Maps constant names to their default CSS class values
 * @namespace SocialCalc.ConstantsDefaultClasses
 */
SocialCalc.ConstantsDefaultClasses = {
   /** @type {string} */ defaultComment: "",
   /** @type {string} */ defaultCommentNoGrid: "",
   /** @type {string} */ defaultHighlightTypeCursor: "",
   /** @type {string} */ defaultHighlightTypeRange: "",
   /** @type {string} */ defaultColname: "",
   /** @type {string} */ defaultSelectedColname: "",
   /** @type {string} */ defaultRowname: "",
   /** @type {string} */ defaultSelectedRowname: "", 
   /** @type {string} */ defaultUpperLeft: "",
   /** @type {string} */ defaultSkippedCell: "",
   /** @type {string} */ defaultPaneDivider: "",
   /** @type {string} */ cteGriddiv: "", // this one has no Style version with it
   /** 
    * Special object with classname and style for FireFox compatibility
    * @type {Object}
    */
   defaultInputEcho: {classname: "", style: "filter:alpha(opacity=90);opacity:.9;"}, // so FireFox won't show warning
   /** @type {string} */ TCmain: "",
   /** @type {string} */ TCendcap: "",
   /** @type {string} */ TCpaneslider: "",
   /** @type {string} */ TClessbutton: "",
   /** @type {string} */ TCmorebutton: "",
   /** @type {string} */ TCscrollarea: "",
   /** @type {string} */ TCthumb: "",
   /** @type {string} */ TCPStrackingline: "",
   /** @type {string} */ TCTDFSthumbstatus: "",
   /** @type {string} */ TDpopupElement: ""
};

/**
 * Sets CSS classes with optional prefix for all constant items
 * 
 * This routine goes through all of the xyzClass/xyzStyle pairs and sets the Class to a default and
 * turns off the Style, if present. The prefix is put before each default.
 * The list of items to set is in SocialCalc.ConstantsDefaultClasses. The names there
 * correspond to the "xyz" parts. If there is a value, it is the default to set. If the
 * default is a null, no change is made. If the default is the null string (""), the
 * name of the item is used (e.g., "defaultComment" would use the classname "defaultComment").
 * If the default is an object, then it expects {classname: classname, style: stylestring} - this
 * lets you combine both.
 * 
 * @param {string} [prefix=''] - Prefix to add to class names
 */
SocialCalc.ConstantsSetClasses = (prefix = '') => {
   const defaults = SocialCalc.ConstantsDefaultClasses;
   const scc = SocialCalc.Constants;

   for (const item in defaults) {
      if (typeof defaults[item] === 'string') {
         scc[`${item}Class`] = prefix + (defaults[item] || item);
         if (scc[`${item}Style`] !== undefined) {
            scc[`${item}Style`] = '';
         }
      } else if (typeof defaults[item] === 'object') {
         scc[`${item}Class`] = prefix + (defaults[item].classname || item);
         scc[`${item}Style`] = defaults[item].style;
      }
   }
};
