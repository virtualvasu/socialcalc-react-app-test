/**
 * @fileoverview SocialCalc Number Formatting Library
 * Part of the SocialCalc package for handling number, date, currency, and data type formatting
 * according to spreadsheet formatting conventions.
 * 
 * @version 1.1.0+
 * @author Dan Bricklin (Software Garden, Inc. for Socialtext, Inc.)
 * @copyright 2008 Socialtext, Inc. All Rights Reserved
 * @license Artistic License 2.0 - http://socialcalc.org/licenses/al-20/
 * 
 * Based in part on the SocialCalc 1.1.0 code written in Perl.
 * The SocialCalc 1.1.0 code was:
 *    Portions (c) Copyright 2005, 2006, 2007 Software Garden, Inc.
 *    Portions (c) Copyright 2007 Socialtext, Inc.
 * The Perl SocialCalc started as modifications to the wikiCalc(R) program, version 1.0.
 * wikiCalc 1.0 was written by Software Garden, Inc.
 */

// Initialize SocialCalc namespace if not exists
var SocialCalc;
if (!SocialCalc) SocialCalc = {};

/**
 * SocialCalc Number Formatting Library
 * Handles formatting of numbers, dates, currencies, and other data types
 * according to spreadsheet formatting conventions.
 * @namespace SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber = {};

/**
 * Global cache for parsed format definitions
 * @type {Object<string, Object>}
 * @memberof SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber.format_definitions = {};

/**
 * Character used for thousands separator in number formatting
 * @type {string}
 * @default ','
 * @memberof SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber.separatorchar = ',';

/**
 * Character used for decimal point in number formatting
 * @type {string}
 * @default '.'
 * @memberof SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber.decimalchar = '.';

/**
 * Full day names for date formatting (Sunday through Saturday)
 * @type {string[]}
 * @memberof SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber.daynames = [
    'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'
];

/**
 * Abbreviated day names for date formatting (Sun through Sat)
 * @type {string[]}
 * @memberof SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber.daynames3 = [
    'Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'
];

/**
 * Abbreviated month names for date formatting (Jan through Dec)
 * @type {string[]}
 * @memberof SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber.monthnames3 = [
    'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'
];

/**
 * Full month names for date formatting (January through December)
 * @type {string[]}
 * @memberof SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber.monthnames = [
    'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September',
    'October', 'November', 'December'
];

/**
 * Allowed color values for formatting with their hex codes
 * @type {Object<string, string>}
 * @memberof SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber.allowedcolors = {
    BLACK: '#000000',
    BLUE: '#0000FF',
    CYAN: '#00FFFF',
    GREEN: '#00FF00',
    MAGENTA: '#FF00FF',
    RED: '#FF0000',
    WHITE: '#FFFFFF',
    YELLOW: '#FFFF00'
};

/**
 * Allowed date format tokens for time formatting
 * @type {Object<string, string>}
 * @memberof SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber.alloweddates = {
    H: 'h]',
    M: 'm]',
    MM: 'mm]',
    S: 's]',
    SS: 'ss]'
};

/**
 * Format command constants for parsing and processing format strings
 * @type {Object<string, number>}
 * @memberof SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber.commands = {
    copy: 1,
    color: 2,
    integer_placeholder: 3,
    fraction_placeholder: 4,
    decimal: 5,
    currency: 6,
    general: 7,
    separator: 8,
    date: 9,
    comparison: 10,
    section: 11,
    style: 12
};

/**
 * Date calculation constants for Julian date conversions and time calculations
 * @type {Object<string, number>}
 * @memberof SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber.datevalues = {
    julian_offset: 2415019,
    seconds_in_a_day: 24 * 60 * 60,
    seconds_in_an_hour: 60 * 60
};

/**
 * Format a number according to the specified format string
 * @param {number|string} rawvalue - The number to format
 * @param {string} format_string - The format specification string
 * @param {string} [currency_char] - Currency symbol to use (optional)
 * @returns {string} The formatted number string
 * @throws {string} Throws "Format not parsed error!" if format parsing fails
 * @memberof SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber.formatNumberWithFormat = (rawvalue, format_string, currency_char) => {
    const scc = SocialCalc.Constants;
    const scfn = SocialCalc.FormatNumber;

    // Variable declarations with better organization
    let op, operandstr, fromend, cval, operandstrlc;
    let startval, estartval;
    let hrs, mins, secs, ehrs, emins, esecs, ampmstr, ymd;
    let minOK, mpos;
    let result = '';
    let thisformat;
    let section, gotcomparison, compop, compval, cpos, oppos;
    let sectioninfo;
    let i, decimalscale, scaledvalue, strvalue, strparts, integervalue, fractionvalue;
    let integerdigits2, integerpos, fractionpos, textcolor, textstyle, separatorchar, decimalchar;
    
    // Ensure numeric input and handle edge cases
    rawvalue = Number(rawvalue);
    let value = rawvalue;
    if (!isFinite(value)) return 'NaN';

    // Handle negative values
    const negativevalue = value < 0;
    if (negativevalue) value = -value;
    const zerovalue = value === 0 ? 1 : 0;

    // Set default currency character
    currency_char = currency_char || scc.FormatNumber_DefaultCurrency;

    // Parse format string and get format structure
    scfn.parse_format_string(scfn.format_definitions, format_string);
    thisformat = scfn.format_definitions[format_string];

    if (!thisformat) throw "Format not parsed error!";

    // Get number of sections - 1
    section = thisformat.sectioninfo.length - 1;

    // Handle format sections with comparisons
    if (thisformat.hascomparison) {
        section = 0; // set to which section we will use
        gotcomparison = 0; // this section has no comparison
        
        // Scan for comparisons
        for (cpos = 0; ; cpos++) {
            op = thisformat.operators[cpos];
            operandstr = thisformat.operands[cpos]; // get next operator and operand
            
            if (!op) { // at end with no match
                if (gotcomparison) { // if comparison but no match
                    format_string = "General"; // use default of General
                    scfn.parse_format_string(scfn.format_definitions, format_string);
                    thisformat = scfn.format_definitions[format_string];
                    section = 0;
                }
                break; // if no comparison, matches on this section
            }
            
            if (op === scfn.commands.section) { // end of section
                if (!gotcomparison) { // no comparison, so it's a match
                    break;
                }
                gotcomparison = 0;
                section++; // check out next one
                continue;
            }
            
            if (op === scfn.commands.comparison) { // found a comparison - do we meet it?
                i = operandstr.indexOf(":");
                compop = operandstr.substring(0, i);
                compval = operandstr.substring(i + 1) - 0;
                
                if ((compop === "<" && rawvalue < compval) ||
                    (compop === "<=" && rawvalue <= compval) ||
                    (compop === "=" && rawvalue === compval) ||
                    (compop === "<>" && rawvalue !== compval) ||
                    (compop === ">=" && rawvalue >= compval) ||
                    (compop === ">" && rawvalue > compval)) { // a match
                    break;
                }
                gotcomparison = 1;
            }
        }
    }
    // Handle multiple sections (separated by ";")
    else if (section > 0) {
        if (section === 1) { // two sections
            if (negativevalue) {
                negativevalue = 0; // sign will be provided by section, not automatically
                section = 1; // use second section for negative values
            } else {
                section = 0; // use first for all others
            }
        }
        else if (section === 2) { // three sections
            if (negativevalue) {
                negativevalue = 0; // sign will be provided by section, not automatically
                section = 1; // use second section for negative values
            }
            else if (zerovalue) {
                section = 2; // use third section for zero values
            }
            else {
                section = 0; // use first for positive
            }
        }
    }

    // Look at values for our section
    sectioninfo = thisformat.sectioninfo[section];
    // Handle scaling by thousands (commas)
    if (sectioninfo.commas > 0) {
        for (i = 0; i < sectioninfo.commas; i++) {
            value /= 1000;
        }
    }

    // Handle percent scaling
    if (sectioninfo.percent > 0) {
        for (i = 0; i < sectioninfo.percent; i++) {
            value *= 100;
        }
    }

    // Cut down to required number of decimal digits
    decimalscale = 1;
    for (i = 0; i < sectioninfo.fractiondigits; i++) {
        decimalscale *= 10;
    }
    scaledvalue = Math.floor(value * decimalscale + 0.5);
    scaledvalue = scaledvalue / decimalscale;

    // Validate scaled value
    if (typeof scaledvalue !== "number") return "NaN";
    if (!isFinite(scaledvalue)) return "NaN";

    // Convert to string (Number.toFixed doesn't do all we need)
    strvalue = scaledvalue + "";

    // Handle zero values to prevent "-0" display
    if (scaledvalue === 0 && (sectioninfo.fractiondigits || sectioninfo.integerdigits)) {
        negativevalue = 0; // no "-0" unless using multiple sections or General
    }

    // Handle scientific notation
    if (strvalue.indexOf("e") >= 0) {
        return rawvalue + ""; // Just return plain converted raw value
    }

    // Parse integer and fraction parts
    strparts = strvalue.match(/^\+{0,1}(\d*)(?:\.(\d*)){0,1}$/);
    if (!strparts) return "NaN"; // if not a number
    
    integervalue = strparts[1];
    if (!integervalue || integervalue === "0") integervalue = "";
    
    fractionvalue = strparts[2];
    if (!fractionvalue) fractionvalue = "";

    // Handle date formatting
    if (sectioninfo.hasdate) {
        if (rawvalue < 0) { // bad date
            return "??-???-??&nbsp;??:??:??";
        }

        // Get date/time parts
        startval = (rawvalue - Math.floor(rawvalue)) * scfn.datevalues.seconds_in_a_day;
        estartval = rawvalue * scfn.datevalues.seconds_in_a_day; // elapsed time version

        // Calculate hours, minutes, seconds
        hrs = Math.floor(startval / scfn.datevalues.seconds_in_an_hour);
        ehrs = Math.floor(estartval / scfn.datevalues.seconds_in_an_hour);
        startval = startval - hrs * scfn.datevalues.seconds_in_an_hour;
        mins = Math.floor(startval / 60);
        emins = Math.floor(estartval / 60);
        secs = startval - mins * 60;

        // Round appropriately depending if there is ss.0
        decimalscale = 1;
        for (i = 0; i < sectioninfo.fractiondigits; i++) {
            decimalscale *= 10;
        }
        secs = Math.floor(secs * decimalscale + 0.5);
        secs = secs / decimalscale;
        esecs = Math.floor(estartval * decimalscale + 0.5);
        esecs = esecs / decimalscale;

        // Handle round up into next second, minute, etc.
        if (secs >= 60) {
            secs = 0;
            mins++; 
            emins++;
            if (mins >= 60) {
                mins = 0;
                hrs++; 
                ehrs++;
                if (hrs >= 24) {
                    hrs = 0;
                    rawvalue++;
                }
            }
        }

        // For "hh:mm:ss.000" formatting
        fractionvalue = (secs - Math.floor(secs)) + "";
        fractionvalue = fractionvalue.substring(2); // skip "0."

        // Convert Julian date to Gregorian
        ymd = SocialCalc.FormatNumber.convert_date_julian_to_gregorian(
            Math.floor(rawvalue + scfn.datevalues.julian_offset)
        );

        // Process minute/month disambiguation and AM/PM
        minOK = 0; // says "m" can be minutes if true
        let mspos = sectioninfo.sectionstart; // m scan position in ops

        // Forward scan for "m" and "mm" to see if any minutes fields, and am/pm
        for (; ; mspos++) {
            op = thisformat.operators[mspos];
            operandstr = thisformat.operands[mspos];
            
            if (!op) break; // don't go past end
            if (op === scfn.commands.section) break;
            
            if (op === scfn.commands.date) {
                // Handle AM/PM formatting
                if ((operandstr.toLowerCase() === "am/pm" || operandstr.toLowerCase() === "a/p") && !ampmstr) {
                    if (hrs >= 12) {
                        hrs -= 12;
                        ampmstr = operandstr.toLowerCase() === "a/p" ? 
                            scc.s_FormatNumber_pm1 : scc.s_FormatNumber_pm; // "P" : "PM"
                    } else {
                        ampmstr = operandstr.toLowerCase() === "a/p" ? 
                            scc.s_FormatNumber_am1 : scc.s_FormatNumber_am; // "A" : "AM"
                    }
                    if (operandstr.indexOf(ampmstr) < 0) {
                        ampmstr = ampmstr.toLowerCase(); // have case match case in format
                    }
                }
                
                // Convert "m" to minutes if following hours
                if (minOK && (operandstr === "m" || operandstr === "mm")) {
                    thisformat.operands[mspos] += "in"; // turn into "min" or "mmin"
                }
                
                if (operandstr.charAt(0) === "h") {
                    minOK = 1; // m following h or hh or [h] is minutes not months
                } else {
                    minOK = 0;
                }
            }
            else if (op !== scfn.commands.copy) { // copying chars can be between h and m
                minOK = 0;
            }
        }

        // Backward scan for seconds after minutes
        minOK = 0;
        for (--mspos; ; mspos--) {
            op = thisformat.operators[mspos];
            operandstr = thisformat.operands[mspos];
            
            if (!op) break; // don't go past end
            if (op === scfn.commands.section) break;
            
            if (op === scfn.commands.date) {
                if (minOK && (operandstr === "m" || operandstr === "mm")) {
                    thisformat.operands[mspos] += "in"; // turn into "min" or "mmin"
                }
                
                if (operandstr === "ss") {
                    minOK = 1; // m before ss is minutes not months
                } else {
                    minOK = 0;
                }
            }
            else if (op !== scfn.commands.copy) { // copying chars can be between ss and m
                minOK = 0;
            }
        }
    }

    // Initialize counters and formatting variables
    integerdigits2 = 0;
    integerpos = 0;
    fractionpos = 0;
    textcolor = "";
    textstyle = "";
    
    // Set up separator and decimal characters
    separatorchar = scc.FormatNumber_separatorchar;
    if (separatorchar.indexOf(" ") >= 0) {
        separatorchar = separatorchar.replace(/ /g, "&nbsp;");
    }
    
    decimalchar = scc.FormatNumber_decimalchar;
    if (decimalchar.indexOf(" ") >= 0) {
        decimalchar = decimalchar.replace(/ /g, "&nbsp;");
    }

    oppos = sectioninfo.sectionstart;

    // Execute format commands
    while (op = thisformat.operators[oppos]) {
        operandstr = thisformat.operands[oppos++]; // get next operator and operand

        if (op === scfn.commands.copy) {
            // Put character in result
            result += operandstr;
        }
        else if (op === scfn.commands.color) {
            // Set text color
            textcolor = operandstr;
        }
        else if (op === scfn.commands.style) {
            // Set text style
            textstyle = operandstr;
        }
        else if (op === scfn.commands.integer_placeholder) {
            // Insert integer part of number
            if (negativevalue) {
                result += "-";
                negativevalue = 0;
            }
            
            integerdigits2++;
            
            if (integerdigits2 === 1) { // first integer placeholder
                // Check if integer is wider than field
                if (integervalue.length > sectioninfo.integerdigits) {
                    for (; integerpos < (integervalue.length - sectioninfo.integerdigits); integerpos++) {
                        result += integervalue.charAt(integerpos);
                        if (sectioninfo.thousandssep) { // check separator position
                            fromend = integervalue.length - integerpos - 1;
                            if (fromend > 2 && fromend % 3 === 0) {
                                result += separatorchar;
                            }
                        }
                    }
                }
            }
            
            // Handle field wider than value
            if (integervalue.length < sectioninfo.integerdigits
                && integerdigits2 <= sectioninfo.integerdigits - integervalue.length) {
                if (operandstr === "0" || operandstr === "?") {
                    result += operandstr === "0" ? "0" : "&nbsp;";
                    if (sectioninfo.thousandssep) { // check separator position
                        fromend = sectioninfo.integerdigits - integerdigits2;
                        if (fromend > 2 && fromend % 3 === 0) {
                            result += separatorchar;
                        }
                    }
                }
            }
            else { // normal integer digit - add it
                result += integervalue.charAt(integerpos);
                if (sectioninfo.thousandssep) { // check separator position
                    fromend = integervalue.length - integerpos - 1;
                    if (fromend > 2 && fromend % 3 === 0) {
                        result += separatorchar;
                    }
                }
                integerpos++;
            }
        }
        else if (op === scfn.commands.fraction_placeholder) {
            // Add fraction part of number
            if (fractionpos >= fractionvalue.length) {
                if (operandstr === "0" || operandstr === "?") {
                    result += operandstr === "0" ? "0" : "&nbsp;";
                }
            } else {
                result += fractionvalue.charAt(fractionpos);
            }
            fractionpos++;
        }
        else if (op === scfn.commands.decimal) {
            // Add decimal point
            if (negativevalue) {
                result += "-";
                negativevalue = 0;
            }
            result += decimalchar;
        }
        else if (op === scfn.commands.currency) {
            // Add currency symbol
            if (negativevalue) {
                result += "-";
                negativevalue = 0;
            }
            result += operandstr;
        }
        else if (op === scfn.commands.general) {
            // Insert "General" conversion
            
            // Cut down number of significant digits to avoid floating point artifacts
            if (value !== 0) { // only if non-zero
                const factor = Math.floor(Math.LOG10E * Math.log(value)); // get integer magnitude as power of 10
                const scaleFactor = Math.pow(10, 13 - factor); // turn into scaling factor
                value = Math.floor(scaleFactor * value + 0.5) / scaleFactor; // scale, round, undo scaling
                if (!isFinite(value)) return "NaN";
            }
            
            if (negativevalue) {
                result += "-";
            }
            
            strvalue = value + ""; // convert original value to string
            if (strvalue.indexOf("e") >= 0) { // converted to scientific notation
                result += strvalue;
                continue;
            }
            
            // Parse integer and fraction parts for general formatting
            strparts = strvalue.match(/^\+{0,1}(\d*)(?:\.(\d*)){0,1}$/);
            integervalue = strparts[1];
            if (!integervalue || integervalue === "0") integervalue = "";
            fractionvalue = strparts[2];
            if (!fractionvalue) fractionvalue = "";
            
            integerpos = 0;
            fractionpos = 0;
            
            // Format integer part
            if (integervalue.length) {
                for (; integerpos < integervalue.length; integerpos++) {
                    result += integervalue.charAt(integerpos);
                    if (sectioninfo.thousandssep) { // check separator position
                        fromend = integervalue.length - integerpos - 1;
                        if (fromend > 2 && fromend % 3 === 0) {
                            result += separatorchar;
                        }
                    }
                }
            } else {
                result += "0";
            }
            
            // Format fraction part
            if (fractionvalue.length) {
                result += decimalchar;
                for (; fractionpos < fractionvalue.length; fractionpos++) {
                    result += fractionvalue.charAt(fractionpos);
                }
            }
        }
        else if (op === scfn.commands.date) {
            // Handle date placeholders
            operandstrlc = operandstr.toLowerCase();
            
            if (operandstrlc === "y" || operandstrlc === "yy") {
                result += (ymd.year + "").substring(2);
            }
            else if (operandstrlc === "yyyy") {
                result += ymd.year + "";
            }
            else if (operandstrlc === "d") {
                result += ymd.day + "";
            }
            else if (operandstrlc === "dd") {
                cval = 1000 + ymd.day;
                result += (cval + "").substr(2);
            }
            else if (operandstrlc === "ddd") {
                cval = Math.floor(rawvalue + 6) % 7;
                result += scc.s_FormatNumber_daynames3[cval];
            }
            else if (operandstrlc === "dddd") {
                cval = Math.floor(rawvalue + 6) % 7;
                result += scc.s_FormatNumber_daynames[cval];
            }
            else if (operandstrlc === "m") {
                result += ymd.month + "";
            }
            else if (operandstrlc === "mm") {
                cval = 1000 + ymd.month;
                result += (cval + "").substr(2);
            }
            else if (operandstrlc === "mmm") {
                result += scc.s_FormatNumber_monthnames3[ymd.month - 1];
            }
            else if (operandstrlc === "mmmm") {
                result += scc.s_FormatNumber_monthnames[ymd.month - 1];
            }
            else if (operandstrlc === "mmmmm") {
                result += scc.s_FormatNumber_monthnames[ymd.month - 1].charAt(0);
            }
            else if (operandstrlc === "h") {
                result += hrs + "";
            }
            else if (operandstrlc === "h]") {
                result += ehrs + "";
            }
            else if (operandstrlc === "mmin") {
                cval = (1000 + mins) + "";
                result += cval.substr(2);
            }
            else if (operandstrlc === "mm]") {
                if (emins < 100) {
                    cval = (1000 + emins) + "";
                    result += cval.substr(2);
                } else {
                    result += emins + "";
                }
            }
            else if (operandstrlc === "min") {
                result += mins + "";
            }
            else if (operandstrlc === "m]") {
                result += emins + "";
            }
            else if (operandstrlc === "hh") {
                cval = (1000 + hrs) + "";
                result += cval.substr(2);
            }
            else if (operandstrlc === "s") {
                cval = Math.floor(secs);
                result += cval + "";
            }
            else if (operandstrlc === "ss") {
                cval = (1000 + Math.floor(secs)) + "";
                result += cval.substr(2);
            }
            else if (operandstrlc === "am/pm" || operandstrlc === "a/p") {
                result += ampmstr;
            }
            else if (operandstrlc === "ss]") {
                if (esecs < 100) {
                    cval = (1000 + Math.floor(esecs)) + "";
                    result += cval.substr(2);
                } else {
                    cval = Math.floor(esecs);
                    result += cval + "";
                }
            }
        }
        else if (op === scfn.commands.section) {
            // End of section
            break;
        }
        else if (op === scfn.commands.comparison) {
            // Ignore comparison operators during formatting
            continue;
        }
        else {
            result += "!! Parse error !!";
        }
    }

    // Apply text color formatting
    if (textcolor) {
        result = `<span style="color:${textcolor};">${result}</span>`;
    }

    // Apply text style formatting  
    if (textstyle) {
        result = `<span style="${textstyle};">${result}</span>`;
    }

    return result;
};
/**
 * Parse a format string and fill in format_defs with the parsed information
 * 
 * Takes a format string (e.g., "#,##0.00_);(#,##0.00)") and creates a structured
 * representation for formatting operations.
 * 
 * @param {Object<string, Object>} format_defs - Hash to store parsed format definitions
 * @param {string} format_string - The format string to parse
 * 
 * @description Format definitions structure:
 * format_defs["#,##0.0"] = {
 *   operators: [],     // array of operators from parsing (each a number)
 *   operands: [],      // array of corresponding operands (each usually a string)  
 *   sectioninfo: [],   // one hash for each section of the format
 *   hascomparison: boolean // true if any section has [<100], etc.
 * }
 * 
 * Each sectioninfo element contains:
 * - start: starting position
 * - integerdigits: number of integer placeholders
 * - fractiondigits: number of fraction placeholders  
 * - commas: comma count for scaling
 * - percent: percent scaling factor
 * - thousandssep: whether thousands separator is used
 * - hasdate: whether section contains date placeholders
 * 
 * @memberof SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber.parse_format_string = function(format_defs, format_string) {
    const scfn = SocialCalc.FormatNumber;

    // Return early if format already parsed
    if (format_defs[format_string]) return;

    // Initialize format structure
    let thisformat, section, sectioninfo;
    let integerpart = 1; // start out in integer part
    let lastwasinteger; // last char was an integer placeholder
    let lastwasslash; // last char was a backslash - escaping following character
    let lastwasasterisk; // repeat next char
    let lastwasunderscore; // last char was _ which picks up following char for width
    let inquote, quotestr; // processing a quoted string
    let inbracket, bracketstr, bracketdata; // processing a bracketed string
    let ingeneral, gpos; // checks for characters "General"
    let ampmstr, part; // checks for characters "A/P" and "AM/PM"
    let indate; // keeps track of date/time placeholders
    let chpos; // character position being looked at
    let ch; // character being looked at

    // Create info structure for this format
    thisformat = {
        operators: [], 
        operands: [], 
        sectioninfo: [{}]
    };
    format_defs[format_string] = thisformat; // add to other format definitions

    // Initialize first section
    section = 0;
    sectioninfo = thisformat.sectioninfo[section]; // get reference to info for current section
    sectioninfo.sectionstart = 0; // position in operands that starts this section
    sectioninfo.integerdigits = 0; // number of integer-part placeholders
    sectioninfo.fractiondigits = 0; // fraction placeholders
    sectioninfo.commas = 0; // commas encountered, to handle scaling
    sectioninfo.percent = 0; // times to scale by 100

    // Parse format string character by character
    for (chpos = 0; chpos < format_string.length; chpos++) {
        ch = format_string.charAt(chpos); // get next char to examine
        
        // Handle quoted strings
        if (inquote) {
            if (ch === '"') {
                inquote = 0;
                thisformat.operators.push(scfn.commands.copy);
                thisformat.operands.push(quotestr);
                continue;
            }
            quotestr += ch;
            continue;
        }
        
        // Handle bracketed expressions
        if (inbracket) {
            if (ch === ']') {
                inbracket = 0;
                bracketdata = SocialCalc.FormatNumber.parse_format_bracket(bracketstr);
                
                if (bracketdata.operator === scfn.commands.separator) {
                    sectioninfo.thousandssep = 1; // explicit [,]
                    continue;
                }
                if (bracketdata.operator === scfn.commands.date) {
                    sectioninfo.hasdate = 1;
                }
                if (bracketdata.operator === scfn.commands.comparison) {
                    thisformat.hascomparison = 1;
                }
                
                thisformat.operators.push(bracketdata.operator);
                thisformat.operands.push(bracketdata.operand);
                continue;
            }
            bracketstr += ch;
            continue;
        }
        
        // Handle escaped characters
        if (lastwasslash) {
            thisformat.operators.push(scfn.commands.copy);
            thisformat.operands.push(ch);
            lastwasslash = false;
            continue;
        }
        
        // Handle repeated characters (asterisk)
        if (lastwasasterisk) {
            thisformat.operators.push(scfn.commands.copy);
            thisformat.operands.push(ch.repeat(5)); // do 5 of them since no real tabs
            lastwasasterisk = false;
            continue;
        }
        
        // Handle underscore spacing
        if (lastwasunderscore) {
            thisformat.operators.push(scfn.commands.copy);
            thisformat.operands.push("&nbsp;");
            lastwasunderscore = false;
            continue;
        }
        
        // Handle "General" keyword detection
        if (ingeneral) {
            if ("general".charAt(ingeneral) === ch.toLowerCase()) {
                ingeneral++;
                if (ingeneral === 7) {
                    thisformat.operators.push(scfn.commands.general);
                    thisformat.operands.push(ch);
                    ingeneral = 0;
                }
                continue;
            }
            ingeneral = 0;
        }
        
        // Handle date/time placeholders
        if (indate) { // last char was part of a date placeholder
            if (indate.charAt(0) === ch) { // another of the same char
                indate += ch; // accumulate it
                continue;
            }
            thisformat.operators.push(scfn.commands.date); // something else, save date info
            thisformat.operands.push(indate);
            sectioninfo.hasdate = 1;
            indate = "";
        }
        
        // Handle AM/PM detection
        if (ampmstr) {
            ampmstr += ch;
            part = ampmstr.toLowerCase();
            
            if (part !== "am/pm".substring(0, part.length) && 
                part !== "a/p".substring(0, part.length)) {
                ampmstr = "";
            }
            else if (part === "am/pm" || part === "a/p") {
                thisformat.operators.push(scfn.commands.date);
                thisformat.operands.push(ampmstr);
                ampmstr = "";
            }
            continue;
        }
        
        // Handle number placeholders
        if (ch === "#" || ch === "0" || ch === "?") {
            if (integerpart) {
                sectioninfo.integerdigits++;
                if (sectioninfo.commas) { // comma inside of integer placeholders
                    sectioninfo.thousandssep = 1; // any number is thousands separator
                    sectioninfo.commas = 0; // reset count of "thousand" factors
                }
                lastwasinteger = 1;
                thisformat.operators.push(scfn.commands.integer_placeholder);
                thisformat.operands.push(ch);
            }
            else {
                sectioninfo.fractiondigits++;
                thisformat.operators.push(scfn.commands.fraction_placeholder);
                thisformat.operands.push(ch);
            }
        }
        // Handle decimal point
        else if (ch === ".") {
            lastwasinteger = 0;
            thisformat.operators.push(scfn.commands.decimal);
            thisformat.operands.push(ch);
            integerpart = 0;
        }
        // Handle currency symbol
        else if (ch === '$') {
            lastwasinteger = 0;
            thisformat.operators.push(scfn.commands.currency);
            thisformat.operands.push(ch);
        }
        // Handle comma separator/scaling
        else if (ch === ",") {
            if (lastwasinteger) {
                sectioninfo.commas++;
            }
            else {
                thisformat.operators.push(scfn.commands.copy);
                thisformat.operands.push(ch);
            }
        }
        // Handle percent symbol
        else if (ch === "%") {
            lastwasinteger = 0;
            sectioninfo.percent++;
            thisformat.operators.push(scfn.commands.copy);
            thisformat.operands.push(ch);
        }
        // Handle quoted string start
        else if (ch === '"') {
            lastwasinteger = 0;
            inquote = 1;
            quotestr = "";
        }
        // Handle bracket start
        else if (ch === '[') {
            lastwasinteger = 0;
            inbracket = 1;
            bracketstr = "";
        }
        // Handle escape character
        else if (ch === '\\') {
            lastwasslash = 1;
            lastwasinteger = 0;
        }
        // Handle repeat character
        else if (ch === '*') {
            lastwasasterisk = 1;
            lastwasinteger = 0;
        }
        // Handle underscore spacing
        else if (ch === '_') {
            lastwasunderscore = 1;
            lastwasinteger = 0;
        }
        // Handle section separator
        else if (ch === ";") {
            section++; // start next section
            thisformat.sectioninfo[section] = {}; // create a new section
            sectioninfo = thisformat.sectioninfo[section]; // get reference to info for current section
            sectioninfo.sectionstart = 1 + thisformat.operators.length; // remember where it starts
            sectioninfo.integerdigits = 0; // number of integer-part placeholders
            sectioninfo.fractiondigits = 0; // fraction placeholders
            sectioninfo.commas = 0; // commas encountered, to handle scaling
            sectioninfo.percent = 0; // times to scale by 100
            integerpart = 1; // reset for new section
            lastwasinteger = 0;
            thisformat.operators.push(scfn.commands.section);
            thisformat.operands.push(ch);
        }
        // Handle "General" keyword start
        else if (ch.toLowerCase() === "g") {
            ingeneral = 1;
            lastwasinteger = 0;
        }
        // Handle AM/PM start
        else if (ch.toLowerCase() === "a") {
            ampmstr = ch;
            lastwasinteger = 0;
        }
        // Handle date/time format characters
        else if ("dmyhHs".indexOf(ch) >= 0) {
            indate = ch;
        }
        // Handle all other characters as literal copies
        else {
            lastwasinteger = 0;
            thisformat.operators.push(scfn.commands.copy);
            thisformat.operands.push(ch);
        }
    }

    // Handle any remaining unsaved date placeholder
    if (indate) {
        thisformat.operators.push(scfn.commands.date);
        thisformat.operands.push(indate);
        sectioninfo.hasdate = 1;
    }

    return;
};
/**
 * Parse bracket contents and return operator and operand information
 * 
 * Takes bracket contents (e.g., "RED", ">10") and returns structured data
 * for formatting operations.
 * 
 * @param {string} bracketstr - The contents inside the brackets (without the brackets)
 * @returns {Object} bracketdata - Object containing operator and operand
 * @returns {number} bracketdata.operator - The command operator number
 * @returns {string} bracketdata.operand - The operand string for the operator
 * 
 * @example
 * // Color bracket
 * parse_format_bracket("RED") // returns {operator: 2, operand: "#FF0000"}
 * 
 * // Comparison bracket  
 * parse_format_bracket(">10") // returns {operator: 10, operand: ">:10"}
 * 
 * // Currency bracket
 * parse_format_bracket("$USD") // returns {operator: 6, operand: "USD"}
 * 
 * @memberof SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber.parse_format_bracket = function(bracketstr) {
    const scfn = SocialCalc.FormatNumber;
    const scc = SocialCalc.Constants;

    const bracketdata = {};
    let parts;

    // Handle currency brackets [$...]
    if (bracketstr.charAt(0) === '$') {
        bracketdata.operator = scfn.commands.currency;
        parts = bracketstr.match(/^\$(.+?)(\-.+?){0,1}$/);
        
        if (parts) {
            bracketdata.operand = parts[1] || scc.FormatNumber_defaultCurrency || '$';
        } else {
            bracketdata.operand = bracketstr.substring(1) || scc.FormatNumber_defaultCurrency || '$';
        }
    }
    // Handle special currency placeholder [?$]
    else if (bracketstr === '?$') {
        bracketdata.operator = scfn.commands.currency;
        bracketdata.operand = '[?$]';
    }
    // Handle color brackets (RED, BLUE, etc.)
    else if (scfn.allowedcolors[bracketstr.toUpperCase()]) {
        bracketdata.operator = scfn.commands.color;
        bracketdata.operand = scfn.allowedcolors[bracketstr.toUpperCase()];
    }
    // Handle style brackets [style=...]
    else if (parts = bracketstr.match(/^style=([^"]*)$/)) {
        bracketdata.operator = scfn.commands.style;
        bracketdata.operand = parts[1];
    }
    // Handle separator bracket [,]
    else if (bracketstr === ",") {
        bracketdata.operator = scfn.commands.separator;
        bracketdata.operand = bracketstr;
    }
    // Handle allowed date format brackets ([h], [m], [s], etc.)
    else if (scfn.alloweddates[bracketstr.toUpperCase()]) {
        bracketdata.operator = scfn.commands.date;
        bracketdata.operand = scfn.alloweddates[bracketstr.toUpperCase()];
    }
    // Handle comparison operators [<100], [>=50], etc.
    else if (parts = bracketstr.match(/^[<>=]/)) {
        parts = bracketstr.match(/^([<>=]+)(.+)$/); // split operator and value
        bracketdata.operator = scfn.commands.comparison;
        bracketdata.operand = parts[1] + ":" + parts[2];
    }
    // Handle unknown brackets - treat as literal text
    else {
        bracketdata.operator = scfn.commands.copy;
        bracketdata.operand = "[" + bracketstr + "]";
    }

    return bracketdata;
};

/**
 * Convert Gregorian date to Julian date number
 * 
 * Uses the algorithm from: http://aa.usno.navy.mil/faq/docs/JD_Formula.html
 * Based on: Fliegel, H. F. and van Flandern, T. C. (1968). Communications of the ACM, 
 * Vol. 11, No. 10 (October, 1968). Translated from the FORTRAN.
 * 
 * @param {number} year - The year (e.g., 2023)
 * @param {number} month - The month (1-12, where 1 = January)
 * @param {number} day - The day of the month (1-31)
 * @returns {number} juliandate - The Julian day number
 * 
 * @example
 * // Convert January 1, 2000 to Julian date
 * convert_date_gregorian_to_julian(2000, 1, 1) // returns 2451545
 * 
 * @memberof SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber.convert_date_gregorian_to_julian = function(year, month, day) {
    let juliandate;

    // Apply the Fliegel and van Flandern algorithm
    juliandate = day - 32075 + SocialCalc.intFunc(1461 * (year + 4800 + SocialCalc.intFunc((month - 14) / 12)) / 4);
    juliandate += SocialCalc.intFunc(367 * (month - 2 - SocialCalc.intFunc((month - 14) / 12) * 12) / 12);
    juliandate = juliandate - SocialCalc.intFunc(3 * SocialCalc.intFunc((year + 4900 + SocialCalc.intFunc((month - 14) / 12)) / 100) / 4);

    return juliandate;
};

/**
 * Convert Julian date number to Gregorian date
 * 
 * Uses the algorithm from: http://aa.usno.navy.mil/faq/docs/JD_Formula.html  
 * Based on: Fliegel, H. F. and van Flandern, T. C. (1968). Communications of the ACM,
 * Vol. 11, No. 10 (October, 1968). Translated from the FORTRAN.
 * 
 * @param {number} juliandate - The Julian day number
 * @returns {Object} ymd - Object containing year, month, and day
 * @returns {number} ymd.year - The year
 * @returns {number} ymd.month - The month (1-12, where 1 = January)
 * @returns {number} ymd.day - The day of the month
 * 
 * @example
 * // Convert Julian date to January 1, 2000
 * convert_date_julian_to_gregorian(2451545) // returns {year: 2000, month: 1, day: 1}
 * 
 * @memberof SocialCalc.FormatNumber
 */
SocialCalc.FormatNumber.convert_date_julian_to_gregorian = function(juliandate) {
    let L, N, I, J, K;

    // Apply the Fliegel and van Flandern reverse algorithm
    L = juliandate + 68569;
    N = Math.floor(4 * L / 146097);
    L = L - Math.floor((146097 * N + 3) / 4);
    I = Math.floor(4000 * (L + 1) / 1461001);
    L = L - Math.floor(1461 * I / 4) + 31;
    J = Math.floor(80 * L / 2447);
    K = L - Math.floor(2447 * J / 80);
    L = Math.floor(J / 11);
    J = J + 2 - 12 * L;
    I = 100 * (N - 49) + I + L;

    return {
        year: I, 
        month: J, 
        day: K
    };
};

/**
 * Integer function that handles negative numbers correctly
 * 
 * This function provides proper integer truncation for both positive and negative numbers,
 * which is needed for the Julian date conversion algorithms.
 * 
 * @param {number} n - The number to truncate to an integer
 * @returns {number} The truncated integer value
 * 
 * @example
 * intFunc(3.7)   // returns 3
 * intFunc(-3.7)  // returns -3 (not -4 like Math.floor would)
 * intFunc(0)     // returns 0
 * 
 * @memberof SocialCalc
 */
SocialCalc.intFunc = function(n) {
    if (n < 0) {
        return -Math.floor(-n);
    } else {
        return Math.floor(n);
    }
};

