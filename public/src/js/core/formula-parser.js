/**
 * @fileoverview SocialCalc Spreadsheet Formula Library
 * Part of the SocialCalc package
 * 
 * @copyright (c) Copyright 2008 Socialtext, Inc. All Rights Reserved.
 * @license Artistic License 2.0 - http://socialcalc.org/licenses/al-20/
 * 
 * @description This library handles formula parsing and processing for SocialCalc.
 * Some of the other files in the SocialCalc package are licensed under different licenses.
 * Please note the licenses of the modules you use.
 * 
 * Code History:
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

if (typeof SocialCalc === 'undefined') {
    var SocialCalc = {}; // May be used with other SocialCalc libraries or standalone
}                        // In any case, requires SocialCalc.Constants.

/**
 * @namespace SocialCalc.Formula
 * @description Main formula processing namespace containing parsing and evaluation functions
 */
SocialCalc.Formula = {};

/**
 * @enum {number}
 * @description Formula parsing state constants for the state machine
 */
SocialCalc.Formula.ParseState = {
   /** @type {number} Parsing numeric values */
   num: 1,
   /** @type {number} Parsing alphabetic characters */
   alpha: 2,
   /** @type {number} Parsing coordinate references */
   coord: 3,
   /** @type {number} Parsing string literals */
   string: 4,
   /** @type {number} Parsing quoted strings */
   stringquote: 5,
   /** @type {number} Parsing numeric exponential part 1 */
   numexp1: 6,
   /** @type {number} Parsing numeric exponential part 2 */
   numexp2: 7,
   /** @type {number} Parsing alphanumeric sequences */
   alphanumeric: 8,
   /** @type {number} Parsing special values */
   specialvalue: 9
};

/**
 * @enum {number}
 * @description Token type constants for parsed formula elements
 */
SocialCalc.Formula.TokenType = {
   /** @type {number} Numeric token */
   num: 1,
   /** @type {number} Coordinate reference token */
   coord: 2,
   /** @type {number} Operator token */
   op: 3,
   /** @type {number} Name/function token */
   name: 4,
   /** @type {number} Error token */
   error: 5,
   /** @type {number} String token */
   string: 6,
   /** @type {number} Space token */
   space: 7
};

/**
 * @enum {number}
 * @description Character classification constants for parsing
 */
SocialCalc.Formula.CharClass = {
   /** @type {number} Numeric character */
   num: 1,
   /** @type {number} Character that can start a number */
   numstart: 2,
   /** @type {number} Operator character */
   op: 3,
   /** @type {number} End of file */
   eof: 4,
   /** @type {number} Alphabetic character */
   alpha: 5,
   /** @type {number} Character valid in coordinates */
   incoord: 6,
   /** @type {number} Error character */
   error: 7,
   /** @type {number} Quote character */
   quote: 8,
   /** @type {number} Space character */
   space: 9,
   /** @type {number} Special value start character */
   specialstart: 10
};

/**
 * @type {Object<string, number>}
 * @description Lookup table mapping characters to their classification types
 */
SocialCalc.Formula.CharClassTable = {
   " ": 9, "!": 3, '"': 8, "#": 10, "$":6, "%":3, "&":3, "(": 3, ")": 3, "*": 3, "+": 3, ",": 3, "-": 3, ".": 2, "/": 3,
   "0": 1, "1": 1, "2": 1, "3": 1, "4": 1, "5": 1, "6": 1, "7": 1, "8": 1, "9": 1,
   ":": 3, "<": 3, "=": 3, ">": 3,
   "A": 5, "B": 5, "C": 5, "D": 5, "E": 5, "F": 5, "G": 5, "H": 5, "I": 5, "J": 5, "K": 5, "L": 5, "M": 5, "N": 5,
   "O": 5, "P": 5, "Q": 5, "R": 5, "S": 5, "T": 5, "U": 5, "V": 5, "W": 5, "X": 5, "Y": 5, "Z": 5,
   "^": 3, "_": 5,
   "a": 5, "b": 5, "c": 5, "d": 5, "e": 5, "f": 5, "g": 5, "h": 5, "i": 5, "j": 5, "k": 5, "l": 5, "m": 5, "n": 5,
   "o": 5, "p": 5, "q": 5, "r": 5, "s": 5, "t": 5, "u": 5, "v": 5, "w": 5, "x": 5, "y": 5, "z": 5
};

/**
 * @type {Object<string, string>}
 * @description Lookup table for converting lowercase letters to uppercase
 */
SocialCalc.Formula.UpperCaseTable = {
   "a": "A", "b": "B", "c": "C", "d": "D", "e": "E", "f": "F", "g": "G", "h": "H", "i": "I", "j": "J", "k": "K", "l": "L", "m": "M",
   "n": "N", "o": "O", "p": "P", "q": "Q", "r": "R", "s": "S", "t": "T", "u": "U", "v": "V", "w": "W", "x": "X", "y": "Y", "z": "Z"
};

/**
 * @type {Object<string, string>}
 * @description Special constant names that convert to error values during name lookup
 */
SocialCalc.Formula.SpecialConstants = {
   "#NULL!": "0,e#NULL!",
   "#NUM!": "0,e#NUM!",
   "#DIV/0!": "0,e#DIV/0!",
   "#VALUE!": "0,e#VALUE!",
   "#REF!": "0,e#REF!",
   "#NAME?": "0,e#NAME?"
};

/**
 * @type {Object<string, number>}
 * @description Operator precedence table for formula evaluation
 * Order: 1- !, 2- : ,, 3- M P, 4- %, 5- ^, 6- * /, 7- + -, 8- &, 9- < > = G(>=) L(<=) N(<>)
 * Negative values indicate right associative operators
 */
SocialCalc.Formula.TokenPrecedence = {
   "!": 1,
   ":": 2, ",": 2,
   "M": -3, "P": -3,
   "%": 4,
   "^": 5,
   "*": 6, "/": 6,
   "+": 7, "-": 7,
   "&": 8,
   "<": 9, ">": 9, "G": 9, "L": 9, "N": 9
};

/**
 * @type {Object<string, string>}
 * @description Converts single-character token text to input text representation
 */
SocialCalc.Formula.TokenOpExpansion = {
   'G': '>=',
   'L': '<=',
   'M': '-',
   'N': '<>',
   'P': '+'
};

/**
 * @type {Object<string, Object<string, string>>}
 * @description Type lookup table for determining result types when performing operations on values
 * 
 * Each entry contains type mappings with format:
 * 'type1a': '|type2a:resulta|type2b:resultb|...'
 * 
 * Type patterns:
 * - t* or n* matches any of those types not specifically listed
 * - Results may be a type or numbers 1/2 specifying to return type1 or type2
 */
SocialCalc.Formula.TypeLookupTable = {
   unaryminus: { 'n*': '|n*:1|', 'e*': '|e*:1|', 't*': '|t*:e#VALUE!|', 'b': '|b:n|'},
   unaryplus: { 'n*': '|n*:1|', 'e*': '|e*:1|', 't*': '|t*:e#VALUE!|', 'b': '|b:n|'},
   unarypercent: { 'n*': '|n:n%|n*:n|', 'e*': '|e*:1|', 't*': '|t*:e#VALUE!|', 'b': '|b:n|'},
   plus: {
      'n%': '|n%:n%|nd:n|nt:n|ndt:n|n$:n|n:n|n*:n|b:n|e*:2|t*:e#VALUE!|',
      'nd': '|n%:n|nd:nd|nt:ndt|ndt:ndt|n$:n|n:nd|n*:n|b:n|e*:2|t*:e#VALUE!|',
      'nt': '|n%:n|nd:ndt|nt:nt|ndt:ndt|n$:n|n:nt|n*:n|b:n|e*:2|t*:e#VALUE!|',
      'ndt': '|n%:n|nd:ndt|nt:ndt|ndt:ndt|n$:n|n:ndt|n*:n|b:n|e*:2|t*:e#VALUE!|',
      'n$': '|n%:n|nd:n|nt:n|ndt:n|n$:n$|n:n$|n*:n|b:n|e*:2|t*:e#VALUE!|',
      'nl': '|n%:n|nd:n|nt:n|ndt:n|n$:n|n:n|n*:n|b:n|e*:2|t*:e#VALUE!|',
      'n': '|n%:n|nd:nd|nt:nt|ndt:ndt|n$:n$|n:n|n*:n|b:n|e*:2|t*:e#VALUE!|',
      'b': '|n%:n%|nd:nd|nt:nt|ndt:ndt|n$:n$|n:n|n*:n|b:n|e*:2|t*:e#VALUE!|',
      't*': '|n*:e#VALUE!|t*:e#VALUE!|b:e#VALUE!|e*:2|',
      'e*': '|e*:1|n*:1|t*:1|b:1|'
   },
   concat: {
      't': '|t:t|th:th|tw:tw|tl:t|t*:2|e*:2|',
      'th': '|t:th|th:th|tw:t|tl:th|t*:t|e*:2|',
      'tw': '|t:tw|th:t|tw:tw|tl:tw|t*:t|e*:2|',
      'tl': '|t:tl|th:th|tw:tw|tl:tl|t*:t|e*:2|',
      'e*': '|e*:1|n*:1|t*:1|'
   },
   oneargnumeric: { 'n*': '|n*:n|', 'e*': '|e*:1|', 't*': '|t*:e#VALUE!|', 'b': '|b:n|'},
   twoargnumeric: { 'n*': '|n*:n|t*:e#VALUE!|e*:2|', 'e*': '|e*:1|n*:1|t*:1|', 't*': '|t*:e#VALUE!|n*:e#VALUE!|e*:2|'},
   propagateerror: { 'n*': '|n*:2|e*:2|', 'e*': '|e*:2|', 't*': '|t*:2|e*:2|', 'b': '|b:2|e*:2|'}
};

/**
 * @function ParseFormulaIntoTokens
 * @memberof SocialCalc.Formula
 * @description Parses a text string as if it was a spreadsheet formula using a state machine
 * 
 * @param {string} line - The formula text to parse
 * @returns {Array<Object>} parseinfo - Array of parsed tokens, each containing:
 *   - {string} text - The characters making up the parsed token
 *   - {number} type - The type of the token (from TokenType enum)
 *   - {string} opcode - Single character version of operator for precedence table
 * 
 * @example
 * const tokens = SocialCalc.Formula.ParseFormulaIntoTokens("=A1+B2*3");
 * // Returns array of token objects representing the parsed formula
 */
SocialCalc.Formula.ParseFormulaIntoTokens = function(line) {
   let i, ch, chclass, haddecimal, last_token, last_token_type, last_token_text, t;

   const scf = SocialCalc.Formula;
   const scc = SocialCalc.Constants;
   const parsestate = scf.ParseState;
   const tokentype = scf.TokenType;
   const charclass = scf.CharClass;
   const charclasstable = scf.CharClassTable;
   const uppercasetable = scf.UpperCaseTable; // much faster than toUpperCase function
   const pushtoken = scf.ParsePushToken;
   const coordregex = /^\$?[A-Z]{1,2}\$?[1-9]\d*$/i;

   const parseinfo = [];
   let str = "";
   let state = 0;
   haddecimal = false;

   for (i = 0; i <= line.length; i++) {
      if (i < line.length) {
         ch = line.charAt(i);
         cclass = charclasstable[ch];
      } else {
         ch = "";
         cclass = charclass.eof;
      }

      if (state == parsestate.num) {
         if (cclass == charclass.num) {
            str += ch;
         } else if (cclass == charclass.numstart && !haddecimal) {
            haddecimal = true;
            str += ch;
         } else if (ch == "E" || ch == "e") {
            str += ch;
            haddecimal = false;
            state = parsestate.numexp1;
         } else { // end of number - save it
            pushtoken(parseinfo, str, tokentype.num, 0);
            haddecimal = false;
            state = 0;
         }
      }

      if (state == parsestate.numexp1) {
         if (cclass == parsestate.num) {
            state = parsestate.numexp2;
         } else if ((ch == '+' || ch == '-') && (uppercasetable[str.charAt(str.length-1)] == 'E')) {
            str += ch;
         } else if (ch == 'E' || ch == 'e') {
            // Continue parsing
         } else {
            pushtoken(parseinfo, scc.s_parseerrexponent, tokentype.error, 0);
            state = 0;
         }
      }

      if (state == parsestate.numexp2) {
         if (cclass == charclass.num) {
            str += ch;
         } else { // end of number - save it
            pushtoken(parseinfo, str, tokentype.num, 0);
            state = 0;
         }
      }

      if (state == parsestate.alpha) {
         if (cclass == charclass.num) {
            state = parsestate.coord;
         } else if (cclass == charclass.alpha || ch == ".") { // alpha may be letters, numbers, "_", or "."
            str += ch;
         } else if (cclass == charclass.incoord) {
            state = parsestate.coord;
         } else if (cclass == charclass.op || cclass == charclass.numstart
                || cclass == charclass.space || cclass == charclass.eof) {
            pushtoken(parseinfo, str.toUpperCase(), tokentype.name, 0);
            state = 0;
         } else {
            pushtoken(parseinfo, scc.s_parseerrchar, tokentype.error, 0);
            state = 0;
         }
      }

      if (state == parsestate.coord) {
         if (cclass == charclass.num) {
            str += ch;
         } else if (cclass == charclass.incoord) {
            str += ch;
         } else if (cclass == charclass.alpha) {
            state = parsestate.alphanumeric;
         } else if (cclass == charclass.op || cclass == charclass.numstart ||
                  cclass == charclass.eof || cclass == charclass.space) {
            if (coordregex.test(str)) {
               t = tokentype.coord;
            } else {
               t = tokentype.name;
            }
            pushtoken(parseinfo, str.toUpperCase(), t, 0);
            state = 0;
         } else {
            pushtoken(parseinfo, scc.s_parseerrchar, tokentype.error, 0);
            state = 0;
         }
      }
      if (state == parsestate.alphanumeric) {
         if (cclass == charclass.num || cclass == charclass.alpha) {
            str += ch;
         } else if (cclass == charclass.op || cclass == charclass.numstart
                || cclass == charclass.space || cclass == charclass.eof) {
            pushtoken(parseinfo, str.toUpperCase(), tokentype.name, 0);
            state = 0;
         } else {
            pushtoken(parseinfo, scc.s_parseerrchar, tokentype.error, 0);
            state = 0;
         }
      }

      if (state == parsestate.string) {
         if (cclass == charclass.quote) {
            state = parsestate.stringquote; // got quote in string: is it doubled (quote in string) or by itself (end of string)?
         } else if (cclass == charclass.eof) {
            pushtoken(parseinfo, scc.s_parseerrstring, tokentype.error, 0);
            state = 0;
         } else {
            str += ch;
         }
      } else if (state == parsestate.stringquote) { // note else if here
         if (cclass == charclass.quote) {
            str += '"';
            state = parsestate.string; // double quote: add one then continue getting string
         } else { // something else -- end of string
            pushtoken(parseinfo, str, tokentype.string, 0);
            state = 0; // drop through to process
         }
      } else if (state == parsestate.specialvalue) { // special values like #REF!
         if (str.charAt(str.length-1) == "!") { // done - save value as a name
            pushtoken(parseinfo, str, tokentype.name, 0);
            state = 0; // drop through to process
         } else if (cclass == charclass.eof) {
            pushtoken(parseinfo, scc.s_parseerrspecialvalue, tokentype.error, 0);
            state = 0;
         } else {
            str += ch;
         }
      }

      if (state == 0) {
         if (cclass == charclass.num) {
            str = ch;
            state = parsestate.num;
         } else if (cclass == charclass.numstart) {
            str = ch;
            haddecimal = true;
            state = parsestate.num;
         } else if (cclass == charclass.alpha || cclass == charclass.incoord) {
            str = ch;
            state = parsestate.alpha;
         } else if (cclass == charclass.specialstart) {
            str = ch;
            state = parsestate.specialvalue;
         } else if (cclass == charclass.op) {
            str = ch;
            if (parseinfo.length > 0) {
               last_token = parseinfo[parseinfo.length-1];
               last_token_type = last_token.type;
               last_token_text = last_token.text;
               if (last_token_type == charclass.op) {
                  if (last_token_text == '<' || last_token_text == ">") {
                     str = last_token_text + str;
                     parseinfo.pop();
                     if (parseinfo.length > 0) {
                        last_token = parseinfo[parseinfo.length-1];
                        last_token_type = last_token.type;
                        last_token_text = last_token.text;
                     } else {
                        last_token_type = charclass.eof;
                        last_token_text = "EOF";
                     }
                  }
               }
            } else {
               last_token_type = charclass.eof;
               last_token_text = "EOF";
            }
            t = tokentype.op;
            if ((parseinfo.length == 0)
                || (last_token_type == charclass.op && last_token_text != ')' && last_token_text != '%')) { // Unary operator
               if (str == '-') { // M is unary minus
                  str = "M";
                  ch = "M";
               } else if (str == '+') { // P is unary plus
                  str = "P";
                  ch = "P";
               } else if (str == ')' && last_token_text == '(') { // null arg list OK
                  // Continue processing
               } else if (str != '(') { // binary-op open-paren OK, others no
                  t = tokentype.error;
                  str = scc.s_parseerrtwoops;
               }
            } else if (str.length > 1) {
               if (str == '>=') { // G is >=
                  str = "G";
                  ch = "G";
               } else if (str == '<=') { // L is <=
                  str = "L";
                  ch = "L";
               } else if (str == '<>') { // N is <>
                  str = "N";
                  ch = "N";
               } else {
                  t = tokentype.error;
                  str = scc.s_parseerrtwoops;
               }
            }
            pushtoken(parseinfo, str, t, ch);
            state = 0;
         } else if (cclass == charclass.quote) { // starting a string
            str = "";
            state = parsestate.string;
         } else if (cclass == charclass.space) { // store so can reconstruct spacing
            pushtoken(parseinfo, " ", tokentype.space, 0);
         } else if (cclass == charclass.eof) { // ignore -- needed to have extra loop to close out other things
            // Continue processing
         } else { // unknown class - such as unknown char
            pushtoken(parseinfo, scc.s_parseerrchar, tokentype.error, 0);
         }
      }
   }

   return parseinfo;
};

/**
 * @function ParsePushToken
 * @memberof SocialCalc.Formula
 * @description Helper function to push a token into the parseinfo array
 * 
 * @param {Array<Object>} parseinfo - Array to store parsed tokens
 * @param {string} ttext - The text content of the token
 * @param {number} ttype - The type of the token (from TokenType enum)
 * @param {string|number} topcode - The opcode for operators
 * 
 * @example
 * SocialCalc.Formula.ParsePushToken(parseinfo, "123", TokenType.num, 0);
 */
SocialCalc.Formula.ParsePushToken = function(parseinfo, ttext, ttype, topcode) {
   parseinfo.push({text: ttext, type: ttype, opcode: topcode});
};

/**
 * @function evaluate_parsed_formula
 * @memberof SocialCalc.Formula
 * @description Evaluates a parsed formula and returns the calculated result
 * 
 * @param {Array<Object>} parseinfo - Array of parsed tokens from ParseFormulaIntoTokens
 * @param {Object} sheet - The spreadsheet object containing cell data and context
 * @param {boolean} [allowrangereturn] - If true, allows returning range references (e.g., "A1:A10")
 * @returns {Object} Result object containing:
 *   - {*} value - The calculated value
 *   - {string} type - The type of the result value
 *   - {string} error - Error text if calculation failed
 * 
 * @example
 * const result = SocialCalc.Formula.evaluate_parsed_formula(tokens, sheet, false);
 * // Returns: {value: 42, type: "n", error: ""}
 */
SocialCalc.Formula.evaluate_parsed_formula = function(parseinfo, sheet, allowrangereturn) {
   let result;

   const scf = SocialCalc.Formula;
   const tokentype = scf.TokenType;

   let revpolish;
   const parsestack = [];

   let errortext = "";

   revpolish = scf.ConvertInfixToPolish(parseinfo); // result is either an array or a string with error text

   result = scf.EvaluatePolish(parseinfo, revpolish, sheet, allowrangereturn);

   return result;
};

/**
 * @function ConvertInfixToPolish
 * @memberof SocialCalc.Formula
 * @description Converts infix notation to reverse polish notation for formula evaluation
 * 
 * Based upon the algorithm shown in Wikipedia "Reverse Polish notation" article
 * and then enhanced for additional spreadsheet-specific functionality.
 * 
 * @param {Array<Object>} parseinfo - Array of parsed tokens from ParseFormulaIntoTokens
 * @returns {Array<number>|string} Returns array of token indices in reverse polish order,
 *                                  or error string if conversion fails
 * 
 * @example
 * const rpn = SocialCalc.Formula.ConvertInfixToPolish(tokens);
 * // Returns: [0, 2, 1] for "A1 + B2" (indices to tokens in RPN order)
 */
SocialCalc.Formula.ConvertInfixToPolish = function(parseinfo) {
   const scf = SocialCalc.Formula;
   const scc = SocialCalc.Constants;
   const tokentype = scf.TokenType;
   const token_precedence = scf.TokenPrecedence;

   const revpolish = [];
   const parsestack = [];

   let errortext = "";

   /** @type {number} Marker for function start in reverse polish notation */
   const function_start = -1;

   let i, pii, ttype, ttext, tprecedence, tstackprecedence;

   for (i = 0; i < parseinfo.length; i++) {
      pii = parseinfo[i];
      ttype = pii.type;
      ttext = pii.text;
      
      if (ttype == tokentype.num || ttype == tokentype.coord || ttype == tokentype.string) {
         revpolish.push(i);
      } else if (ttype == tokentype.name) {
         parsestack.push(i);
         revpolish.push(function_start);
      } else if (ttype == tokentype.space) { // ignore
         continue;
      } else if (ttext == ',') {
         while (parsestack.length && parseinfo[parsestack[parsestack.length-1]].text != "(") {
            revpolish.push(parsestack.pop());
         }
         if (parsestack.length == 0) { // no ( -- error
            errortext = scc.s_parseerrmissingopenparen;
            break;
         }
      } else if (ttext == '(') {
         parsestack.push(i);
      } else if (ttext == ')') {
         while (parsestack.length && parseinfo[parsestack[parsestack.length-1]].text != "(") {
            revpolish.push(parsestack.pop());
         }
         if (parsestack.length == 0) { // no ( -- error
            errortext = scc.s_parseerrcloseparennoopen;
            break;
         }
         parsestack.pop();
         if (parsestack.length && parseinfo[parsestack[parsestack.length-1]].type == tokentype.name) {
            revpolish.push(parsestack.pop());
         }
      } else if (ttype == tokentype.op) {
         if (parsestack.length && parseinfo[parsestack[parsestack.length-1]].type == tokentype.name) {
            revpolish.push(parsestack.pop());
         }
         while (parsestack.length && parseinfo[parsestack[parsestack.length-1]].type == tokentype.op
                && parseinfo[parsestack[parsestack.length-1]].text != '(') {
            tprecedence = token_precedence[pii.opcode];
            tstackprecedence = token_precedence[parseinfo[parsestack[parsestack.length-1]].opcode];
            if (tprecedence >= 0 && tprecedence < tstackprecedence) {
               break;
            } else if (tprecedence < 0) {
               tprecedence = -tprecedence;
               if (tstackprecedence < 0) tstackprecedence = -tstackprecedence;
               if (tprecedence <= tstackprecedence) {
                  break;
               }
            }
            revpolish.push(parsestack.pop());
         }
         parsestack.push(i);
      } else if (ttype == tokentype.error) {
         errortext = ttext;
         break;
      } else {
         errortext = "Internal error while processing parsed formula. ";
         break;
      }
   }
   
   while (parsestack.length > 0) {
      if (parseinfo[parsestack[parsestack.length-1]].text == '(') {
         errortext = scc.s_parseerrmissingcloseparen;
         break;
      }
      revpolish.push(parsestack.pop());
   }

   if (errortext) {
      return errortext;
   }

   return revpolish;
};
/**
 * @function EvaluatePolish
 * @memberof SocialCalc.Formula
 * @description Execute reverse polish representation of formula
 * 
 * Operand values are objects in the operand array with a "type" and an optional "value".
 * Type can have these values (many are type and sub-type as two or more letters):
 *    "tw", "th", "t", "n", "nt", "coord", "range", "start", "eErrorType", "b" (blank)
 * The value of a coord is in the form A57 or A57!sheetname
 * The value of a range is coord|coord|number where number starts at 0 and is
 * the offset of the next item to fetch if you are going through the range one by one
 * The number starts as a null string ("A1|B3|")
 * 
 * @param {Array<Object>} parseinfo - Array of parsed tokens from ParseFormulaIntoTokens
 * @param {Array<number>|string} revpolish - Reverse polish notation array or error string
 * @param {Object} sheet - The spreadsheet object containing cell data and context
 * @param {boolean} [allowrangereturn] - If true, allows returning range references
 * @returns {Object} Result object containing:
 *   - {*} value - The calculated value
 *   - {string} type - The type of the result value
 *   - {string} error - Error text if calculation failed
 * 
 * @example
 * const result = SocialCalc.Formula.EvaluatePolish(parseinfo, rpn, sheet, false);
 * // Returns: {value: 42, type: "n", error: ""}
 */
SocialCalc.Formula.EvaluatePolish = function(parseinfo, revpolish, sheet, allowrangereturn) {
   const scf = SocialCalc.Formula;
   const scc = SocialCalc.Constants;
   const tokentype = scf.TokenType;
   const lookup_result_type = scf.LookupResultType;
   const typelookup = scf.TypeLookupTable;
   const operand_as_number = scf.OperandAsNumber;
   const operand_as_text = scf.OperandAsText;
   const operand_value_and_type = scf.OperandValueAndType;
   const operands_as_coord_on_sheet = scf.OperandsAsCoordOnSheet;
   const format_number_for_display = SocialCalc.format_number_for_display || function(v, t, f) {return v+"";};

   let errortext = "";
   /** @type {number} Marker for function start in evaluation stack */
   const function_start = -1;
   /** @type {Object} Standard error object for missing operands */
   const missingOperandError = {value: "", type: "e#VALUE!", error: scc.s_parseerrmissingoperand};

   /** @type {Array<Object>} Stack of operands during evaluation */
   const operand = [];
   
   /**
    * @function PushOperand
    * @description Helper function to push an operand onto the evaluation stack
    * @param {string} t - Type of the operand
    * @param {*} v - Value of the operand
    */
   const PushOperand = function(t, v) {operand.push({type: t, value: v});};

   let i, rii, prii, ttype, ttext, value1, value2, tostype, tostype2, resulttype, valuetype, cond, vmatch, smatch;

   if (!parseinfo.length || (! (revpolish instanceof Array))) {
      return ({value: "", type: "e#VALUE!", error: (typeof revpolish == "string" ? revpolish : "")});
   }

   for (i = 0; i < revpolish.length; i++) {
      rii = revpolish[i];
      if (rii == function_start) { // Remember the start of a function argument list
         PushOperand("start", 0);
         continue;
      }

      prii = parseinfo[rii];
      ttype = prii.type;
      ttext = prii.text;

      if (ttype == tokentype.num) {
         PushOperand("n", ttext-0);
      } else if (ttype == tokentype.coord) {
         PushOperand("coord", ttext);
      } else if (ttype == tokentype.string) {
         PushOperand("t", ttext);
      } else if (ttype == tokentype.op) {
         if (operand.length <= 0) { // Nothing on the stack...
            return missingOperandError;
            break; // done
         }

         // Unary minus
         if (ttext == 'M') {
            value1 = operand_as_number(sheet, operand);
            resulttype = lookup_result_type(value1.type, value1.type, typelookup.unaryminus);
            PushOperand(resulttype, -value1.value);
         }
         // Unary plus
         else if (ttext == 'P') {
            value1 = operand_as_number(sheet, operand);
            resulttype = lookup_result_type(value1.type, value1.type, typelookup.unaryplus);
            PushOperand(resulttype, value1.value);
         }
         // Unary % - percent, left associative
         else if (ttext == '%') {
            value1 = operand_as_number(sheet, operand);
            resulttype = lookup_result_type(value1.type, value1.type, typelookup.unarypercent);
            PushOperand(resulttype, 0.01*value1.value);
         }
         // & - string concatenate
         else if (ttext == '&') {
            if (operand.length <= 1) { // Need at least two things on the stack...
               return missingOperandError;
            }
            value2 = operand_as_text(sheet, operand);
            value1 = operand_as_text(sheet, operand);
            resulttype = lookup_result_type(value1.type, value1.type, typelookup.concat);
            PushOperand(resulttype, value1.value + value2.value);
         }
         // : - Range constructor
         else if (ttext == ':') {
            if (operand.length <= 1) { // Need at least two things on the stack...
               return missingOperandError;
            }
            value1 = scf.OperandsAsRangeOnSheet(sheet, operand); // get coords even if use name on other sheet
            if (value1.error) { // not available
               errortext = errortext || value1.error;
            }
            PushOperand(value1.type, value1.value); // push sheetname with range on that sheet
         }
         // ! - sheetname!coord
         else if (ttext == '!') {
            if (operand.length <= 1) { // Need at least two things on the stack...
               return missingOperandError;
            }
            value1 = operands_as_coord_on_sheet(sheet, operand); // get coord even if name on other sheet
            if (value1.error) { // not available
               errortext = errortext || value1.error;
            }
            PushOperand(value1.type, value1.value); // push sheetname with coord or range on that sheet
         }
         // Comparison operators: < L = G > N (< <= = >= > <>)
         else if (ttext == "<" || ttext == "L" || ttext == "=" || ttext == "G" || ttext == ">" || ttext == "N") {
            if (operand.length <= 1) { // Need at least two things on the stack...
               errortext = scc.s_parseerrmissingoperand; // remember error
               break;
            }
            value2 = operand_value_and_type(sheet, operand);
            value1 = operand_value_and_type(sheet, operand);
            if (value1.type.charAt(0) == "n" && value2.type.charAt(0) == "n") { // compare two numbers
               cond = 0;
               if (ttext == "<") { cond = value1.value < value2.value ? 1 : 0; }
               else if (ttext == "L") { cond = value1.value <= value2.value ? 1 : 0; }
               else if (ttext == "=") { cond = value1.value == value2.value ? 1 : 0; }
               else if (ttext == "G") { cond = value1.value >= value2.value ? 1 : 0; }
               else if (ttext == ">") { cond = value1.value > value2.value ? 1 : 0; }
               else if (ttext == "N") { cond = value1.value != value2.value ? 1 : 0; }
               PushOperand("nl", cond);
            } else if (value1.type.charAt(0) == "e") { // error on left
               PushOperand(value1.type, 0);
            } else if (value2.type.charAt(0) == "e") { // error on right
               PushOperand(value2.type, 0);
            } else { // text maybe mixed with numbers or blank
               tostype = value1.type.charAt(0);
               tostype2 = value2.type.charAt(0);
               if (tostype == "n") {
                  value1.value = format_number_for_display(value1.value, "n", "");
               } else if (tostype == "b") {
                  value1.value = "";
               }
               if (tostype2 == "n") {
                  value2.value = format_number_for_display(value2.value, "n", "");
               } else if (tostype2 == "b") {
                  value2.value = "";
               }
               cond = 0;
               value1.value = value1.value.toLowerCase(); // ignore case
               value2.value = value2.value.toLowerCase();
               if (ttext == "<") { cond = value1.value < value2.value ? 1 : 0; }
               else if (ttext == "L") { cond = value1.value <= value2.value ? 1 : 0; }
               else if (ttext == "=") { cond = value1.value == value2.value ? 1 : 0; }
               else if (ttext == "G") { cond = value1.value >= value2.value ? 1 : 0; }
               else if (ttext == ">") { cond = value1.value > value2.value ? 1 : 0; }
               else if (ttext == "N") { cond = value1.value != value2.value ? 1 : 0; }
               PushOperand("nl", cond);
            }
         }
         // Normal infix arithmetic operators: +, -, *, /, ^
         else { // what's left are the normal infix arithmetic operators
            if (operand.length <= 1) { // Need at least two things on the stack...
               errortext = scc.s_parseerrmissingoperand; // remember error
               break;
            }
            value2 = operand_as_number(sheet, operand);
            value1 = operand_as_number(sheet, operand);
            if (ttext == '+') {
               resulttype = lookup_result_type(value1.type, value2.type, typelookup.plus);
               PushOperand(resulttype, value1.value + value2.value);
            } else if (ttext == '-') {
               resulttype = lookup_result_type(value1.type, value2.type, typelookup.plus);
               PushOperand(resulttype, value1.value - value2.value);
            } else if (ttext == '*') {
               resulttype = lookup_result_type(value1.type, value2.type, typelookup.plus);
               PushOperand(resulttype, value1.value * value2.value);
            } else if (ttext == '/') {
               if (value2.value != 0) {
                  PushOperand("n", value1.value / value2.value); // gives plain numeric result type
               } else {
                  PushOperand("e#DIV/0!", 0);
               }
            } else if (ttext == '^') {
               value1.value = Math.pow(value1.value, value2.value);
               value1.type = "n"; // gives plain numeric result type
               if (isNaN(value1.value)) {
                  value1.value = 0;
                  value1.type = "e#NUM!";
               }
               PushOperand(value1.type, value1.value);
            }
         }
      }
      // function or name
      else if (ttype == tokentype.name) {
         errortext = scf.CalculateFunction(ttext, operand, sheet);
         if (errortext) break;
      } else {
         errortext = scc.s_InternalError+"Unknown token "+ttype+" ("+ttext+"). ";
         break;
      }
   }
   // Look at final value and handle special cases

   value = operand[0] ? operand[0].value : "";
   tostype = operand[0] ? operand[0].type : "";

   if (tostype == "name") { // name - expand it
      value1 = SocialCalc.Formula.LookupName(sheet, value);
      value = value1.value;
      tostype = value1.type;
      errortext = errortext || value1.error;
   }

   if (tostype == "coord") { // the value is a coord reference, get its value and type
      value1 = operand_value_and_type(sheet, operand);
      value = value1.value;
      tostype = value1.type;
      if (tostype == "b") {
         tostype = "n";
         value = 0;
      }
   }

   if (operand.length > 1 && !errortext) { // something left - error
      errortext += scc.s_parseerrerrorinformula;
   }

   // Set return type
   valuetype = tostype;

   if (tostype.charAt(0) == "e") { // error value
      errortext = errortext || tostype.substring(1) || scc.s_calcerrerrorvalueinformula;
   } else if (tostype == "range") {
      vmatch = value.match(/^(.*)\|(.*)\|/);
      smatch = vmatch[1].indexOf("!");
      if (smatch >= 0) { // swap sheetname
         vmatch[1] = vmatch[1].substring(smatch+1) + "!" + vmatch[1].substring(0, smatch).toUpperCase();
      } else {
         vmatch[1] = vmatch[1].toUpperCase();
      }
      value = vmatch[1] + ":" + vmatch[2].toUpperCase();
      if (!allowrangereturn) {
         errortext = scc.s_formularangeresult+" "+value;
      }
   }

   if (errortext && valuetype.charAt(0) != "e") {
      value = errortext;
      valuetype = "e";
   }

   // Look for overflow
   if (valuetype.charAt(0) == "n" && (isNaN(value) || !isFinite(value))) {
      value = 0;
      valuetype = "e#NUM!";
      errortext = isNaN(value) ? scc.s_calcerrnumericnan: scc.s_calcerrnumericoverflow;
   }

   return ({value: value, type: valuetype, error: errortext});
};

/**
 * @function LookupResultType
 * @memberof SocialCalc.Formula
 * @description Determines the result type when performing operations on two operands
 * 
 * The typelookup parameter has values of the following form:
 * typelookup{"typespec1"} = "|typespec2A:resultA|typespec2B:resultB|..."
 * 
 * First type1 is looked up. If no match, then the first letter (major type) of type1 plus "*" is looked up.
 * Result type is type1 if result is "1", type2 if result is "2", otherwise the value of result.
 * 
 * @param {string} type1 - Type of the first operand
 * @param {string} type2 - Type of the second operand
 * @param {Object<string, string>} typelookup - Lookup table for type combinations
 * @returns {string} The resulting type after the operation
 * 
 * @example
 * const resultType = SocialCalc.Formula.LookupResultType("n", "t", typelookup.plus);
 * // Returns appropriate type based on lookup table
 */
SocialCalc.Formula.LookupResultType = function(type1, type2, typelookup) {
   let pos1, pos2, result;

   let table1 = typelookup[type1];

   if (!table1) {
      table1 = typelookup[type1.charAt(0)+'*'];
      if (!table1) {
         return "e#VALUE! (internal error, missing LookupResultType "+type1.charAt(0)+"*)"; // missing from table -- please add it
      }
   }
   
   pos1 = table1.indexOf("|"+type2+":");
   if (pos1 >= 0) {
      pos2 = table1.indexOf("|", pos1+1);
      if (pos2 < 0) return "e#VALUE! (internal error, incorrect LookupResultType "+table1+")";
      result = table1.substring(pos1+type2.length+2, pos2);
      if (result == "1") return type1;
      if (result == "2") return type2;
      return result;
   }
   
   pos1 = table1.indexOf("|"+type2.charAt(0)+"*:");
   if (pos1 >= 0) {
      pos2 = table1.indexOf("|", pos1+1);
      if (pos2 < 0) return "e#VALUE! (internal error, incorrect LookupResultType "+table1+")";
      result = table1.substring(pos1+4, pos2);
      if (result == "1") return type1;
      if (result == "2") return type2;
      return result;
   }
   
   return "e#VALUE!";
};

/**
 * @function TopOfStackValueAndType
 * @memberof SocialCalc.Formula
 * @description Returns top of stack value and type, then pops the stack
 * 
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @param {Array<Object>} operand - The operand stack
 * @returns {Object} Result object containing:
 *   - {*} value - The value from the top of stack
 *   - {string} type - The type of the value
 *   - {string} [error] - Error message if operation failed
 * 
 * @example
 * const stackTop = SocialCalc.Formula.TopOfStackValueAndType(sheet, operandStack);
 * // Returns: {value: 42, type: "n"} and pops the stack
 */
SocialCalc.Formula.TopOfStackValueAndType = function(sheet, operand) {
   let cellvtype, cell, pos, coordsheet;
   const scf = SocialCalc.Formula;

   const result = {type: "", value: ""};

   const stacklen = operand.length;

   if (!stacklen) { // make sure something is there
      result.error = SocialCalc.Constants.s_InternalError+"no operand on stack";
      return result;
   }

   result.value = operand[stacklen-1].value; // get top of stack
   result.type = operand[stacklen-1].type;
   operand.pop(); // we have data - pop stack

   if (result.type == "name") {
      result = scf.LookupName(sheet, result.value);
   }

   return result;
};

/**
 * @function OperandAsNumber
 * @memberof SocialCalc.Formula
 * @description Gets top of stack operand and converts it to a numeric value
 * 
 * Uses OperandValueAndType to get top of stack and pops it.
 * Returns numeric value and type. Text values are treated as 0 if they can't be converted.
 * 
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @param {Array<Object>} operand - The operand stack
 * @returns {Object} Result object containing:
 *   - {number} value - The numeric value
 *   - {string} type - The type of the operand
 * 
 * @example
 * const numOperand = SocialCalc.Formula.OperandAsNumber(sheet, operandStack);
 * // Returns: {value: 42, type: "n"}
 */
SocialCalc.Formula.OperandAsNumber = function(sheet, operand) {
   let t, valueinfo;
   const operandinfo = SocialCalc.Formula.OperandValueAndType(sheet, operand);

   t = operandinfo.type.charAt(0);

   if (t == "n") {
      operandinfo.value = operandinfo.value-0;
   } else if (t == "b") { // blank cell
      operandinfo.type = "n";
      operandinfo.value = 0;
   } else if (t == "e") { // error
      operandinfo.value = 0;
   } else {
      valueinfo = SocialCalc.DetermineValueType ? SocialCalc.DetermineValueType(operandinfo.value) :
                                                  {value: operandinfo.value-0, type: "n"}; // if without rest of SocialCalc
      if (valueinfo.type.charAt(0) == "n") {
         operandinfo.value = valueinfo.value-0;
         operandinfo.type = valueinfo.type;
      } else {
         operandinfo.value = 0;
         operandinfo.type = valueinfo.type;
      }
   }

   return operandinfo;
};

/**
 * @function OperandAsText
 * @memberof SocialCalc.Formula
 * @description Gets top of stack operand and converts it to a text value
 * 
 * Uses OperandValueAndType to get top of stack and pops it.
 * Returns text value, preserving sub-type where appropriate.
 * 
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @param {Array<Object>} operand - The operand stack
 * @returns {Object} Result object containing:
 *   - {string} value - The text value
 *   - {string} type - The type of the operand (typically "t" for text)
 * 
 * @example
 * const textOperand = SocialCalc.Formula.OperandAsText(sheet, operandStack);
 * // Returns: {value: "Hello", type: "t"}
 */
SocialCalc.Formula.OperandAsText = function(sheet, operand) {
   let t, valueinfo;
   const operandinfo = SocialCalc.Formula.OperandValueAndType(sheet, operand);

   t = operandinfo.type.charAt(0);

   if (t == "t") { // any flavor of text returns as is
      // Keep as is
   } else if (t == "n") {
      operandinfo.value = SocialCalc.format_number_for_display ?
                         SocialCalc.format_number_for_display(operandinfo.value, operandinfo.type, "") :
                         operandinfo.value = operandinfo.value+"";
      operandinfo.type = "t";
   } else if (t == "b") { // blank
      operandinfo.value = "";
      operandinfo.type = "t";
   } else if (t == "e") { // error
      operandinfo.value = "";
   } else {
      operand.value = operandinfo.value + "";
      operand.type = "t";
   }

   return operandinfo;
};
/**
 * @function OperandValueAndType
 * @memberof SocialCalc.Formula
 * @description Pops the top of stack and returns it, following coordinate references if necessary
 * 
 * Ranges are returned as if they were pushed onto the stack first coord first.
 * Also sets type with "t", "n", "th", etc., as appropriate.
 * 
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @param {Array<Object>} operand - The operand stack
 * @returns {Object} Result object containing:
 *   - {*} value - The value from the operand
 *   - {string} type - The type of the value ("t", "n", "coord", "range", etc.)
 *   - {string} [error] - Error message if operation failed
 * 
 * @example
 * const result = SocialCalc.Formula.OperandValueAndType(sheet, operandStack);
 * // Returns: {value: 42, type: "n"} for a numeric cell value
 */
SocialCalc.Formula.OperandValueAndType = function(sheet, operand) {
   let cellvtype, cell, pos, coordsheet;
   const scf = SocialCalc.Formula;

   const result = {type: "", value: ""};

   const stacklen = operand.length;

   if (!stacklen) { // make sure something is there
      result.error = SocialCalc.Constants.s_InternalError+"no operand on stack";
      return result;
   }

   result.value = operand[stacklen-1].value; // get top of stack
   result.type = operand[stacklen-1].type;
   operand.pop(); // we have data - pop stack

   if (result.type == "name") {
      result = scf.LookupName(sheet, result.value);
   }

   if (result.type == "range") {
      result = scf.StepThroughRangeDown(operand, result.value);
   }

   if (result.type == "coord") { // value is a coord reference
      coordsheet = sheet;
      pos = result.value.indexOf("!");
      if (pos != -1) { // sheet reference
         coordsheet = scf.FindInSheetCache(result.value.substring(pos+1)); // get other sheet
         if (coordsheet == null) { // unavailable
            result.type = "e#REF!";
            result.error = SocialCalc.Constants.s_sheetunavailable+" "+result.value.substring(pos+1);
            result.value = 0;
            return result;
         }
         result.value = result.value.substring(0, pos); // get coord part
      }

      if (coordsheet) {
         cell = coordsheet.cells[SocialCalc.Formula.PlainCoord(result.value)];
         if (cell) {
            cellvtype = cell.valuetype; // get type of value in the cell it points to
            result.value = cell.datavalue;
         } else {
            cellvtype = "b";
         }
      } else {
         cellvtype = "e#N/A";
         result.value = 0;
      }
      result.type = cellvtype || "b";
      if (result.type == "b") { // blank
         result.value = 0;
      }
   }

   return result;
};

/**
 * @function OperandAsCoord
 * @memberof SocialCalc.Formula
 * @description Gets top of stack and pops it, expecting a coordinate value
 * 
 * Returns coordinate value. All other types are treated as an error.
 * 
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @param {Array<Object>} operand - The operand stack
 * @returns {Object} Result object containing:
 *   - {string} value - The coordinate value or error message
 *   - {string} type - "coord" if successful, "e#REF!" if error
 * 
 * @example
 * const coord = SocialCalc.Formula.OperandAsCoord(sheet, operandStack);
 * // Returns: {value: "A1", type: "coord"}
 */
SocialCalc.Formula.OperandAsCoord = function(sheet, operand) {
   const scf = SocialCalc.Formula;

   const result = {type: "", value: ""};

   const stacklen = operand.length;

   result.value = operand[stacklen-1].value; // get top of stack
   result.type = operand[stacklen-1].type;
   operand.pop(); // we have data - pop stack
   
   if (result.type == "name") {
      result = SocialCalc.Formula.LookupName(sheet, result.value);
   }
   
   if (result.type == "coord") { // value is a coord reference
      return result;
   } else {
      result.value = SocialCalc.Constants.s_calcerrcellrefmissing;
      result.type = "e#REF!";
      return result;
   }
};

/**
 * @function OperandsAsCoordOnSheet
 * @memberof SocialCalc.Formula
 * @description Gets 2 operands from top of stack, treating them as sheetname!coord-or-name
 * 
 * Returns stack-style coord value (coord!sheetname, or coord!sheetname|coord|) with
 * a type of coord or range. All others are treated as an error.
 * If sheetname not available, sets result.error.
 * 
 * @param {Object} sheet - The current spreadsheet object
 * @param {Array<Object>} operand - The operand stack
 * @returns {Object} Result object containing:
 *   - {string} value - The coordinate with sheet reference
 *   - {string} type - "coord", "range", or "e#REF!" if error
 *   - {string} [error] - Error message if sheet unavailable
 * 
 * @example
 * const result = SocialCalc.Formula.OperandsAsCoordOnSheet(sheet, operandStack);
 * // Returns: {value: "A1!Sheet2", type: "coord"}
 */
SocialCalc.Formula.OperandsAsCoordOnSheet = function(sheet, operand) {
   let sheetname, othersheet, pos1, pos2;
   const value1 = {};
   const result = {};
   const scf = SocialCalc.Formula;

   const stacklen = operand.length;
   value1.value = operand[stacklen-1].value; // get top of stack - coord or name
   value1.type = operand[stacklen-1].type;
   operand.pop(); // we have data - pop stack

   sheetname = scf.OperandAsSheetName(sheet, operand); // get sheetname as text
   othersheet = scf.FindInSheetCache(sheetname.value);
   if (othersheet == null) { // unavailable
      result.type = "e#REF!";
      result.value = 0;
      result.error = SocialCalc.Constants.s_sheetunavailable+" "+sheetname.value;
      return result;
   }

   if (value1.type == "name") {
      value1 = scf.LookupName(othersheet, value1.value);
   }
   
   result.type = value1.type;
   if (value1.type == "coord") { // value is a coord reference
      result.value = value1.value + "!" + sheetname.value; // return in the format as used on stack
   } else if (value1.type == "range") { // value is a range reference
      pos1 = value1.value.indexOf("|");
      pos2 = value1.value.indexOf("|", pos1+1);
      result.value = value1.value.substring(0, pos1) + "!" + sheetname.value +
                    "|" + value1.value.substring(pos1+1, pos2) + "|";
   } else if (value1.type.charAt(0)=="e") {
      result.value = value1.value;
   } else {
      result.error = SocialCalc.Constants.s_calcerrcellrefmissing;
      result.type = "e#REF!";
      result.value = 0;
   }
   
   return result;
};

/**
 * @function OperandsAsRangeOnSheet
 * @memberof SocialCalc.Formula
 * @description Gets 2 operands from top of stack, treating them as coord2-or-name:coord1
 * 
 * Name is evaluated on sheet of coord1.
 * Returns result with "value" of stack-style range value (coord!sheetname|coord|) and
 * "type" of "range". All others are treated as an error.
 * 
 * @param {Object} sheet - The current spreadsheet object
 * @param {Array<Object>} operand - The operand stack
 * @returns {Object} Result object containing:
 *   - {string} value - The range in stack format or error message
 *   - {string} type - "range" if successful, "e#REF!" if error
 *   - {string} [errortext] - Error message if sheet unavailable
 * 
 * @example
 * const range = SocialCalc.Formula.OperandsAsRangeOnSheet(sheet, operandStack);
 * // Returns: {value: "A1!Sheet1|B2|", type: "range"}
 */
SocialCalc.Formula.OperandsAsRangeOnSheet = function(sheet, operand) {
   let value1, othersheet, pos1, pos2;
   const value2 = {};
   const scf = SocialCalc.Formula;
   const scc = SocialCalc.Constants;

   const stacklen = operand.length;
   value2.value = operand[stacklen-1].value; // get top of stack - coord or name for "right" side
   value2.type = operand[stacklen-1].type;
   operand.pop(); // we have data - pop stack

   value1 = scf.OperandAsCoord(sheet, operand); // get "left" coord
   if (value1.type != "coord") { // not a coord, which it must be
      return {value: 0, type: "e#REF!"};
   }

   othersheet = sheet;
   pos1 = value1.value.indexOf("!");
   if (pos1 != -1) { // sheet reference
      pos2 = value1.value.indexOf("|", pos1+1);
      if (pos2 < 0) pos2 = value1.value.length;
      othersheet = scf.FindInSheetCache(value1.value.substring(pos1+1,pos2)); // get other sheet
      if (othersheet == null) { // unavailable
         return {value: 0, type: "e#REF!", errortext: scc.s_sheetunavailable+" "+value1.value.substring(pos1+1,pos2)};
      }
   }

   if (value2.type == "name") { // coord:name is allowed, if name is just one cell
      value2 = scf.LookupName(othersheet, value2.value);
   }

   if (value2.type == "coord") { // value is a coord reference, so return the combined range
      return {value: value1.value+"|"+value2.value+"|", type: "range"}; // return range in the format as used on stack
   } else { // bad form
      return {value: scc.s_calcerrcellrefmissing, type: "e#REF!"};
   }
};

/**
 * @function OperandAsSheetName
 * @memberof SocialCalc.Formula
 * @description Gets top of stack and pops it, expecting a sheet name value
 * 
 * Returns sheetname value. All other types are treated as an error.
 * Accepts text, cell reference, and named value which is one of those two.
 * 
 * @param {Object} sheet - The current spreadsheet object
 * @param {Array<Object>} operand - The operand stack
 * @returns {Object} Result object containing:
 *   - {string} value - The sheet name or empty string if error
 *   - {string} type - Type of the value or error type
 *   - {string} [error] - Error message if sheet name missing
 * 
 * @example
 * const sheetName = SocialCalc.Formula.OperandAsSheetName(sheet, operandStack);
 * // Returns: {value: "Sheet2", type: "t"}
 */
SocialCalc.Formula.OperandAsSheetName = function(sheet, operand) {
   let nvalue, cell;

   const scf = SocialCalc.Formula;

   const result = {type: "", value: ""};

   const stacklen = operand.length;

   result.value = operand[stacklen-1].value; // get top of stack
   result.type = operand[stacklen-1].type;
   operand.pop(); // we have data - pop stack
   
   if (result.type == "name") {
      nvalue = SocialCalc.Formula.LookupName(sheet, result.value);
      if (!nvalue.value) { // not a known name - return bare name as the name value
         return result;
      }
      result.value = nvalue.value;
      result.type = nvalue.type;
   }
   
   if (result.type == "coord") { // value is a coord reference, follow it to find sheet name
      cell = sheet.cells[SocialCalc.Formula.PlainCoord(result.value)];
      if (cell) {
         result.value = cell.datavalue;
         result.type = cell.valuetype;
      } else {
         result.value = "";
         result.type = "b";
      }
   }
   
   if (result.type.charAt(0) == "t") { // value is a string which could be a sheet name
      return result;
   } else {
      result.value = "";
      result.error = SocialCalc.Constants.s_calcerrsheetnamemissing;
      return result;
   }
};
/**
 * @function LookupName
 * @memberof SocialCalc.Formula
 * @description Returns value and type of a named value
 * 
 * Names are case insensitive. Names may have a definition which is:
 * - A coord (A1)
 * - A range (A1:B7) 
 * - A formula (=OFFSET(A1,0,0,5,1))
 * Note: The range must not have sheet names ("!") in them.
 * 
 * @param {Object} sheet - The spreadsheet object containing names and cell data
 * @param {string} name - The name to lookup (case insensitive)
 * @returns {Object} Result object containing:
 *   - {*} value - The resolved value of the name
 *   - {string} type - The type ("coord", "range", "n", "t", etc.)
 *   - {string} [error] - Error message if lookup failed
 * 
 * @example
 * const result = SocialCalc.Formula.LookupName(sheet, "MyRange");
 * // Returns: {value: "A1|B3|", type: "range"} for a range definition
 */
SocialCalc.Formula.LookupName = function(sheet, name) {
   let pos, specialc, parseinfo;
   const names = sheet.names;
   const value = {};
   let startedwalk = false;

   if (names[name.toUpperCase()]) { // is name defined?
      value.value = names[name.toUpperCase()].definition; // yes

      if (value.value.charAt(0) == "=") { // formula
         if (!sheet.checknamecirc) { // are we possibly walking the name tree?
            sheet.checknamecirc = {}; // not yet
            startedwalk = true; // remember we are the reference that started it
         } else {
            if (sheet.checknamecirc[name]) { // circular reference
               value.type = "e#NAME?";
               value.error = SocialCalc.Constants.s_circularnameref+' "' + name + '".';
               return value;
            }
         }
         sheet.checknamecirc[name] = true;

         parseinfo = SocialCalc.Formula.ParseFormulaIntoTokens(value.value.substring(1));
         value = SocialCalc.Formula.evaluate_parsed_formula(parseinfo, sheet, 1); // parse formula, allowing range return

         delete sheet.checknamecirc[name]; // done with us
         if (startedwalk) {
            delete sheet.checknamecirc; // done with walk
         }

         if (value.type != "range") {
            return value;
         }
      }

      pos = value.value.indexOf(":");
      if (pos != -1) { // range
         value.type = "range";
         value.value = value.value.substring(0, pos) + "|" + value.value.substring(pos+1)+"|";
         value.value = value.value.toUpperCase();
      } else {
         value.type = "coord";
         value.value = value.value.toUpperCase();
      }
      return value;
   } else if (specialc = SocialCalc.Formula.SpecialConstants[name.toUpperCase()]) { // special constant, like #REF!
      pos = specialc.indexOf(",");
      value.value = specialc.substring(0,pos)-0;
      value.type = specialc.substring(pos+1);
      return value;
   } else {
      value.value = "";
      value.type = "e#NAME?";
      value.error = SocialCalc.Constants.s_calcerrunknownname+' "'+name+'"';
      return value;
   }
};

/**
 * @function StepThroughRangeDown
 * @memberof SocialCalc.Formula
 * @description Returns next coordinate in a range, keeping track on the operand stack
 * 
 * Goes from upper left across and down to bottom right through the range.
 * 
 * @param {Array<Object>} operand - The operand stack for tracking position
 * @param {string} rangevalue - Range value in format "coord1|coord2|sequence"
 * @returns {Object} Result object containing:
 *   - {string} value - The next coordinate in the range
 *   - {string} type - Always "coord"
 * 
 * @example
 * const nextCoord = SocialCalc.Formula.StepThroughRangeDown(operand, "A1|B2|0");
 * // Returns: {value: "A1", type: "coord"} and updates operand stack
 */
SocialCalc.Formula.StepThroughRangeDown = function(operand, rangevalue) {
   let value1, value2, sequence, pos1, pos2, sheet1, rp, c, r, count;
   const scf = SocialCalc.Formula;

   pos1 = rangevalue.indexOf("|");
   pos2 = rangevalue.indexOf("|", pos1+1);
   value1 = rangevalue.substring(0, pos1);
   value2 = rangevalue.substring(pos1+1, pos2);
   sequence = rangevalue.substring(pos2+1) - 0;

   pos1 = value1.indexOf("!");
   if (pos1 != -1) {
      sheet1 = value1.substring(pos1);
      value1 = value1.substring(0, pos1);
   } else {
      sheet1 = "";
   }
   
   pos1 = value2.indexOf("!");
   if (pos1 != -1) {
      value2 = value2.substring(0, pos1);
   }

   rp = scf.OrderRangeParts(value1, value2);
   
   count = 0;
   for (r = rp.r1; r <= rp.r2; r++) {
      for (c = rp.c1; c <= rp.c2; c++) {
         count++;
         if (count > sequence) {
            if (r != rp.r2 || c != rp.c2) { // keep on stack until done
               scf.PushOperand(operand, "range", value1+sheet1+"|"+value2+"|"+count);
            }
            return {value: SocialCalc.crToCoord(c, r)+sheet1, type: "coord"};
         }
      }
   }
};

/**
 * @function DecodeRangeParts
 * @memberof SocialCalc.Formula
 * @description Decodes a range into its component parts and sheet information
 * 
 * Returns sheetdata for the sheet where the range is, as well as
 * the number of the first column in the range, the number of columns,
 * and equivalent row information.
 * 
 * @param {Object} sheetdata - The current sheet data object
 * @param {string} range - Range in format "coord1|coord2|sequence"
 * @returns {Object|null} Result object containing:
 *   - {Object} sheetdata - Sheet data object for the range
 *   - {string} sheetname - Sheet name or empty string
 *   - {number} col1num - First column number
 *   - {number} ncols - Number of columns
 *   - {number} row1num - First row number  
 *   - {number} nrows - Number of rows
 *   Returns null if any errors occur.
 * 
 * @example
 * const parts = SocialCalc.Formula.DecodeRangeParts(sheet, "A1|C3|");
 * // Returns: {sheetdata: sheet, sheetname: "", col1num: 1, ncols: 3, row1num: 1, nrows: 3}
 */
SocialCalc.Formula.DecodeRangeParts = function(sheetdata, range) {
   let value1, value2, pos1, pos2, sheet1, coordsheetdata, rp;

   const scf = SocialCalc.Formula;

   pos1 = range.indexOf("|");
   pos2 = range.indexOf("|", pos1+1);
   value1 = range.substring(0, pos1);
   value2 = range.substring(pos1+1, pos2);

   pos1 = value1.indexOf("!");
   if (pos1 != -1) {
      sheet1 = value1.substring(pos1+1);
      value1 = value1.substring(0, pos1);
   } else {
      sheet1 = "";
   }
   
   pos1 = value2.indexOf("!");
   if (pos1 != -1) {
      value2 = value2.substring(0, pos1);
   }

   coordsheetdata = sheetdata;
   if (sheet1) { // sheet reference
      coordsheetdata = scf.FindInSheetCache(sheet1);
      if (coordsheetdata == null) { // unavailable
         return null;
      }
   }

   rp = scf.OrderRangeParts(value1, value2);

   return {
      sheetdata: coordsheetdata, 
      sheetname: sheet1, 
      col1num: rp.c1, 
      ncols: rp.c2-rp.c1+1, 
      row1num: rp.r1, 
      nrows: rp.r2-rp.r1+1
   };
};

/**
 * @namespace SocialCalc.Formula Function Handling
 * @description Function handling system for spreadsheet formulas
 */

/**
 * @type {Object<string, Array>}
 * @description List of functions with their definitions
 * 
 * Format: SocialCalc.Formula.FunctionList["function_name"] = [function_subroutine, number_of_arguments, arg_def, func_def, func_class]
 * 
 * - function_subroutine: takes arguments (fname, operand, foperand, sheet), returns errortext or null
 * - number_of_arguments: 0=no args, >0=exact count, <0=minimum count, 100=don't check
 * - arg_def: optional name of element in SocialCalc.Formula.FunctionArgDefs
 * - func_def: optional string explaining the function, or looked up in SocialCalc.Constants
 * - func_class: optional comma-separated class names from SocialCalc.Formula.FunctionClasses
 */
if (!SocialCalc.Formula.FunctionList) { // make sure it is defined (could have been in another module)
   SocialCalc.Formula.FunctionList = {};
}

/**
 * @type {Object<string, Object>|null}
 * @description Function classes for organizing functions
 * Format: FunctionClasses[classname] = {name: full-name-string, items: [sorted list of function names]}
 * Filled in by SocialCalc.Formula.FillFunctionInfo
 */
SocialCalc.Formula.FunctionClasses = null; // start null to say needs filling in

/**
 * @type {Object<string, string>}
 * @description Function argument definitions
 * Format: FunctionArgDef[argname] = explicit-string-for-arg-list
 * Filled in by SocialCalc.Formula.FillFunctionInfo
 */
SocialCalc.Formula.FunctionArgDefs = {};

/**
 * @function CalculateFunction
 * @memberof SocialCalc.Formula
 * @description Dispatches function calls to their appropriate handlers
 * 
 * @param {string} fname - The function name to calculate
 * @param {Array<Object>} operand - The main operand stack
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {string} Error text if function call failed, empty string if successful
 * 
 * @example
 * const error = SocialCalc.Formula.CalculateFunction("SUM", operand, sheet);
 * // Returns: "" if successful, error message if failed
 */
SocialCalc.Formula.CalculateFunction = function(fname, operand, sheet) {
   let fobj, foperand, ffunc, argnum, ttext;
   const scf = SocialCalc.Formula;
   let ok = 1;
   let errortext = "";

   fobj = scf.FunctionList[fname];

   if (fobj) {
      foperand = [];
      ffunc = fobj[0];
      argnum = fobj[1];
      scf.CopyFunctionArgs(operand, foperand);
      
      if (argnum != 100) {
         if (argnum < 0) {
            if (foperand.length < -argnum) {
               errortext = scf.FunctionArgsError(fname, operand);
               return errortext;
            }
         } else {
            if (foperand.length != argnum) {
               errortext = scf.FunctionArgsError(fname, operand);
               return errortext;
            }
         }
      }
      errortext = ffunc(fname, operand, foperand, sheet);
   } else {
      ttext = fname;

      if (operand.length && operand[operand.length-1].type == "start") { // no arguments - name or zero arg function
         operand.pop();
         scf.PushOperand(operand, "name", ttext);
      } else {
         errortext = SocialCalc.Constants.s_sheetfuncunknownfunction+" " + ttext +". ";
      }
   }

   return errortext;
};
/**
 * @function PushOperand
 * @memberof SocialCalc.Formula
 * @description Pushes a type and value onto the operand stack
 * 
 * @param {Array<Object>} operand - The operand stack to push onto
 * @param {string} t - The type of the operand
 * @param {*} v - The value of the operand
 * 
 * @example
 * SocialCalc.Formula.PushOperand(operand, "n", 42);
 * // Pushes {type: "n", value: 42} onto the operand stack
 */
SocialCalc.Formula.PushOperand = function(operand, t, v) {
   operand.push({type: t, value: v});
};

/**
 * @function CopyFunctionArgs
 * @memberof SocialCalc.Formula
 * @description Pops operands from main stack and pushes onto function operand stack
 * 
 * Pops operands from operand and pushes on foperand up to function start,
 * reversing order in the process.
 * 
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack to populate
 * 
 * @example
 * SocialCalc.Formula.CopyFunctionArgs(operand, foperand);
 * // Transfers function arguments from main stack to function stack
 */
SocialCalc.Formula.CopyFunctionArgs = function(operand, foperand) {
   let fobj, ffunc, argnum;
   const scf = SocialCalc.Formula;
   let ok = 1;
   let errortext = null;

   while (operand.length > 0 && operand[operand.length-1].type != "start") { // get each arg
      foperand.push(operand.pop()); // copy it
   }
   operand.pop(); // get rid of "start"

   return;
};

/**
 * @function FunctionArgsError
 * @memberof SocialCalc.Formula
 * @description Pushes appropriate error on operand stack for incorrect function arguments
 * 
 * @param {string} fname - The function name that had the error
 * @param {Array<Object>} operand - The operand stack to push error onto
 * @returns {string} Error text including the function name
 * 
 * @example
 * const error = SocialCalc.Formula.FunctionArgsError("SUM", operand);
 * // Returns: "Incorrect arguments to function SUM. " and pushes error to stack
 */
SocialCalc.Formula.FunctionArgsError = function(fname, operand) {
   const errortext = SocialCalc.Constants.s_calcerrincorrectargstofunction+" " + fname + ". ";
   SocialCalc.Formula.PushOperand(operand, "e#VALUE!", errortext);

   return errortext;
};

/**
 * @function FunctionSpecificError
 * @memberof SocialCalc.Formula
 * @description Pushes specified error and text on operand stack
 * 
 * @param {string} fname - The function name (for context)
 * @param {Array<Object>} operand - The operand stack to push error onto
 * @param {string} errortype - The error type (e.g., "e#VALUE!", "e#DIV/0!")
 * @param {string} errortext - The error message text
 * @returns {string} The error text that was pushed
 * 
 * @example
 * const error = SocialCalc.Formula.FunctionSpecificError("SQRT", operand, "e#NUM!", "Negative number");
 * // Pushes custom error to stack and returns error text
 */
SocialCalc.Formula.FunctionSpecificError = function(fname, operand, errortype, errortext) {
   SocialCalc.Formula.PushOperand(operand, errortype, errortext);

   return errortext;
};

/**
 * @function CheckForErrorValue
 * @memberof SocialCalc.Formula
 * @description Checks if a value is an error and pushes it to stack if so
 * 
 * @param {Array<Object>} operand - The operand stack
 * @param {Object} v - The value object to check (with type and value properties)
 * @returns {boolean} True if v.type is an error (and pushed to stack), false otherwise
 * 
 * @example
 * const hasError = SocialCalc.Formula.CheckForErrorValue(operand, {type: "e#DIV/0!", value: 0});
 * // Returns: true and pushes error to operand stack
 */
SocialCalc.Formula.CheckForErrorValue = function(operand, v) {
   if (v.type.charAt(0) == "e") {
      operand.push(v);
      return true;
   } else {
      return false;
   }
};

/**
 * @namespace SocialCalc.Formula Function Information Routines
 * @description Routines for managing function definitions and metadata
 */

/**
 * @function FillFunctionInfo
 * @memberof SocialCalc.Formula
 * @description Processes function definitions and fills out FunctionArgDefs and FunctionClasses
 * 
 * Goes through function definitions and fills out FunctionArgDefs and FunctionClasses.
 * Execute this after any changes to SocialCalc.Constants but before UI is used.
 * 
 * @example
 * SocialCalc.Formula.FillFunctionInfo();
 * // Initializes function argument definitions and classes from constants
 */
SocialCalc.Formula.FillFunctionInfo = function() {
   const scf = SocialCalc.Formula;
   const scc = SocialCalc.Constants;

   let fname, f, classes, cname, i;

   if (scf.FunctionClasses) { // only do once
      return;
   }

   for (fname in scf.FunctionList) {
      f = scf.FunctionList[fname];
      if (f[2]) { // has an arg def
         scf.FunctionArgDefs[f[2]] = scc["s_farg_"+f[2]] || ""; // get it from constants
      }
      if (!f[3]) { // no text def, see if in constants
         if (scc["s_fdef_"+fname]) {
            scf.FunctionList[fname][3] = scc["s_fdef_"+fname];
         }
      }
   }

   scf.FunctionClasses = {};
 
   for (i = 0; i < scc.function_classlist.length; i++) {
      cname = scc.function_classlist[i];
      scf.FunctionClasses[cname] = {name: scc["s_fclass_"+cname], items: []};
   }

   for (fname in scf.FunctionList) {
      f = scf.FunctionList[fname];
      classes = f[4] ? f[4].split(",") : []; // get classes
      classes.push("all");
      for (i = 0; i < classes.length; i++) {
         cname = classes[i];
         scf.FunctionClasses[cname].items.push(fname);
      }
   }
   
   for (cname in scf.FunctionClasses) {
      scf.FunctionClasses[cname].items.sort();
   }
};

/**
 * @function FunctionArgString
 * @memberof SocialCalc.Formula
 * @description Returns a string representing the arguments to a function
 * 
 * @param {string} fname - The function name to get argument string for
 * @returns {string} String representation of the function's arguments
 * 
 * @example
 * const argStr = SocialCalc.Formula.FunctionArgString("SUM");
 * // Returns: "v1, v2, ..." or specific argument description
 */
SocialCalc.Formula.FunctionArgString = function(fname) {
   const scf = SocialCalc.Formula;
   const fdata = scf.FunctionList[fname];
   let nargs, i, str;

   const adef = fdata[2];

   if (!adef) {
      nargs = fdata[1];
      if (nargs == 0) {
         adef = " ";
      } else if (nargs > 0) {
         str = "v1";
         for (i = 2; i <= nargs; i++) {
            str += ", v"+i;
         }
         return str;
      } else if (nargs < 0) {
         str = "v1";
         for (i = 2; i < -nargs; i++) {
            str += ", v"+i;
         }
         return str+", ...";
      } else {
         return "nargs: "+nargs;
      }
   }

   str = scf.FunctionArgDefs[adef] || adef;

   return str;
};

/**
 * @namespace SocialCalc.Formula Function Definitions
 * @description Standard function definitions for spreadsheet calculations
 * Note that some functions require SocialCalc.DetermineValueType to be defined.
 */

/**
 * @function SeriesFunctions
 * @memberof SocialCalc.Formula
 * @description Calculates statistical functions that operate on series of values
 * 
 * Handles: AVERAGE, COUNT, COUNTA, COUNTBLANK, MAX, MIN, PRODUCT, STDEV, STDEVP, SUM, VAR, VARP
 * 
 * Calculates all statistical measures and then returns the desired one.
 * The overhead is in accessing values, not calculating, so this approach is efficient.
 * If this routine is changed, check the dseries_functions, too.
 * 
 * @param {string} fname - The specific function name to calculate
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {null} Always returns null (errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction for functions like SUM, AVERAGE, etc.
 * SocialCalc.Formula.SeriesFunctions("SUM", operand, foperand, sheet);
 */
SocialCalc.Formula.SeriesFunctions = function(fname, operand, foperand, sheet) {
   let value1, t, v1;

   const scf = SocialCalc.Formula;
   const operand_value_and_type = scf.OperandValueAndType;
   const lookup_result_type = scf.LookupResultType;
   const typelookupplus = scf.TypeLookupTable.plus;

   /**
    * @function PushOperand
    * @description Helper function to push operand onto stack
    * @param {string} t - Type of operand
    * @param {*} v - Value of operand
    */
   const PushOperand = function(t, v) {operand.push({type: t, value: v});};

   let sum = 0;
   let resulttypesum = "";
   let count = 0;
   let counta = 0;
   let countblank = 0;
   let product = 1;
   let maxval;
   let minval;
   /** @type {number} M sub k for variance calculations as per Knuth */
   let mk, sk, mk1, sk1; // For variance, etc.: M sub k, k-1, and S sub k-1
                          // as per Knuth "The Art of Computer Programming" Vol. 2 3rd edition, page 232

   while (foperand.length > 0) {
      value1 = operand_value_and_type(sheet, foperand);
      t = value1.type.charAt(0);
      if (t == "n") count += 1;
      if (t != "b") counta += 1;
      if (t == "b") countblank += 1;

      if (t == "n") {
         v1 = value1.value-0; // get it as a number
         sum += v1;
         product *= v1;
         maxval = (maxval != undefined) ? (v1 > maxval ? v1 : maxval) : v1;
         minval = (minval != undefined) ? (v1 < minval ? v1 : minval) : v1;
         
         if (count == 1) { // initialize with first values for variance used in STDEV, VAR, etc.
            mk1 = v1;
            sk1 = 0;
         } else { // Accumulate S sub 1 through n as per Knuth noted above
            mk = mk1 + (v1 - mk1) / count;
            sk = sk1 + (v1 - mk1) * (v1 - mk);
            sk1 = sk;
            mk1 = mk;
         }
         resulttypesum = lookup_result_type(value1.type, resulttypesum || value1.type, typelookupplus);
      } else if (t == "e" && resulttypesum.charAt(0) != "e") {
         resulttypesum = value1.type;
      }
   }

   resulttypesum = resulttypesum || "n";

   switch (fname) {
      case "SUM":
         PushOperand(resulttypesum, sum);
         break;

      case "PRODUCT": // may handle cases with text differently than some other spreadsheets
         PushOperand(resulttypesum, product);
         break;

      case "MIN":
         PushOperand(resulttypesum, minval || 0);
         break;

      case "MAX":
         PushOperand(resulttypesum, maxval || 0);
         break;

      case "COUNT":
         PushOperand("n", count);
         break;

      case "COUNTA":
         PushOperand("n", counta);
         break;

      case "COUNTBLANK":
         PushOperand("n", countblank);
         break;

      case "AVERAGE":
         if (count > 0) {
            PushOperand(resulttypesum, sum/count);
         } else {
            PushOperand("e#DIV/0!", 0);
         }
         break;

      case "STDEV":
         if (count > 1) {
            PushOperand(resulttypesum, Math.sqrt(sk / (count - 1))); // sk is never negative according to Knuth
         } else {
            PushOperand("e#DIV/0!", 0);
         }
         break;

      case "STDEVP":
         if (count > 1) {
            PushOperand(resulttypesum, Math.sqrt(sk / count));
         } else {
            PushOperand("e#DIV/0!", 0);
         }
         break;

      case "VAR":
         if (count > 1) {
            PushOperand(resulttypesum, sk / (count - 1));
         } else {
            PushOperand("e#DIV/0!", 0);
         }
         break;

      case "VARP":
         if (count > 1) {
            PushOperand(resulttypesum, sk / count);
         } else {
            PushOperand("e#DIV/0!", 0);
         }
         break;
   }

   return null;
};
/**
 * @description Function list registrations for statistical series functions
 * Each entry follows the format: [function_handler, arg_count, arg_def, func_def, func_class]
 * - arg_count: -1 means one or more arguments
 * - arg_def: "vn" refers to variable number of numeric arguments
 * - func_class: "stat" indicates statistical function category
 */

// Add to function list
SocialCalc.Formula.FunctionList["AVERAGE"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["COUNT"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["COUNTA"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["COUNTBLANK"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["MAX"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["MIN"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["PRODUCT"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["STDEV"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["STDEVP"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["SUM"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["VAR"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];
SocialCalc.Formula.FunctionList["VARP"] = [SocialCalc.Formula.SeriesFunctions, -1, "vn", null, "stat"];

/**
 * @function DSeriesFunctions
 * @memberof SocialCalc.Formula
 * @description Calculates database statistical functions that operate on filtered data
 * 
 * Handles: DAVERAGE, DCOUNT, DCOUNTA, DGET, DMAX, DMIN, DPRODUCT, DSTDEV, DSTDEVP, DSUM, DVAR, DVARP
 * 
 * These functions work on database ranges with criteria filtering:
 * - Database range contains the data with headers in first row
 * - Field name specifies which column to perform the calculation on
 * - Criteria range specifies filtering conditions
 * 
 * Calculates all statistical measures and then returns the desired one.
 * The overhead is in accessing values, not calculating, so this approach is efficient.
 * If this routine is changed, check the series_functions, too.
 * 
 * @param {string} fname - The specific database function name to calculate
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction for functions like DSUM, DAVERAGE, etc.
 * // =DSUM(A1:D10, "Sales", F1:F2) - sums Sales column where criteria in F1:F2 are met
 * SocialCalc.Formula.DSeriesFunctions("DSUM", operand, foperand, sheet);
 */
SocialCalc.Formula.DSeriesFunctions = function(fname, operand, foperand, sheet) {
   let value1, tostype, cr, dbrange, fieldname, criteriarange, dbinfo, criteriainfo;
   let fieldasnum, targetcol, i, j, k, cell, criteriafieldnums;
   let testok, criteriacr, criteria, testcol, testcr;
   let t;

   const scf = SocialCalc.Formula;
   const operand_value_and_type = scf.OperandValueAndType;
   const lookup_result_type = scf.LookupResultType;
   const typelookupplus = scf.TypeLookupTable.plus;

   /**
    * @function PushOperand
    * @description Helper function to push operand onto stack
    * @param {string} t - Type of operand
    * @param {*} v - Value of operand
    */
   const PushOperand = function(t, v) {operand.push({type: t, value: v});};

   value1 = {};

   // Statistical accumulator variables
   let sum = 0;
   let resulttypesum = "";
   let count = 0;
   let counta = 0;
   let countblank = 0;
   let product = 1;
   let maxval;
   let minval;
   /** @type {number} Variance calculation variables as per Knuth */
   let mk, sk, mk1, sk1; // For variance, etc.: M sub k, k-1, and S sub k-1
                          // as per Knuth "The Art of Computer Programming" Vol. 2 3rd edition, page 232

   // Get function arguments: database range, field name, criteria range
   dbrange = scf.TopOfStackValueAndType(sheet, foperand); // get a range
   fieldname = scf.OperandValueAndType(sheet, foperand); // get a value
   criteriarange = scf.TopOfStackValueAndType(sheet, foperand); // get a range

   if (dbrange.type != "range" || criteriarange.type != "range") {
      return scf.FunctionArgsError(fname, operand);
   }

   // Decode the ranges into usable information
   dbinfo = scf.DecodeRangeParts(sheet, dbrange.value);
   criteriainfo = scf.DecodeRangeParts(sheet, criteriarange.value);

   // Find the target column number for the field to calculate on
   fieldasnum = scf.FieldToColnum(dbinfo.sheetdata, dbinfo.col1num, dbinfo.ncols, dbinfo.row1num, fieldname.value, fieldname.type);
   if (fieldasnum <= 0) {
      PushOperand("e#VALUE!", 0);
      return;
   }

   targetcol = dbinfo.col1num + fieldasnum - 1;
   criteriafieldnums = [];

   // Map criteria field names to column numbers
   for (i = 0; i < criteriainfo.ncols; i++) { // get criteria field colnums
      cell = criteriainfo.sheetdata.GetAssuredCell(SocialCalc.crToCoord(criteriainfo.col1num + i, criteriainfo.row1num));
      criterianum = scf.FieldToColnum(dbinfo.sheetdata, dbinfo.col1num, dbinfo.ncols, dbinfo.row1num, cell.datavalue, cell.valuetype);
      if (criterianum <= 0) {
         PushOperand("e#VALUE!", 0);
         return;
      }
      criteriafieldnums.push(dbinfo.col1num + criterianum - 1);
   }

   // Process each row of the database (skip header row)
   for (i = 1; i < dbinfo.nrows; i++) { // go through each row of the database
      testok = false;
      
      // Test against all criteria rows
CRITERIAROW:
      for (j = 1; j < criteriainfo.nrows; j++) { // go through each criteria row
         // Check all criteria columns for this criteria row
         for (k = 0; k < criteriainfo.ncols; k++) { // look at each column
            criteriacr = SocialCalc.crToCoord(criteriainfo.col1num + k, criteriainfo.row1num + j); // where criteria is
            cell = criteriainfo.sheetdata.GetAssuredCell(criteriacr);
            criteria = cell.datavalue;
            if (typeof criteria == "string" && criteria.length == 0) continue; // blank items are OK
            
            testcol = criteriafieldnums[k];
            testcr = SocialCalc.crToCoord(testcol, dbinfo.row1num + i); // cell to check
            cell = criteriainfo.sheetdata.GetAssuredCell(testcr);
            
            if (!scf.TestCriteria(cell.datavalue, cell.valuetype || "b", criteria)) {
               continue CRITERIAROW; // does not meet criteria - check next row
            }
         }
         testok = true; // met all the criteria
         break CRITERIAROW;
      }
      
      if (!testok) {
         continue; // Skip this database row if it doesn't meet any criteria row
      }

      // Get the target cell value for statistical calculation
      cr = SocialCalc.crToCoord(targetcol, dbinfo.row1num + i); // get cell of this row to do the function on
      cell = dbinfo.sheetdata.GetAssuredCell(cr);

      value1.value = cell.datavalue;
      value1.type = cell.valuetype;
      t = value1.type.charAt(0);
      
      // Update counters
      if (t == "n") count += 1;
      if (t != "b") counta += 1;
      if (t == "b") countblank += 1;

      // Perform statistical calculations for numeric values
      if (t == "n") {
         v1 = value1.value-0; // get it as a number
         sum += v1;
         product *= v1;
         maxval = (maxval != undefined) ? (v1 > maxval ? v1 : maxval) : v1;
         minval = (minval != undefined) ? (v1 < minval ? v1 : minval) : v1;
         
         if (count == 1) { // initialize with first values for variance used in STDEV, VAR, etc.
            mk1 = v1;
            sk1 = 0;
         } else { // Accumulate S sub 1 through n as per Knuth noted above
            mk = mk1 + (v1 - mk1) / count;
            sk = sk1 + (v1 - mk1) * (v1 - mk);
            sk1 = sk;
            mk1 = mk;
         }
         resulttypesum = lookup_result_type(value1.type, resulttypesum || value1.type, typelookupplus);
      } else if (t == "e" && resulttypesum.charAt(0) != "e") {
         resulttypesum = value1.type;
      }
   }

   resulttypesum = resulttypesum || "n";

   // Return the appropriate result based on the function name
   switch (fname) {
      case "DSUM":
         PushOperand(resulttypesum, sum);
         break;

      case "DPRODUCT": // may handle cases with text differently than some other spreadsheets
         PushOperand(resulttypesum, product);
         break;

      case "DMIN":
         PushOperand(resulttypesum, minval || 0);
         break;

      case "DMAX":
         PushOperand(resulttypesum, maxval || 0);
         break;

      case "DCOUNT":
         PushOperand("n", count);
         break;

      case "DCOUNTA":
         PushOperand("n", counta);
         break;

      case "DAVERAGE":
         if (count > 0) {
            PushOperand(resulttypesum, sum/count);
         } else {
            PushOperand("e#DIV/0!", 0);
         }
         break;

      case "DSTDEV":
         if (count > 1) {
            PushOperand(resulttypesum, Math.sqrt(sk / (count - 1))); // sk is never negative according to Knuth
         } else {
            PushOperand("e#DIV/0!", 0);
         }
         break;

      case "DSTDEVP":
         if (count > 1) {
            PushOperand(resulttypesum, Math.sqrt(sk / count));
         } else {
            PushOperand("e#DIV/0!", 0);
         }
         break;

      case "DVAR":
         if (count > 1) {
            PushOperand(resulttypesum, sk / (count - 1));
         } else {
            PushOperand("e#DIV/0!", 0);
         }
         break;

      case "DVARP":
         if (count > 1) {
            PushOperand(resulttypesum, sk / count);
         } else {
            PushOperand("e#DIV/0!", 0);
         }
         break;

      case "DGET":
         if (count == 1) {
            PushOperand(resulttypesum, sum);
         } else if (count == 0) {
            PushOperand("e#VALUE!", 0);
         } else {
            PushOperand("e#NUM!", 0);
         }
         break;
   }

   return;
};
/**
 * @description Function list registrations for database statistical functions
 * Each entry follows the format: [function_handler, arg_count, arg_def, func_def, func_class]
 * - arg_count: 3 means exactly three arguments (database range, field name, criteria range)
 * - arg_def: "dfunc" refers to database function arguments
 * - func_class: "stat" indicates statistical function category
 */

SocialCalc.Formula.FunctionList["DAVERAGE"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DCOUNT"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DCOUNTA"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DGET"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DMAX"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DMIN"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DPRODUCT"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DSTDEV"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DSTDEVP"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DSUM"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DVAR"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];
SocialCalc.Formula.FunctionList["DVARP"] = [SocialCalc.Formula.DSeriesFunctions, 3, "dfunc", "", "stat"];

/**
 * @function FieldToColnum
 * @memberof SocialCalc.Formula
 * @description Converts a field name to a column number within a database range
 * 
 * If fieldname is a number, uses it directly (if valid), otherwise looks up the string
 * in the header row cells to find the matching field number.
 * 
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @param {number} col1num - The first column number of the database range
 * @param {number} ncols - The number of columns in the database range
 * @param {number} row1num - The header row number of the database range
 * @param {*} fieldname - The field name to look up (string or number)
 * @param {string} fieldtype - The type of the fieldname parameter
 * @returns {number} Column number (1-based) if found, 0 if not found
 * 
 * @example
 * const colNum = SocialCalc.Formula.FieldToColnum(sheet, 1, 5, 1, "Sales", "t");
 * // Returns: 3 if "Sales" is found in the 3rd column header
 */
SocialCalc.Formula.FieldToColnum = function(sheet, col1num, ncols, row1num, fieldname, fieldtype) {
   let colnum, cell, value;

   if (fieldtype.charAt(0) == "n") { // number - return it if legal
      colnum = fieldname - 0; // make sure a number
      if (colnum <= 0 || colnum > ncols) {
         return 0;
      }
      return Math.floor(colnum);
   }

   if (fieldtype.charAt(0) != "t") { // must be text otherwise
      return 0;
   }

   fieldname = fieldname ? fieldname.toLowerCase() : "";

   for (colnum = 0; colnum < ncols; colnum++) { // look through column headers for a match
      cell = sheet.GetAssuredCell(SocialCalc.crToCoord(col1num+colnum, row1num));
      value = cell.datavalue;
      value = (value+"").toLowerCase(); // ignore case
      if (value == fieldname) { // match
         return colnum+1;
      }         
   }
   return 0; // looked at all and no match
};

/**
 * @function LookupFunctions
 * @memberof SocialCalc.Formula
 * @description Implements HLOOKUP, VLOOKUP, and MATCH functions for table lookups
 * 
 * These functions search for values in tables and return corresponding values or positions:
 * - HLOOKUP: Horizontal lookup in the first row, returns value from specified row
 * - VLOOKUP: Vertical lookup in the first column, returns value from specified column  
 * - MATCH: Returns the position of a value in a range
 * 
 * Supports both exact match and range lookup modes.
 * 
 * @param {string} fname - The specific lookup function name (HLOOKUP, VLOOKUP, or MATCH)
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction for lookup functions
 * // =VLOOKUP("Product A", A1:D10, 3, FALSE) - exact match lookup
 * // =HLOOKUP(100, A1:Z3, 2, TRUE) - range lookup with approximate match
 * SocialCalc.Formula.LookupFunctions("VLOOKUP", operand, foperand, sheet);
 */
SocialCalc.Formula.LookupFunctions = function(fname, operand, foperand, sheet) {
   let lookupvalue, range, offset, rangelookup, offsetvalue, rangeinfo;
   let c, r, cincr, rincr, previousOK, csave, rsave, cell, value, valuetype, cr;

   const scf = SocialCalc.Formula;
   const operand_value_and_type = scf.OperandValueAndType;
   const lookup_result_type = scf.LookupResultType;
   const typelookupplus = scf.TypeLookupTable.plus;

   /**
    * @function PushOperand
    * @description Helper function to push operand onto stack
    * @param {string} t - Type of operand
    * @param {*} v - Value of operand
    */
   const PushOperand = function(t, v) {operand.push({type: t, value: v});};

   // Get the lookup value and convert strings to lowercase for case-insensitive comparison
   lookupvalue = operand_value_and_type(sheet, foperand);
   if (typeof lookupvalue.value == "string") {
      lookupvalue.value = lookupvalue.value.toLowerCase();
   }

   // Get the range to search in
   range = scf.TopOfStackValueAndType(sheet, foperand);

   // Set default range lookup behavior and process optional arguments
   rangelookup = 1; // default to true or 1
   if (fname == "MATCH") {
      if (foperand.length) {
         rangelookup = scf.OperandAsNumber(sheet, foperand);
         if (rangelookup.type.charAt(0) != "n") {
            PushOperand("e#VALUE!", 0);
            return;
         }
         if (foperand.length) {
            scf.FunctionArgsError(fname, operand);
            return 0;
         }
         rangelookup = rangelookup.value - 0;
      }
   } else {
      // For HLOOKUP/VLOOKUP, get the offset (row/column index)
      offsetvalue = scf.OperandAsNumber(sheet, foperand);
      if (offsetvalue.type.charAt(0) != "n") {
         PushOperand("e#VALUE!", 0);
         return;
      }
      offsetvalue = Math.floor(offsetvalue.value);
      
      if (foperand.length) {
         rangelookup = scf.OperandAsNumber(sheet, foperand);
         if (rangelookup.type.charAt(0) != "n") {
            PushOperand("e#VALUE!", 0);
            return;
         }
         if (foperand.length) {
            scf.FunctionArgsError(fname, operand);
            return 0;
         }
         rangelookup = rangelookup.value ? 1 : 0; // convert to 1 or 0
      }
   }
   
   // Normalize lookup value type and ensure numbers are numeric
   lookupvalue.type = lookupvalue.type.charAt(0); // only deal with general type
   if (lookupvalue.type == "n") { // if number, make sure a number
      lookupvalue.value = lookupvalue.value - 0;
   }

   // Validate range argument
   if (range.type != "range") {
      scf.FunctionArgsError(fname, operand);
      return 0;
   }

   rangeinfo = scf.DecodeRangeParts(sheet, range.value, range.type);
   if (!rangeinfo) {
      PushOperand("e#REF!", 0);
      return;
   }

   // Set up search direction and validate offset bounds
   c = 0;
   r = 0;
   cincr = 0;
   rincr = 0;
   
   if (fname == "HLOOKUP") {
      cincr = 1; // search horizontally
      if (offsetvalue > rangeinfo.nrows) {
         PushOperand("e#REF!", 0);
         return;
      }
   } else if (fname == "VLOOKUP") {
      rincr = 1; // search vertically
      if (offsetvalue > rangeinfo.ncols) {
         PushOperand("e#REF!", 0);
         return;
      }
   } else if (fname == "MATCH") {
      // Determine search direction based on range shape
      if (rangeinfo.ncols > 1) {
         if (rangeinfo.nrows > 1) {
            PushOperand("e#N/A", 0);
            return;
         }
         cincr = 1; // horizontal range
      } else {
         rincr = 1; // vertical range
      }
   } else {
      scf.FunctionArgsError(fname, operand);
      return 0;
   }
   
   if (offsetvalue < 1 && fname != "MATCH") {
      PushOperand("e#VALUE!", 0);
      return 0;
   }

   /** @type {number|undefined} Tracks previous acceptable match for range lookups */
   previousOK; // if 1, previous test was <. If 2, also this one wasn't

   // Main search loop
   while (1) {
      cr = SocialCalc.crToCoord(rangeinfo.col1num + c, rangeinfo.row1num + r);
      cell = rangeinfo.sheetdata.GetAssuredCell(cr);
      value = cell.datavalue;
      valuetype = cell.valuetype ? cell.valuetype.charAt(0) : "b"; // only deal with general types
      
      if (valuetype == "n") {
         value = value - 0; // make sure number
      }
      
      if (rangelookup) { // rangelookup type 1 or -1: look for within brackets for matches
         if (lookupvalue.type == "n" && valuetype == "n") {
            if (lookupvalue.value == value) { // exact match
               break;
            }
            if ((rangelookup > 0 && lookupvalue.value > value)
                || (rangelookup < 0 && lookupvalue.value < value)) { // possible match: wait and see
               previousOK = 1;
               csave = c; // remember col and row of last OK
               rsave = r;
            } else if (previousOK) { // last one was OK, this one isn't
               previousOK = 2;
               break;
            }
         } else if (lookupvalue.type == "t" && valuetype == "t") {
            value = typeof value == "string" ? value.toLowerCase() : "";
            if (lookupvalue.value == value) { // exact match
               break;
            }
            if ((rangelookup > 0 && lookupvalue.value > value)
                || (rangelookup < 0 && lookupvalue.value < value)) { // possible match: wait and see
               previousOK = 1;
               csave = c;
               rsave = r;
            } else if (previousOK) { // last one was OK, this one isn't
               previousOK = 2;
               break;
            }
         }
      } else { // exact value matches only
         if (lookupvalue.type == "n" && valuetype == "n") {
            if (lookupvalue.value == value) { // exact match
               break;
            }
         } else if (lookupvalue.type == "t" && valuetype == "t") {
            value = typeof value == "string" ? value.toLowerCase() : "";
            if (lookupvalue.value == value) { // exact match
               break;
            }
         }
      }

      // Move to next position
      r += rincr;
      c += cincr;
      
      if (r >= rangeinfo.nrows || c >= rangeinfo.ncols) { // end of range to check, no exact match
         if (previousOK) { // at least one could have been OK
            previousOK = 2;
            break;
         }
         PushOperand("e#N/A", 0);
         return;
      }
   }

   // Use the best available match
   if (previousOK == 2) { // back to last OK
      r = rsave;
      c = csave;
   }

   // Return appropriate result based on function type
   if (fname == "MATCH") {
      value = c + r + 1; // only one may be <> 0
      valuetype = "n";
   } else {
      // Get the value from the offset position
      cr = SocialCalc.crToCoord(
         rangeinfo.col1num + c + (fname == "VLOOKUP" ? offsetvalue-1 : 0), 
         rangeinfo.row1num + r + (fname == "HLOOKUP" ? offsetvalue-1 : 0)
      );
      cell = rangeinfo.sheetdata.GetAssuredCell(cr);
      value = cell.datavalue;
      valuetype = cell.valuetype;
   }
   
   PushOperand(valuetype, value);

   return;
};

/**
 * @description Function list registrations for lookup functions
 * Each entry follows the format: [function_handler, arg_count, arg_def, func_def, func_class]
 * - arg_count: negative values indicate minimum number of arguments (-2 means 2+, -3 means 3+)
 * - func_class: "lookup" indicates lookup/reference function category
 */
SocialCalc.Formula.FunctionList["HLOOKUP"] = [SocialCalc.Formula.LookupFunctions, -3, "hlookup", "", "lookup"];
SocialCalc.Formula.FunctionList["MATCH"] = [SocialCalc.Formula.LookupFunctions, -2, "match", "", "lookup"];
SocialCalc.Formula.FunctionList["VLOOKUP"] = [SocialCalc.Formula.LookupFunctions, -3, "vlookup", "", "lookup"];
/**
 * @function IndexFunction
 * @memberof SocialCalc.Formula
 * @description Implements the INDEX function for retrieving values from ranges by position
 * 
 * INDEX(range, rownum, colnum) returns a reference to a cell or range at the specified position.
 * - If both row and column are specified, returns the cell at that intersection
 * - If only row is specified (or only column for single-row ranges), returns the entire row/column
 * - If neither is specified, returns the entire range
 * 
 * @param {string} fname - The function name ("INDEX")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =INDEX(A1:C5, 2, 3) returns reference to C2
 * // =INDEX(A1:A10, 5) returns reference to A5 (single column)
 * SocialCalc.Formula.IndexFunction("INDEX", operand, foperand, sheet);
 */
SocialCalc.Formula.IndexFunction = function(fname, operand, foperand, sheet) {
   let range, sheetname, indexinfo, rowindex, colindex, result, resulttype;

   const scf = SocialCalc.Formula;

   /**
    * @function PushOperand
    * @description Helper function to push operand onto stack
    * @param {string} t - Type of operand
    * @param {*} v - Value of operand
    */
   const PushOperand = function(t, v) {operand.push({type: t, value: v});};

   // Get the range argument
   range = scf.TopOfStackValueAndType(sheet, foperand); // get range
   if (range.type != "range") {
      scf.FunctionArgsError(fname, operand);
      return 0;
   }
   
   indexinfo = scf.DecodeRangeParts(sheet, range.value, range.type);
   if (indexinfo.sheetname) {
      sheetname = "!" + indexinfo.sheetname;
   } else {
      sheetname = "";
   }

   // Initialize row and column indices
   rowindex = {value: 0};
   colindex = {value: 0};

   if (foperand.length) { // look for row number
      rowindex = scf.OperandAsNumber(sheet, foperand);
      if (rowindex.type.charAt(0) != "n" || rowindex.value < 0) {
         PushOperand("e#VALUE!", 0);
         return;
      }
      
      if (foperand.length) { // look for col number
         colindex = scf.OperandAsNumber(sheet, foperand);
         if (colindex.type.charAt(0) != "n" || colindex.value < 0) {
            PushOperand("e#VALUE!", 0);
            return;
         }
         if (foperand.length) {
            scf.FunctionArgsError(fname, operand);
            return 0;
         }
      } else { // col number missing
         if (indexinfo.nrows == 1) { // if only one row, then rowindex is really colindex
            colindex.value = rowindex.value;
            rowindex.value = 0;
         }
      }
   }

   // Validate indices are within range bounds
   if (rowindex.value > indexinfo.nrows || colindex.value > indexinfo.ncols) {
      PushOperand("e#REF!", 0);
      return;
   }

   // Generate the appropriate result based on the indices provided
   if (rowindex.value == 0) {
      if (colindex.value == 0) {
         // Return entire range
         if (indexinfo.nrows == 1 && indexinfo.ncols == 1) {
            result = SocialCalc.crToCoord(indexinfo.col1num, indexinfo.row1num) + sheetname;
            resulttype = "coord";
         } else {
            result = SocialCalc.crToCoord(indexinfo.col1num, indexinfo.row1num) + sheetname + "|" +
                     SocialCalc.crToCoord(indexinfo.col1num+indexinfo.ncols-1, indexinfo.row1num+indexinfo.nrows-1) + 
                     "|";
            resulttype = "range";
         }
      } else {
         // Return entire column
         if (indexinfo.nrows == 1) {
            result = SocialCalc.crToCoord(indexinfo.col1num+colindex.value-1, indexinfo.row1num) + sheetname;
            resulttype = "coord";
         } else {
            result = SocialCalc.crToCoord(indexinfo.col1num+colindex.value-1, indexinfo.row1num) + sheetname + "|" +
                     SocialCalc.crToCoord(indexinfo.col1num+colindex.value-1, indexinfo.row1num+indexinfo.nrows-1) +
                     "|";
            resulttype = "range";
         }
      }
   } else {
      if (colindex.value == 0) {
         // Return entire row
         if (indexinfo.ncols == 1) {
            result = SocialCalc.crToCoord(indexinfo.col1num, indexinfo.row1num+rowindex.value-1) + sheetname;
            resulttype = "coord";
         } else {
            result = SocialCalc.crToCoord(indexinfo.col1num, indexinfo.row1num+rowindex.value-1) + sheetname + "|" +
                     SocialCalc.crToCoord(indexinfo.col1num+indexinfo.ncols-1, indexinfo.row1num+rowindex.value-1) +
                     "|";
            resulttype = "range";
         }
      } else {
         // Return specific cell
         result = SocialCalc.crToCoord(indexinfo.col1num+colindex.value-1, indexinfo.row1num+rowindex.value-1) + sheetname;
         resulttype = "coord";
      }
   }

   PushOperand(resulttype, result);

   return;
};

SocialCalc.Formula.FunctionList["INDEX"] = [SocialCalc.Formula.IndexFunction, -1, "index", "", "lookup"];

/**
 * @function CountifSumifFunctions
 * @memberof SocialCalc.Formula
 * @description Implements COUNTIF and SUMIF functions for conditional counting and summing
 * 
 * - COUNTIF(range, criteria): Counts cells in range that meet criteria
 * - SUMIF(range, criteria, [sum_range]): Sums values where criteria are met
 *   If sum_range is omitted, sums the values in the criteria range
 * 
 * @param {string} fname - The function name ("COUNTIF" or "SUMIF")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =COUNTIF(A1:A10, ">5") counts cells > 5
 * // =SUMIF(A1:A10, ">5", B1:B10) sums B values where A > 5
 * SocialCalc.Formula.CountifSumifFunctions("SUMIF", operand, foperand, sheet);
 */
SocialCalc.Formula.CountifSumifFunctions = function(fname, operand, foperand, sheet) {
   let range, criteria, sumrange, f2operand, result, resulttype, value1, value2;
   let sum = 0;
   let resulttypesum = "";
   let count = 0;

   const scf = SocialCalc.Formula;
   const operand_value_and_type = scf.OperandValueAndType;
   const lookup_result_type = scf.LookupResultType;
   const typelookupplus = scf.TypeLookupTable.plus;

   /**
    * @function PushOperand
    * @description Helper function to push operand onto stack
    * @param {string} t - Type of operand
    * @param {*} v - Value of operand
    */
   const PushOperand = function(t, v) {operand.push({type: t, value: v});};

   // Get function arguments
   range = scf.TopOfStackValueAndType(sheet, foperand); // get range or coord
   criteria = scf.OperandAsText(sheet, foperand); // get criteria
   
   if (fname == "SUMIF") {
      if (foperand.length == 1) { // three arg form of SUMIF
         sumrange = scf.TopOfStackValueAndType(sheet, foperand);
      } else if (foperand.length == 0) { // two arg form
         sumrange = {value: range.value, type: range.type};
      } else {
         scf.FunctionArgsError(fname, operand);
         return 0;
      }
   } else {
      sumrange = {value: range.value, type: range.type};
   }

   // Process criteria value
   if (criteria.type.charAt(0) == "n") {
      criteria.value = criteria.value + ""; // make text
   } else if (criteria.type.charAt(0) == "e") { // error
      criteria.value = null;
   } else if (criteria.type.charAt(0) == "b") { // blank here is undefined
      criteria.value = null;
   }

   // Validate range arguments
   if (range.type != "coord" && range.type != "range") {
      scf.FunctionArgsError(fname, operand);
      return 0;
   }

   if (fname == "SUMIF" && sumrange.type != "coord" && sumrange.type != "range") {
      scf.FunctionArgsError(fname, operand);
      return 0;
   }

   // Set up operand stacks for parallel processing of criteria and sum ranges
   foperand.push(range);
   f2operand = []; // to allow for 3 arg form
   f2operand.push(sumrange);

   // Process each cell in the ranges
   while (foperand.length) {
      value1 = operand_value_and_type(sheet, foperand); // criteria range value
      value2 = operand_value_and_type(sheet, f2operand); // sum range value

      if (!scf.TestCriteria(value1.value, value1.type, criteria.value)) {
         continue; // doesn't meet criteria
      }

      count += 1;

      if (value2.type.charAt(0) == "n") {
         sum += value2.value-0;
         resulttypesum = lookup_result_type(value2.type, resulttypesum || value2.type, typelookupplus);
      } else if (value2.type.charAt(0) == "e" && resulttypesum.charAt(0) != "e") {
         resulttypesum = value2.type;
      }
   }

   resulttypesum = resulttypesum || "n";

   // Return appropriate result
   if (fname == "SUMIF") {
      PushOperand(resulttypesum, sum);
   } else if (fname == "COUNTIF") {
      PushOperand("n", count);
   }

   return;
};

SocialCalc.Formula.FunctionList["COUNTIF"] = [SocialCalc.Formula.CountifSumifFunctions, 2, "rangec", "", "stat"];
SocialCalc.Formula.FunctionList["SUMIF"] = [SocialCalc.Formula.CountifSumifFunctions, -2, "sumif", "", "stat"];

/**
 * @function IfFunction
 * @memberof SocialCalc.Formula
 * @description Implements the IF function for conditional value selection
 * 
 * IF(condition, true_value, false_value) returns true_value if condition is true,
 * false_value otherwise. The condition must be numeric or boolean.
 * 
 * @param {string} fname - The function name ("IF")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {null} Always returns null (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =IF(A1>5, "High", "Low") returns "High" if A1>5, "Low" otherwise
 * SocialCalc.Formula.IfFunction("IF", operand, foperand, sheet);
 */
SocialCalc.Formula.IfFunction = function(fname, operand, foperand, sheet) {
   let cond, t;

   cond = SocialCalc.Formula.OperandValueAndType(sheet, foperand);
   t = cond.type.charAt(0);
   
   if (t != "n" && t != "b") {
      operand.push({type: "e#VALUE!", value: 0});
      return;
   }

   // Select the appropriate value based on condition
   if (!cond.value) foperand.pop(); // remove true_value if condition is false
   operand.push(foperand.pop()); // push the selected value
   if (cond.value) foperand.pop(); // remove false_value if condition is true

   return null;
};

// Add to function list
SocialCalc.Formula.FunctionList["IF"] = [SocialCalc.Formula.IfFunction, 3, "iffunc", "", "test"];

/**
 * @function DateFunction
 * @memberof SocialCalc.Formula
 * @description Implements the DATE function for creating date values
 * 
 * DATE(year, month, day) returns a date serial number for the specified date.
 * Uses the Gregorian to Julian calendar conversion system.
 * 
 * @param {string} fname - The function name ("DATE")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =DATE(2023, 12, 25) returns the serial number for December 25, 2023
 * SocialCalc.Formula.DateFunction("DATE", operand, foperand, sheet);
 */
SocialCalc.Formula.DateFunction = function(fname, operand, foperand, sheet) {
   const scf = SocialCalc.Formula;
   let result = 0;
   
   const year = scf.OperandAsNumber(sheet, foperand);
   const month = scf.OperandAsNumber(sheet, foperand);
   const day = scf.OperandAsNumber(sheet, foperand);
   
   let resulttype = scf.LookupResultType(year.type, month.type, scf.TypeLookupTable.twoargnumeric);
   resulttype = scf.LookupResultType(resulttype, day.type, scf.TypeLookupTable.twoargnumeric);
   
   if (resulttype.charAt(0) == "n") {
      result = SocialCalc.FormatNumber.convert_date_gregorian_to_julian(
                  Math.floor(year.value), Math.floor(month.value), Math.floor(day.value)
                  ) - SocialCalc.FormatNumber.datevalues.julian_offset;
      resulttype = "nd";
   }
   
   scf.PushOperand(operand, resulttype, result);
   return;
};

SocialCalc.Formula.FunctionList["DATE"] = [SocialCalc.Formula.DateFunction, 3, "date", "", "datetime"];

/**
 * @function TimeFunction
 * @memberof SocialCalc.Formula
 * @description Implements the TIME function for creating time values
 * 
 * TIME(hour, minute, second) returns a time serial number (fraction of a day)
 * for the specified time components.
 * 
 * @param {string} fname - The function name ("TIME")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =TIME(14, 30, 0) returns 0.604167 (2:30 PM as fraction of day)
 * SocialCalc.Formula.TimeFunction("TIME", operand, foperand, sheet);
 */
SocialCalc.Formula.TimeFunction = function(fname, operand, foperand, sheet) {
   const scf = SocialCalc.Formula;
   let result = 0;
   
   const hours = scf.OperandAsNumber(sheet, foperand);
   const minutes = scf.OperandAsNumber(sheet, foperand);
   const seconds = scf.OperandAsNumber(sheet, foperand);
   
   let resulttype = scf.LookupResultType(hours.type, minutes.type, scf.TypeLookupTable.twoargnumeric);
   resulttype = scf.LookupResultType(resulttype, seconds.type, scf.TypeLookupTable.twoargnumeric);
   
   if (resulttype.charAt(0) == "n") {
      result = ((hours.value * 60 * 60) + (minutes.value * 60) + seconds.value) / (24*60*60);
      resulttype = "nt";
   }
   
   scf.PushOperand(operand, resulttype, result);
   return;
};

SocialCalc.Formula.FunctionList["TIME"] = [SocialCalc.Formula.TimeFunction, 3, "hms", "", "datetime"];

/**
 * @function DMYFunctions
 * @memberof SocialCalc.Formula
 * @description Implements date extraction functions (DAY, MONTH, YEAR, WEEKDAY)
 * 
 * - DAY(date): Returns the day of the month (1-31)
 * - MONTH(date): Returns the month (1-12)
 * - YEAR(date): Returns the year
 * - WEEKDAY(date, [type]): Returns day of week (1-7), with optional type parameter
 * 
 * @param {string} fname - The function name ("DAY", "MONTH", "YEAR", or "WEEKDAY")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =DAY(DATE(2023,12,25)) returns 25
 * // =WEEKDAY(DATE(2023,12,25), 1) returns day of week for Christmas 2023
 * SocialCalc.Formula.DMYFunctions("DAY", operand, foperand, sheet);
 */
SocialCalc.Formula.DMYFunctions = function(fname, operand, foperand, sheet) {
   let ymd, dtype, doffset;
   const scf = SocialCalc.Formula;
   let result = 0;

   const datevalue = scf.OperandAsNumber(sheet, foperand);
   const resulttype = scf.LookupResultType(datevalue.type, datevalue.type, scf.TypeLookupTable.oneargnumeric);

   if (resulttype.charAt(0) == "n") {
      ymd = SocialCalc.FormatNumber.convert_date_julian_to_gregorian(
               Math.floor(datevalue.value + SocialCalc.FormatNumber.datevalues.julian_offset));
      
      switch (fname) {
         case "DAY":
            result = ymd.day;
            break;

         case "MONTH":
            result = ymd.month;
            break;

         case "YEAR":
            result = ymd.year;
            break;

         case "WEEKDAY":
            dtype = {value: 1};
            if (foperand.length) { // get type if present
               dtype = scf.OperandAsNumber(sheet, foperand);
               if (dtype.type.charAt(0) != "n" || dtype.value < 1 || dtype.value > 3) {
                  scf.PushOperand(operand, "e#VALUE!", 0);
                  return;
               }
               if (foperand.length) { // extra args
                  scf.FunctionArgsError(fname, operand);
                  return;
               }
            }
            doffset = 6;
            if (dtype.value > 1) {
               doffset -= 1;
            }
            result = Math.floor(datevalue.value+doffset) % 7 + (dtype.value < 3 ? 1 : 0);
            break;
      }
   }

   scf.PushOperand(operand, resulttype, result);
   return;
};

SocialCalc.Formula.FunctionList["DAY"] = [SocialCalc.Formula.DMYFunctions, 1, "v", "", "datetime"];
SocialCalc.Formula.FunctionList["MONTH"] = [SocialCalc.Formula.DMYFunctions, 1, "v", "", "datetime"];
SocialCalc.Formula.FunctionList["YEAR"] = [SocialCalc.Formula.DMYFunctions, 1, "v", "", "datetime"];
SocialCalc.Formula.FunctionList["WEEKDAY"] = [SocialCalc.Formula.DMYFunctions, -1, "weekday", "", "datetime"];
/**
 * @function HMSFunctions
 * @memberof SocialCalc.Formula
 * @description Implements time extraction functions (HOUR, MINUTE, SECOND)
 * 
 * These functions extract time components from datetime values:
 * - HOUR(datetime): Returns the hour component (0-23)
 * - MINUTE(datetime): Returns the minute component (0-59)
 * - SECOND(datetime): Returns the second component (0-59)
 * 
 * The datetime value is expected to be a fraction of a day, where the integer
 * part represents the date and the fractional part represents the time.
 * 
 * @param {string} fname - The function name ("HOUR", "MINUTE", or "SECOND")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =HOUR(0.5) returns 12 (noon)
 * // =MINUTE(TIME(14,30,45)) returns 30
 * // =SECOND(TIME(14,30,45)) returns 45
 * SocialCalc.Formula.HMSFunctions("HOUR", operand, foperand, sheet);
 */
SocialCalc.Formula.HMSFunctions = function(fname, operand, foperand, sheet) {
   let hours, minutes, seconds, fraction;
   const scf = SocialCalc.Formula;
   let result = 0;

   const datetime = scf.OperandAsNumber(sheet, foperand);
   const resulttype = scf.LookupResultType(datetime.type, datetime.type, scf.TypeLookupTable.oneargnumeric);

   if (resulttype.charAt(0) == "n") {
      if (datetime.value < 0) {
         scf.PushOperand(operand, "e#NUM!", 0); // must be non-negative
         return;
      }
      
      // Extract time components from fractional part of day
      fraction = datetime.value - Math.floor(datetime.value); // fraction of a day
      fraction *= 24; // convert to hours
      hours = Math.floor(fraction);
      fraction -= Math.floor(fraction);
      fraction *= 60; // convert to minutes
      minutes = Math.floor(fraction);
      fraction -= Math.floor(fraction);
      fraction *= 60; // convert to seconds
      seconds = Math.floor(fraction + (datetime.value >= 0 ? 0.5: -0.5)); // round to nearest second
      
      if (fname == "HOUR") {
         result = hours;
      } else if (fname == "MINUTE") {
         result = minutes;
      } else if (fname == "SECOND") {
         result = seconds;
      }
   }

   scf.PushOperand(operand, resulttype, result);
   return;
};

SocialCalc.Formula.FunctionList["HOUR"] = [SocialCalc.Formula.HMSFunctions, 1, "v", "", "datetime"];
SocialCalc.Formula.FunctionList["MINUTE"] = [SocialCalc.Formula.HMSFunctions, 1, "v", "", "datetime"];
SocialCalc.Formula.FunctionList["SECOND"] = [SocialCalc.Formula.HMSFunctions, 1, "v", "", "datetime"];

/**
 * @function ExactFunction
 * @memberof SocialCalc.Formula
 * @description Implements the EXACT function for case-sensitive comparison
 * 
 * EXACT(value1, value2) compares two values and returns TRUE (1) if they are exactly equal,
 * FALSE (0) otherwise. The comparison is case-sensitive for text values.
 * 
 * Type conversion rules:
 * - Text vs Text: Case-sensitive string comparison
 * - Number vs Number: Numeric comparison
 * - Text vs Number: Convert number to text for comparison
 * - Blank vs Text: TRUE only if text is empty string
 * - Blank vs Blank: Always TRUE
 * - Any vs Error: Returns the error
 * 
 * @param {string} fname - The function name ("EXACT")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =EXACT("Hello", "hello") returns 0 (FALSE - case sensitive)
 * // =EXACT("Hello", "Hello") returns 1 (TRUE - exact match)
 * // =EXACT(123, "123") returns 1 (TRUE - number converted to text)
 * SocialCalc.Formula.ExactFunction("EXACT", operand, foperand, sheet);
 */
SocialCalc.Formula.ExactFunction = function(fname, operand, foperand, sheet) {
   const scf = SocialCalc.Formula;
   let result = 0;
   let resulttype = "nl";

   const value1 = scf.OperandValueAndType(sheet, foperand);
   const v1type = value1.type.charAt(0);
   const value2 = scf.OperandValueAndType(sheet, foperand);
   const v2type = value2.type.charAt(0);

   if (v1type == "t") { // First value is text
      if (v2type == "t") {
         result = value1.value == value2.value ? 1 : 0; // Case-sensitive text comparison
      } else if (v2type == "b") {
         result = value1.value.length ? 0 : 1; // Text vs blank
      } else if (v2type == "n") {
         result = value1.value == value2.value+"" ? 1 : 0; // Text vs number (convert number to string)
      } else if (v2type == "e") {
         result = value2.value;
         resulttype = value2.type; // Propagate error
      } else {
         result = 0;
      }
   } else if (v1type == "n") { // First value is number
      if (v2type == "n") {
         result = value1.value-0 == value2.value-0 ? 1 : 0; // Numeric comparison
      } else if (v2type == "b") {
         result = 0; // Number vs blank is never equal
      } else if (v2type == "t") {
         result = value1.value+"" == value2.value ? 1 : 0; // Number vs text (convert number to string)
      } else if (v2type == "e") {
         result = value2.value;
         resulttype = value2.type; // Propagate error
      } else {
         result = 0;
      }
   } else if (v1type == "b") { // First value is blank
      if (v2type == "t") {
         result = value2.value.length ? 0 : 1; // Blank vs text
      } else if (v2type == "b") {
         result = 1; // Blank vs blank is always equal
      } else if (v2type == "n") {
         result = 0; // Blank vs number is never equal
      } else if (v2type == "e") {
         result = value2.value;
         resulttype = value2.type; // Propagate error
      } else {
         result = 0;
      }
   } else if (v1type == "e") { // First value is error
      result = value1.value;
      resulttype = value1.type; // Propagate error
   }

   scf.PushOperand(operand, resulttype, result);
   return;
};

SocialCalc.Formula.FunctionList["EXACT"] = [SocialCalc.Formula.ExactFunction, 2, "", "", "text"];

/**
 * @type {Object<string, Array<number>>}
 * @description Argument type definitions for string functions
 * 
 * Each function has an array specifying argument types:
 * - 1: Text argument (converted to string)
 * - 0: Numeric argument (converted to number)
 * - -1: Accept any type as-is
 * 
 * Text values are manipulated as UTF-8, converting from and back to byte strings.
 */
SocialCalc.Formula.ArgList = {
   FIND: [1, 1, 0],        // key (text), string (text), start (numeric)
   LEFT: [1, 0],           // string (text), length (numeric)
   LEN: [1],               // string (text)
   LOWER: [1],             // string (text)
   MID: [1, 0, 0],         // string (text), start (numeric), length (numeric)
   PROPER: [1],            // string (text)
   REPLACE: [1, 0, 0, 1],  // string (text), start (numeric), length (numeric), new (text)
   REPT: [1, 0],           // string (text), count (numeric)
   RIGHT: [1, 0],          // string (text), length (numeric)
   SUBSTITUTE: [1, 1, 1, 0], // string (text), old (text), new (text), which (numeric)
   TRIM: [1],              // string (text)
   UPPER: [1]              // string (text)
};

/**
 * @function StringFunctions
 * @memberof SocialCalc.Formula
 * @description Implements various string manipulation functions
 * 
 * Handles multiple text functions:
 * - FIND(key, string, [start]): Find position of key in string
 * - LEFT(string, [length]): Extract leftmost characters
 * - LEN(string): Get string length
 * - LOWER(string): Convert to lowercase
 * - MID(string, start, length): Extract middle portion
 * - PROPER(string): Convert to proper case (title case)
 * - REPLACE(string, start, length, new): Replace portion of string
 * - REPT(string, count): Repeat string multiple times
 * - RIGHT(string, [length]): Extract rightmost characters
 * - SUBSTITUTE(string, old, new, [which]): Replace occurrences of substring
 * - TRIM(string): Remove leading/trailing spaces
 * - UPPER(string): Convert to uppercase
 * 
 * @param {string} fname - The specific string function name
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction for string functions
 * // =LEFT("Hello World", 5) returns "Hello"
 * // =FIND("o", "Hello World", 1) returns 5
 * // =UPPER("hello") returns "HELLO"
 * SocialCalc.Formula.StringFunctions("LEFT", operand, foperand, sheet);
 */
SocialCalc.Formula.StringFunctions = function(fname, operand, foperand, sheet) {
   let i, value, offset, len, start, count;
   const scf = SocialCalc.Formula;
   let result = 0;
   let resulttype = "e#VALUE!";

   const numargs = foperand.length;
   const argdef = scf.ArgList[fname];
   const operand_value = [];
   const operand_type = [];

   // Process and validate all arguments
   for (i = 1; i <= numargs; i++) { // go through each arg, get value and type, and check for errors
      if (i > argdef.length) { // too many args
         scf.FunctionArgsError(fname, operand);
         return;
      }
      
      if (argdef[i-1] == 0) {
         value = scf.OperandAsNumber(sheet, foperand);
      } else if (argdef[i-1] == 1) {
         value = scf.OperandAsText(sheet, foperand);
      } else if (argdef[i-1] == -1) {
         value = scf.OperandValueAndType(sheet, foperand);
      }
      
      operand_value[i] = value.value;
      operand_type[i] = value.type;
      
      if (value.type.charAt(0) == "e") {
         scf.PushOperand(operand, value.type, result);
         return;
      }
   }

   // Execute the specific string function
   switch (fname) {
      case "FIND":
         offset = operand_type[3] ? operand_value[3]-1 : 0;
         if (offset < 0) {
            result = "Start is before string"; // !! not displayed, no need to translate
         } else {
            result = operand_value[2].indexOf(operand_value[1], offset); // (null string matches first char)
            if (result >= 0) {
               result += 1; // Convert to 1-based indexing
               resulttype = "n";
            } else {
               result = "Not found"; // !! not displayed, error is e#VALUE!
            }
         }
         break;

      case "LEFT":
         len = operand_type[2] ? operand_value[2]-0 : 1;
         if (len < 0) {
            result = "Negative length";
         } else {
            result = operand_value[1].substring(0, len);
            resulttype = "t";
         }
         break;

      case "LEN":
         result = operand_value[1].length;
         resulttype = "n";
         break;

      case "LOWER":
         result = operand_value[1].toLowerCase();
         resulttype = "t";
         break;

      case "MID":
         start = operand_value[2]-0;
         len = operand_value[3]-0;
         if (len < 1 || start < 1) {
            result = "Bad arguments";
         } else {
            result = operand_value[1].substring(start-1, start+len-1);
            resulttype = "t";
         }
         break;

      case "PROPER":
         result = operand_value[1].replace(/\b\w+\b/g, function(word) {
                     return word.substring(0,1).toUpperCase() + 
                        word.substring(1);
                     }); // uppercase first character of words (see JavaScript, Flanagan, 5th edition, page 704)
         resulttype = "t";
         break;

      case "REPLACE":
         start = operand_value[2]-0;
         len = operand_value[3]-0;
         if (len < 0 || start < 1) {
            result = "Bad arguments";
         } else {
            result = operand_value[1].substring(0, start-1) + operand_value[4] + 
               operand_value[1].substring(start-1+len);
            resulttype = "t";
         }
         break;
      case "REPT":
         count = operand_value[2]-0;
         if (count < 0) {
            result = "Negative count";
         } else {
            result = "";
            for (; count > 0; count--) {
               result += operand_value[1];
            }
            resulttype = "t";
         }
         break;

      case "RIGHT":
         len = operand_type[2] ? operand_value[2]-0 : 1;
         if (len < 0) {
            result = "Negative length";
         } else {
            result = operand_value[1].slice(-len);
            resulttype = "t";
         }
         break;

      case "SUBSTITUTE":
         fulltext = operand_value[1];
         oldtext = operand_value[2];
         newtext = operand_value[3];
         if (operand_value[4] != null) {
            which = operand_value[4]-0;
            if (which <= 0) {
               result = "Non-positive instance number";
               break;
            }
         } else {
            which = 0; // Replace all occurrences
         }
         count = 0;
         oldpos = 0;
         result = "";
         while (true) {
            pos = fulltext.indexOf(oldtext, oldpos);
            if (pos >= 0) {
               count++; //!!!!!! old test just in case: if (count>1000) {alert(pos); break;}
               result += fulltext.substring(oldpos, pos);
               if (which == 0) {
                  result += newtext; // substitute all occurrences
               } else if (which == count) {
                  result += newtext + fulltext.substring(pos+oldtext.length);
                  break; // Found the specific occurrence to replace
               } else {
                  result += oldtext; // leave as was - not the target occurrence
               }
               oldpos = pos + oldtext.length;
            } else { // no more occurrences found
               result += fulltext.substring(oldpos);
               break;
            }
         }
         resulttype = "t";
         break;

      case "TRIM":
         result = operand_value[1];
         result = result.replace(/^ */, "");    // Remove leading spaces
         result = result.replace(/ *$/, "");    // Remove trailing spaces
         result = result.replace(/ +/g, " ");   // Replace multiple spaces with single space
         resulttype = "t";
         break;

      case "UPPER":
         result = operand_value[1].toUpperCase();
         resulttype = "t";
         break;
   }

   scf.PushOperand(operand, resulttype, result);
   return;
};

/**
 * @description Function list registrations for string manipulation functions
 * Each entry follows the format: [function_handler, arg_count, arg_def, func_def, func_class]
 * - Negative arg_count indicates minimum number of arguments
 * - arg_def refers to argument definition keys for help text
 * - func_class: "text" indicates text manipulation function category
 */
SocialCalc.Formula.FunctionList["FIND"] = [SocialCalc.Formula.StringFunctions, -2, "find", "", "text"];
SocialCalc.Formula.FunctionList["LEFT"] = [SocialCalc.Formula.StringFunctions, -2, "tc", "", "text"];
SocialCalc.Formula.FunctionList["LEN"] = [SocialCalc.Formula.StringFunctions, 1, "txt", "", "text"];
SocialCalc.Formula.FunctionList["LOWER"] = [SocialCalc.Formula.StringFunctions, 1, "txt", "", "text"];
SocialCalc.Formula.FunctionList["MID"] = [SocialCalc.Formula.StringFunctions, 3, "mid", "", "text"];
SocialCalc.Formula.FunctionList["PROPER"] = [SocialCalc.Formula.StringFunctions, 1, "v", "", "text"];
SocialCalc.Formula.FunctionList["REPLACE"] = [SocialCalc.Formula.StringFunctions, 4, "replace", "", "text"];
SocialCalc.Formula.FunctionList["REPT"] = [SocialCalc.Formula.StringFunctions, 2, "tc", "", "text"];
SocialCalc.Formula.FunctionList["RIGHT"] = [SocialCalc.Formula.StringFunctions, -1, "tc", "", "text"];
SocialCalc.Formula.FunctionList["SUBSTITUTE"] = [SocialCalc.Formula.StringFunctions, -3, "subs", "", "text"];
SocialCalc.Formula.FunctionList["TRIM"] = [SocialCalc.Formula.StringFunctions, 1, "v", "", "text"];
SocialCalc.Formula.FunctionList["UPPER"] = [SocialCalc.Formula.StringFunctions, 1, "v", "", "text"];

/**
 * @function IsFunctions
 * @memberof SocialCalc.Formula
 * @description Implements value testing functions (IS* family)
 * 
 * These functions test the type or characteristics of values:
 * - ISBLANK(value): Returns TRUE if value is blank/empty
 * - ISERR(value): Returns TRUE if value is an error (except #N/A)
 * - ISERROR(value): Returns TRUE if value is any error (including #N/A)
 * - ISLOGICAL(value): Returns TRUE if value is a logical (boolean) value
 * - ISNA(value): Returns TRUE if value is the #N/A error
 * - ISNONTEXT(value): Returns TRUE if value is not text
 * - ISNUMBER(value): Returns TRUE if value is numeric
 * - ISTEXT(value): Returns TRUE if value is text
 * 
 * @param {string} fname - The specific IS function name
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =ISNUMBER(123) returns 1 (TRUE)
 * // =ISTEXT("Hello") returns 1 (TRUE)
 * // =ISERROR(#DIV/0!) returns 1 (TRUE)
 * // =ISNA(#N/A) returns 1 (TRUE)
 * SocialCalc.Formula.IsFunctions("ISNUMBER", operand, foperand, sheet);
 */
SocialCalc.Formula.IsFunctions = function(fname, operand, foperand, sheet) {
   const scf = SocialCalc.Formula;
   let result = 0;
   const resulttype = "nl"; // Logical result type

   const value = scf.OperandValueAndType(sheet, foperand);
   const t = value.type.charAt(0);

   switch (fname) {
      case "ISBLANK":
         result = value.type == "b" ? 1 : 0;
         break;

      case "ISERR":
         result = t == "e" ? (value.type == "e#N/A" ? 0 : 1) : 0; // Error but not #N/A
         break;

      case "ISERROR":
         result = t == "e" ? 1 : 0; // Any error including #N/A
         break;

      case "ISLOGICAL":
         result = value.type == "nl" ? 1 : 0; // Logical/boolean type
         break;

      case "ISNA":
         result = value.type == "e#N/A" ? 1 : 0; // Specifically #N/A error
         break;

      case "ISNONTEXT":
         result = t == "t" ? 0 : 1; // Not text
         break;

      case "ISNUMBER":
         result = t == "n" ? 1 : 0; // Numeric type
         break;

      case "ISTEXT":
         result = t == "t" ? 1 : 0; // Text type
         break;
   }

   scf.PushOperand(operand, resulttype, result);

   return;
};

/**
 * @description Function list registrations for value testing functions
 * All IS* functions take exactly 1 argument and belong to the "test" category
 */
SocialCalc.Formula.FunctionList["ISBLANK"] = [SocialCalc.Formula.IsFunctions, 1, "v", "", "test"];
SocialCalc.Formula.FunctionList["ISERR"] = [SocialCalc.Formula.IsFunctions, 1, "v", "", "test"];
SocialCalc.Formula.FunctionList["ISERROR"] = [SocialCalc.Formula.IsFunctions, 1, "v", "", "test"];
SocialCalc.Formula.FunctionList["ISLOGICAL"] = [SocialCalc.Formula.IsFunctions, 1, "v", "", "test"];
SocialCalc.Formula.FunctionList["ISNA"] = [SocialCalc.Formula.IsFunctions, 1, "v", "", "test"];
SocialCalc.Formula.FunctionList["ISNONTEXT"] = [SocialCalc.Formula.IsFunctions, 1, "v", "", "test"];
SocialCalc.Formula.FunctionList["ISNUMBER"] = [SocialCalc.Formula.IsFunctions, 1, "v", "", "test"];
SocialCalc.Formula.FunctionList["ISTEXT"] = [SocialCalc.Formula.IsFunctions, 1, "v", "", "test"];

/**
 * @function NTVFunctions
 * @memberof SocialCalc.Formula
 * @description Implements value conversion functions (N, T, VALUE)
 * 
 * These functions convert between different value types:
 * - N(value): Converts value to number, returns 0 if not numeric
 * - T(value): Converts value to text, returns empty string if not text
 * - VALUE(value): Converts text representation of number to actual number
 * 
 * @param {string} fname - The specific conversion function name ("N", "T", or "VALUE")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =N("123") returns 0 (text is not numeric)
 * // =N(123) returns 123 (number stays number)
 * // =T(123) returns "" (number converted to empty text)
 * // =T("Hello") returns "Hello" (text stays text)
 * // =VALUE("123") returns 123 (text parsed as number)
 * SocialCalc.Formula.NTVFunctions("VALUE", operand, foperand, sheet);
 */
SocialCalc.Formula.NTVFunctions = function(fname, operand, foperand, sheet) {
   const scf = SocialCalc.Formula;
   let result = 0;
   let resulttype = "e#VALUE!";

   const value = scf.OperandValueAndType(sheet, foperand);
   const t = value.type.charAt(0);

   switch (fname) {
      case "N":
         result = t == "n" ? value.value-0 : 0; // Convert to number or return 0
         resulttype = "n";
         break;

      case "T":
         result = t == "t" ? value.value+"" : ""; // Convert to text or return empty string
         resulttype = "t";
         break;

      case "VALUE":
         if (t == "n" || t == "b") {
            result = value.value || 0; // Numbers and blanks convert directly
            resulttype = "n";
         } else if (t == "t") {
            // Try to parse text as a number using SocialCalc's value determination
            value = SocialCalc.DetermineValueType(value.value);
            if (value.type.charAt(0) != "n") {
               result = 0;
               resulttype = "e#VALUE!"; // Text cannot be parsed as number
            } else {
               result = value.value-0;
               resulttype = "n";
            }
         }
         break;
   }

   if (t == "e") { // error trumps all other processing
      resulttype = value.type;
   }

   scf.PushOperand(operand, resulttype, result);

   return;
};

/**
 * @description Function list registrations for value conversion functions
 * - N: Converts to number (math category)
 * - T: Converts to text (text category)  
 * - VALUE: Parses text as number (text category)
 */
SocialCalc.Formula.FunctionList["N"] = [SocialCalc.Formula.NTVFunctions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["T"] = [SocialCalc.Formula.NTVFunctions, 1, "v", "", "text"];
SocialCalc.Formula.FunctionList["VALUE"] = [SocialCalc.Formula.NTVFunctions, 1, "v", "", "text"];

/**
 * @function Math1Functions
 * @memberof SocialCalc.Formula
 * @description Implements single-argument mathematical functions
 * 
 * Handles various mathematical operations on a single numeric argument:
 * - ABS(value): Absolute value
 * - ACOS(value): Arc cosine (inverse cosine) in radians
 * - ASIN(value): Arc sine (inverse sine) in radians  
 * - ATAN(value): Arc tangent (inverse tangent) in radians
 * - COS(value): Cosine of angle in radians
 * - DEGREES(value): Convert radians to degrees
 * - EVEN(value): Round up to nearest even integer
 * - EXP(value): e raised to the power of value
 * - FACT(value): Factorial of value
 * - INT(value): Integer part (floor function)
 * - LN(value): Natural logarithm (base e)
 * - LOG10(value): Base 10 logarithm
 * - ODD(value): Round up to nearest odd integer
 * - RADIANS(value): Convert degrees to radians
 * - SIN(value): Sine of angle in radians
 * - SQRT(value): Square root
 * - TAN(value): Tangent of angle in radians
 * 
 * @param {string} fname - The specific mathematical function name
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {null} Always returns null (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction for math functions
 * // =ABS(-5) returns 5
 * // =SIN(PI()/2) returns 1
 * // =FACT(5) returns 120
 * // =SQRT(16) returns 4
 * SocialCalc.Formula.Math1Functions("ABS", operand, foperand, sheet);
 */
SocialCalc.Formula.Math1Functions = function(fname, operand, foperand, sheet) {
   let v1, value, f;
   const result = {};

   const scf = SocialCalc.Formula;

   v1 = scf.OperandAsNumber(sheet, foperand);
   value = v1.value;
   result.type = scf.LookupResultType(v1.type, v1.type, scf.TypeLookupTable.oneargnumeric);

   if (result.type == "n") {
      switch (fname) {
         case "ABS":
            value = Math.abs(value);
            break;

         case "ACOS":
            if (value >= -1 && value <= 1) {
               value = Math.acos(value);
            } else {
               result.type = "e#NUM!"; // Domain error: input must be [-1, 1]
            }
            break;

         case "ASIN":
            if (value >= -1 && value <= 1) {
               value = Math.asin(value);
            } else {
               result.type = "e#NUM!"; // Domain error: input must be [-1, 1]
            }
            break;

         case "ATAN":
            value = Math.atan(value);
            break;

         case "COS":
            value = Math.cos(value);
            break;

         case "DEGREES":
            value = value * 180/Math.PI; // Convert radians to degrees
            break;

         case "EVEN":
            // Round to nearest even integer away from zero
            value = value < 0 ? -value : value;
            if (value != Math.floor(value)) {
               value = Math.floor(value + 1) + (Math.floor(value + 1) % 2);
            } else { // integer
               value = value + (value % 2);
            }
            if (v1.value < 0) value = -value;
            break;

         case "EXP":
            value = Math.exp(value); // e^value
            break;

         case "FACT":
            f = 1;
            value = Math.floor(value);
            for (; value > 0; value--) {
               f *= value;
            }
            value = f;
            break;

         case "INT":
            value = Math.floor(value); // spreadsheet INT is floor(), not int()
            break;

         case "LN":
            if (value <= 0) {
               result.type = "e#NUM!";
               result.error = SocialCalc.Constants.s_sheetfunclnarg;
            } else {
               value = Math.log(value); // Natural logarithm
            }
            break;

         case "LOG10":
            if (value <= 0) {
               result.type = "e#NUM!";
               result.error = SocialCalc.Constants.s_sheetfunclog10arg;
            } else {
               value = Math.log(value)/Math.log(10); // Base 10 logarithm
            }
            break;

         case "ODD":
            // Round to nearest odd integer away from zero
            value = value < 0 ? -value : value;
            if (value != Math.floor(value)) {
               value = Math.floor(value + 1) + (1 - (Math.floor(value + 1) % 2));
            } else { // integer
               value = value + (1 - (value % 2));
            }
            if (v1.value < 0) value = -value;
            break;

         case "RADIANS":
            value = value * Math.PI/180; // Convert degrees to radians
            break;

         case "SIN":
            value = Math.sin(value);
            break;

         case "SQRT":
            if (value >= 0) {
               value = Math.sqrt(value);
            } else {
               result.type = "e#NUM!"; // Cannot take square root of negative number
            }
            break;

         case "TAN":
            if (Math.cos(value) != 0) {
               value = Math.tan(value);
            } else {
               result.type = "e#NUM!"; // Undefined at π/2 + nπ
            }
            break;
      }
   }

   result.value = value;
   operand.push(result);

   return null;
};

/**
 * @description Function list registrations for single-argument mathematical functions
 * All functions take exactly 1 argument and belong to the "math" category
 */
SocialCalc.Formula.FunctionList["ABS"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["ACOS"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["ASIN"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["ATAN"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["COS"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["DEGREES"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["EVEN"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["EXP"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["FACT"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["INT"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["LN"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["LOG10"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["ODD"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["RADIANS"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["SIN"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["SQRT"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];
SocialCalc.Formula.FunctionList["TAN"] = [SocialCalc.Formula.Math1Functions, 1, "v", "", "math"];

/**
 * @function Math2Functions
 * @memberof SocialCalc.Formula
 * @description Implements two-argument mathematical functions
 * 
 * Handles mathematical operations requiring two numeric arguments:
 * - ATAN2(x, y): Arc tangent of y/x in radians, handling quadrant correctly
 * - MOD(a, b): Modulo operation (remainder after division)
 * - POWER(a, b): Raises a to the power of b
 * - TRUNC(value, precision): Truncates value to specified decimal places
 * 
 * @param {string} fname - The specific mathematical function name
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {null} Always returns null (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction for two-argument math functions
 * // =ATAN2(1, 1) returns π/4 (45 degrees in radians)
 * // =MOD(7, 3) returns 1 (7 divided by 3 leaves remainder 1)
 * // =POWER(2, 3) returns 8 (2^3)
 * // =TRUNC(3.14159, 2) returns 3.14
 * SocialCalc.Formula.Math2Functions("POWER", operand, foperand, sheet);
 */
SocialCalc.Formula.Math2Functions = function(fname, operand, foperand, sheet) {
   let xval, yval, value, quotient, decimalscale, i;
   const result = {};

   const scf = SocialCalc.Formula;

   xval = scf.OperandAsNumber(sheet, foperand);
   yval = scf.OperandAsNumber(sheet, foperand);
   value = 0;
   result.type = scf.LookupResultType(xval.type, yval.type, scf.TypeLookupTable.twoargnumeric);

   if (result.type == "n") {
      switch (fname) {
         case "ATAN2":
            if (xval.value == 0 && yval.value == 0) {
               result.type = "e#DIV/0!"; // Undefined for (0,0)
            } else {
               result.value = Math.atan2(yval.value, xval.value); // Note: y, x order for Math.atan2
            }
            break;

         case "POWER":
            result.value = Math.pow(xval.value, yval.value);
            if (isNaN(result.value)) {
               result.value = 0;
               result.type = "e#NUM!"; // Invalid operation (e.g., negative base with fractional exponent)
            }
            break;

         case "MOD": // Modulo operation following Excel semantics
            // Reference: en.wikipedia.org/wiki/Modulo_operation
            if (yval.value == 0) {
               result.type = "e#DIV/0!"; // Division by zero
            } else {
               quotient = xval.value/yval.value;
               quotient = Math.floor(quotient); // Use floor division for Excel compatibility
               result.value = xval.value - (quotient * yval.value);
            }
            break;

         case "TRUNC":
            // Truncate to specified number of decimal places
            decimalscale = 1; // cut down to required number of decimal digits
            if (yval.value >= 0) {
               yval.value = Math.floor(yval.value);
               for (i = 0; i < yval.value; i++) {
                  decimalscale *= 10;
               }
               result.value = Math.floor(Math.abs(xval.value) * decimalscale) / decimalscale;
            } else if (yval.value < 0) {
               yval.value = Math.floor(-yval.value);
               for (i = 0; i < yval.value; i++) {
                  decimalscale *= 10;
               }
               result.value = Math.floor(Math.abs(xval.value) / decimalscale) * decimalscale;
            }
            if (xval.value < 0) {
               result.value = -result.value; // Preserve original sign
            }
            break;
      }
   }
 
   operand.push(result);

   return null;
};

/**
 * @description Function list registrations for two-argument mathematical functions
 * All functions take exactly 2 arguments and belong to the "math" category
 * - ATAN2: Uses "xy" argument definition for coordinate pair
 * - TRUNC: Uses "valpre" argument definition for value and precision
 */
SocialCalc.Formula.FunctionList["ATAN2"] = [SocialCalc.Formula.Math2Functions, 2, "xy", "", "math"];
SocialCalc.Formula.FunctionList["MOD"] = [SocialCalc.Formula.Math2Functions, 2, "", "", "math"];
SocialCalc.Formula.FunctionList["POWER"] = [SocialCalc.Formula.Math2Functions, 2, "", "", "math"];
SocialCalc.Formula.FunctionList["TRUNC"] = [SocialCalc.Formula.Math2Functions, 2, "valpre", "", "math"];
/**
 * @function LogFunction
 * @memberof SocialCalc.Formula
 * @description Implements the LOG function for logarithms with optional base
 * 
 * LOG(value, [base]) returns the logarithm of value to the specified base.
 * If base is omitted, returns the natural logarithm (base e).
 * 
 * @param {string} fname - The function name ("LOG")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =LOG(100, 10) returns 2 (log base 10 of 100)
 * // =LOG(2.718) returns ~1 (natural log of e)
 * // =LOG(8, 2) returns 3 (log base 2 of 8)
 * SocialCalc.Formula.LogFunction("LOG", operand, foperand, sheet);
 */
SocialCalc.Formula.LogFunction = function(fname, operand, foperand, sheet) {
   let value, value2;
   const result = {};

   const scf = SocialCalc.Formula;

   result.value = 0;

   // Get the value to take logarithm of
   value = scf.OperandAsNumber(sheet, foperand);
   result.type = scf.LookupResultType(value.type, value.type, scf.TypeLookupTable.oneargnumeric);
   
   // Check for optional base argument
   if (foperand.length == 1) {
      value2 = scf.OperandAsNumber(sheet, foperand);
      if (value2.type.charAt(0) != "n" || value2.value <= 0) {
         scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfunclogsecondarg);
         return 0;
      }
   } else if (foperand.length != 0) {
      scf.FunctionArgsError(fname, operand);
      return 0;
   } else {
      value2 = {value: Math.E, type: "n"}; // Default to natural logarithm (base e)
   }

   if (result.type == "n") {
      if (value.value <= 0) {
         scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfunclogfirstarg);
         return 0;
      }
      // Calculate logarithm using change of base formula: log_b(x) = ln(x) / ln(b)
      result.value = Math.log(value.value)/Math.log(value2.value);
   }

   operand.push(result);

   return;
};

SocialCalc.Formula.FunctionList["LOG"] = [SocialCalc.Formula.LogFunction, -1, "log", "", "math"];

/**
 * @function RoundFunction
 * @memberof SocialCalc.Formula
 * @description Implements the ROUND function for rounding numbers to specified precision
 * 
 * ROUND(value, [precision]) rounds value to the specified number of decimal places.
 * - If precision is omitted, rounds to nearest integer
 * - If precision is positive, rounds to that many decimal places
 * - If precision is negative, rounds to that many places left of decimal point
 * 
 * @param {string} fname - The function name ("ROUND")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =ROUND(3.14159, 2) returns 3.14
 * // =ROUND(1234.5, -2) returns 1200
 * // =ROUND(2.5) returns 3 (rounds to nearest integer)
 * SocialCalc.Formula.RoundFunction("ROUND", operand, foperand, sheet);
 */
SocialCalc.Formula.RoundFunction = function(fname, operand, foperand, sheet) {
   let value2, decimalscale, scaledvalue, i;

   const scf = SocialCalc.Formula;
   let result = 0;
   let resulttype = "e#VALUE!";

   const value = scf.OperandValueAndType(sheet, foperand);
   resulttype = scf.LookupResultType(value.type, value.type, scf.TypeLookupTable.oneargnumeric);

   // Check for optional precision argument
   if (foperand.length == 1) {
      value2 = scf.OperandValueAndType(sheet, foperand);
      if (value2.type.charAt(0) != "n") {
         scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfuncroundsecondarg);
         return 0;
      }
   } else if (foperand.length != 0) {
      scf.FunctionArgsError(fname, operand);
      return 0;
   } else {
      value2 = {value: 0, type: "n"}; // if no second arg, assume 0 for simple round
   }

   if (resulttype == "n") {
      value2.value = value2.value-0;
      if (value2.value == 0) {
         // Simple rounding to integer
         result = Math.round(value.value);
      } else if (value2.value > 0) {
         // Round to positive decimal places
         decimalscale = 1; // cut down to required number of decimal digits
         value2.value = Math.floor(value2.value);
         for (i = 0; i < value2.value; i++) {
            decimalscale *= 10;
         }
         scaledvalue = Math.round(value.value * decimalscale);
         result = scaledvalue / decimalscale;
      } else if (value2.value < 0) {
         // Round to negative decimal places (round to tens, hundreds, etc.)
         decimalscale = 1; // cut down to required number of decimal digits
         value2.value = Math.floor(-value2.value);
         for (i = 0; i < value2.value; i++) {
            decimalscale *= 10;
         }
         scaledvalue = Math.round(value.value / decimalscale);
         result = scaledvalue * decimalscale;
      }
   }

   scf.PushOperand(operand, resulttype, result);

   return;
};

SocialCalc.Formula.FunctionList["ROUND"] = [SocialCalc.Formula.RoundFunction, -1, "vp", "", "math"];

/**
 * @function AndOrFunctions
 * @memberof SocialCalc.Formula
 * @description Implements logical AND and OR functions for multiple arguments
 * 
 * - AND(v1, v2, ...): Returns TRUE if all arguments are TRUE, FALSE otherwise
 * - OR(v1, v2, ...): Returns TRUE if any argument is TRUE, FALSE otherwise
 * 
 * Arguments are evaluated as logical values:
 * - Non-zero numbers are TRUE
 * - Zero is FALSE
 * - Errors propagate through the result
 * 
 * @param {string} fname - The function name ("AND" or "OR")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =AND(TRUE, 1, 5) returns TRUE (all non-zero)
 * // =AND(TRUE, FALSE) returns FALSE (one is false)
 * // =OR(FALSE, 0, 1) returns TRUE (at least one is true)
 * // =OR(FALSE, 0) returns FALSE (all are false)
 * SocialCalc.Formula.AndOrFunctions("AND", operand, foperand, sheet);
 */
SocialCalc.Formula.AndOrFunctions = function(fname, operand, foperand, sheet) {
   let value1, result;

   const scf = SocialCalc.Formula;
   let resulttype = "";

   // Initialize result based on function type
   if (fname == "AND") {
      result = 1; // Start with TRUE, becomes FALSE if any argument is false
   } else if (fname == "OR") {
      result = 0; // Start with FALSE, becomes TRUE if any argument is true
   }

   // Process all arguments
   while (foperand.length) {
      value1 = scf.OperandValueAndType(sheet, foperand);
      
      if (value1.type.charAt(0) == "n") {
         value1.value = value1.value-0;
         if (fname == "AND") {
            result = value1.value != 0 ? result : 0; // FALSE if any value is 0
         } else if (fname == "OR") {
            result = value1.value != 0 ? 1 : result; // TRUE if any value is non-zero
         }
         resulttype = scf.LookupResultType(value1.type, resulttype || "nl", scf.TypeLookupTable.propagateerror);
      } else if (value1.type.charAt(0) == "e" && resulttype.charAt(0) != "e") {
         resulttype = value1.type; // Propagate error
      }
   }
   
   if (resulttype.length < 1) {
      resulttype = "e#VALUE!";
      result = 0;
   }

   scf.PushOperand(operand, resulttype, result);

   return;
};

SocialCalc.Formula.FunctionList["AND"] = [SocialCalc.Formula.AndOrFunctions, -1, "vn", "", "test"];
SocialCalc.Formula.FunctionList["OR"] = [SocialCalc.Formula.AndOrFunctions, -1, "vn", "", "test"];

/**
 * @function NotFunction
 * @memberof SocialCalc.Formula
 * @description Implements the NOT function for logical negation
 * 
 * NOT(value) returns the logical opposite of value:
 * - TRUE becomes FALSE
 * - FALSE becomes TRUE
 * - Non-zero numbers become FALSE
 * - Zero becomes TRUE
 * - Text values result in #VALUE! error
 * 
 * @param {string} fname - The function name ("NOT")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =NOT(TRUE) returns FALSE
 * // =NOT(0) returns TRUE
 * // =NOT(5) returns FALSE
 * // =NOT("text") returns #VALUE! error
 * SocialCalc.Formula.NotFunction("NOT", operand, foperand, sheet);
 */
SocialCalc.Formula.NotFunction = function(fname, operand, foperand, sheet) {
   let result = 0;
   const scf = SocialCalc.Formula;
   const value = scf.OperandValueAndType(sheet, foperand);
   let resulttype = scf.LookupResultType(value.type, value.type, scf.TypeLookupTable.propagateerror);

   if (value.type.charAt(0) == "n" || value.type == "b") {
      result = value.value-0 != 0 ? 0 : 1; // do the "not" operation
      resulttype = "nl"; // Logical result type
   } else if (value.type.charAt(0) == "t") {
      resulttype = "e#VALUE!"; // Text values are not valid for logical operations
   }

   scf.PushOperand(operand, resulttype, result);

   return;
};

SocialCalc.Formula.FunctionList["NOT"] = [SocialCalc.Formula.NotFunction, 1, "v", "", "test"];

/**
 * @function ChooseFunction
 * @memberof SocialCalc.Formula
 * @description Implements the CHOOSE function for selecting from a list of values
 * 
 * CHOOSE(index, value1, value2, ...) returns the value at the specified index position.
 * - index: 1-based position of the value to return
 * - value1, value2, ...: List of values to choose from
 * 
 * If index is out of range, returns #VALUE! error.
 * 
 * @param {string} fname - The function name ("CHOOSE")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =CHOOSE(2, "A", "B", "C") returns "B"
 * // =CHOOSE(1, 10, 20, 30) returns 10
 * // =CHOOSE(5, "A", "B") returns #VALUE! (index out of range)
 * SocialCalc.Formula.ChooseFunction("CHOOSE", operand, foperand, sheet);
 */
SocialCalc.Formula.ChooseFunction = function(fname, operand, foperand, sheet) {
   let resulttype, count, value1;
   let result = 0;
   const scf = SocialCalc.Formula;

   // Get the index argument
   const cindex = scf.OperandAsNumber(sheet, foperand);

   if (cindex.type.charAt(0) != "n") {
      cindex.value = 0;
   }
   cindex.value = Math.floor(cindex.value); // Use integer index

   // Search through the value arguments for the one at the specified index
   count = 0;
   while (foperand.length) {
      value1 = scf.TopOfStackValueAndType(sheet, foperand);
      count += 1;
      if (cindex.value == count) {
         result = value1.value;
         resulttype = value1.type;
         break;
      }
   }
   
   if (resulttype) { // found something at the specified index
      scf.PushOperand(operand, resulttype, result);
   } else {
      scf.PushOperand(operand, "e#VALUE!", 0); // index out of range
   }

   return;
};

SocialCalc.Formula.FunctionList["CHOOSE"] = [SocialCalc.Formula.ChooseFunction, -2, "choose", "", "lookup"];
/**
 * @function ColumnsRowsFunctions
 * @memberof SocialCalc.Formula
 * @description Implements COLUMNS and ROWS functions for determining range dimensions
 * 
 * - COLUMNS(range): Returns the number of columns in the specified range
 * - ROWS(range): Returns the number of rows in the specified range
 * 
 * For single cell references, both functions return 1.
 * 
 * @param {string} fname - The function name ("COLUMNS" or "ROWS")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =COLUMNS(A1:C5) returns 3 (3 columns: A, B, C)
 * // =ROWS(A1:C5) returns 5 (5 rows: 1, 2, 3, 4, 5)
 * // =COLUMNS(A1) returns 1 (single cell has 1 column)
 * SocialCalc.Formula.ColumnsRowsFunctions("COLUMNS", operand, foperand, sheet);
 */
SocialCalc.Formula.ColumnsRowsFunctions = function(fname, operand, foperand, sheet) {
   let resulttype, rangeinfo;
   let result = 0;
   const scf = SocialCalc.Formula;

   const value1 = scf.TopOfStackValueAndType(sheet, foperand);

   if (value1.type == "coord") {
      // Single cell reference always has 1 row and 1 column
      result = 1;
      resulttype = "n";
   } else if (value1.type == "range") {
      // Decode range to get dimensions
      rangeinfo = scf.DecodeRangeParts(sheet, value1.value);
      if (fname == "COLUMNS") {
         result = rangeinfo.ncols;
      } else if (fname == "ROWS") {
         result = rangeinfo.nrows;
      }
      resulttype = "n";
   } else {
      // Invalid argument type
      result = 0;
      resulttype = "e#VALUE!";
   }

   scf.PushOperand(operand, resulttype, result);

   return;
};

SocialCalc.Formula.FunctionList["COLUMNS"] = [SocialCalc.Formula.ColumnsRowsFunctions, 1, "range", "", "lookup"];
SocialCalc.Formula.FunctionList["ROWS"] = [SocialCalc.Formula.ColumnsRowsFunctions, 1, "range", "", "lookup"];

/**
 * @function ZeroArgFunctions
 * @memberof SocialCalc.Formula
 * @description Implements functions that take no arguments
 * 
 * Handles various constant and volatile functions:
 * - FALSE(): Returns logical FALSE (0)
 * - TRUE(): Returns logical TRUE (1)
 * - NA(): Returns #N/A error value
 * - PI(): Returns the mathematical constant π (pi)
 * - NOW(): Returns current date and time as serial number
 * - TODAY(): Returns current date as serial number (time portion is 0)
 * 
 * NOW() and TODAY() are volatile functions that recalculate automatically.
 * 
 * @param {string} fname - The function name ("FALSE", "TRUE", "NA", "PI", "NOW", or "TODAY")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack (empty for these functions)
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {null} Always returns null (results are pushed to operand stack)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =PI() returns 3.14159265...
 * // =NOW() returns current datetime as serial number
 * // =TODAY() returns current date as serial number
 * // =TRUE() returns 1, =FALSE() returns 0
 * SocialCalc.Formula.ZeroArgFunctions("PI", operand, foperand, sheet);
 */
SocialCalc.Formula.ZeroArgFunctions = function(fname, operand, foperand, sheet) {
   let startval, tzoffset, start_1_1_1970, seconds_in_a_day, nowdays;
   const result = {value: 0};

   switch (fname) {
      case "FALSE":
         result.type = "nl"; // Logical type
         result.value = 0;
         break;

      case "NA":
         result.type = "e#N/A"; // Not Available error
         break;

      case "NOW":
         // Calculate current date and time as Excel serial number
         startval = new Date();
         tzoffset = startval.getTimezoneOffset(); // Minutes from UTC
         startval = startval.getTime() / 1000; // convert to seconds
         start_1_1_1970 = 25569; // Day number of 1/1/1970 starting with 1/1/1900 as 1
         seconds_in_a_day = 24 * 60 * 60;
         nowdays = start_1_1_1970 + startval / seconds_in_a_day - tzoffset/(24*60);
         result.value = nowdays;
         result.type = "ndt"; // Numeric date-time type
         SocialCalc.Formula.FreshnessInfo.volatile.NOW = true; // Mark as volatile for recalculation
         break;

      case "PI":
         result.type = "n"; // Numeric type
         result.value = Math.PI; // Mathematical constant π
         break;

      case "TODAY":
         // Calculate current date (without time) as Excel serial number
         startval = new Date();
         tzoffset = startval.getTimezoneOffset(); // Minutes from UTC
         startval = startval.getTime() / 1000; // convert to seconds
         start_1_1_1970 = 25569; // Day number of 1/1/1970 starting with 1/1/1900 as 1
         seconds_in_a_day = 24 * 60 * 60;
         nowdays = start_1_1_1970 + startval / seconds_in_a_day - tzoffset/(24*60);
         result.value = Math.floor(nowdays); // Remove time portion
         result.type = "nd"; // Numeric date type
         SocialCalc.Formula.FreshnessInfo.volatile.TODAY = true; // Mark as volatile for recalculation
         break;

      case "TRUE":
         result.type = "nl"; // Logical type
         result.value = 1;
         break;
   }

   operand.push(result);

   return null;
};

/**
 * @description Function list registrations for zero-argument functions
 * All functions take exactly 0 arguments and belong to various categories
 * - FALSE, TRUE, NA: test category (logical/error functions)
 * - NOW, TODAY: datetime category (date/time functions)
 * - PI: math category (mathematical constants)
 */
SocialCalc.Formula.FunctionList["FALSE"] = [SocialCalc.Formula.ZeroArgFunctions, 0, "", "", "test"];
SocialCalc.Formula.FunctionList["NA"] = [SocialCalc.Formula.ZeroArgFunctions, 0, "", "", "test"];
SocialCalc.Formula.FunctionList["NOW"] = [SocialCalc.Formula.ZeroArgFunctions, 0, "", "", "datetime"];
SocialCalc.Formula.FunctionList["PI"] = [SocialCalc.Formula.ZeroArgFunctions, 0, "", "", "math"];
SocialCalc.Formula.FunctionList["TODAY"] = [SocialCalc.Formula.ZeroArgFunctions, 0, "", "", "datetime"];
SocialCalc.Formula.FunctionList["TRUE"] = [SocialCalc.Formula.ZeroArgFunctions, 0, "", "", "test"];

/**
 * @namespace SocialCalc.Formula Financial Functions
 * @description Financial calculation functions for depreciation, investments, and loans
 */

/**
 * @function DDBFunction
 * @memberof SocialCalc.Formula
 * @description Implements the DDB function for double-declining balance depreciation
 * 
 * DDB(cost, salvage, lifetime, period, [method]) calculates depreciation using the
 * double-declining balance method or other declining balance methods.
 * 
 * @param {string} fname - The function name ("DDB")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * Arguments:
 * - cost: Initial cost of the asset
 * - salvage: Salvage value at end of useful life
 * - lifetime: Number of periods over which asset is depreciated
 * - period: Period for which to calculate depreciation
 * - method: Optional, defaults to 2 (double-declining balance rate)
 * 
 * @see {@link http://en.wikipedia.org/wiki/Depreciation} Wikipedia article on depreciation
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =DDB(10000, 1000, 5, 1) - Calculate first year depreciation
 * // =DDB(10000, 1000, 5, 1, 1.5) - Use 1.5 declining balance instead of 2
 * SocialCalc.Formula.DDBFunction("DDB", operand, foperand, sheet);
 */
SocialCalc.Formula.DDBFunction = function(fname, operand, foperand, sheet) {
   let method, depreciation, accumulateddepreciation, i;
   const scf = SocialCalc.Formula;

   // Get required arguments
   const cost = scf.OperandAsNumber(sheet, foperand);
   const salvage = scf.OperandAsNumber(sheet, foperand);
   const lifetime = scf.OperandAsNumber(sheet, foperand);
   const period = scf.OperandAsNumber(sheet, foperand);

   // Check for errors in required arguments
   if (scf.CheckForErrorValue(operand, cost)) return;
   if (scf.CheckForErrorValue(operand, salvage)) return;
   if (scf.CheckForErrorValue(operand, lifetime)) return;
   if (scf.CheckForErrorValue(operand, period)) return;

   if (lifetime.value < 1) {
      scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfuncddblife);
      return 0;
   }

   // Get optional method argument (defaults to 2 for double-declining)
   method = {value: 2, type: "n"};
   if (foperand.length > 0 ) {
      method = scf.OperandAsNumber(sheet, foperand);
   }
   if (foperand.length != 0) {
      scf.FunctionArgsError(fname, operand);
      return 0;
   }
   if (scf.CheckForErrorValue(operand, method)) return;

   depreciation = 0; // calculated for each period
   accumulateddepreciation = 0; // accumulated by adding each period's

   // Calculate depreciation for each period up to the requested period
   for (i = 1; i <= period.value-0 && i <= lifetime.value; i++) { // calculate for each period based on net from previous
      depreciation = (cost.value - accumulateddepreciation) * (method.value / lifetime.value);
      if (cost.value - accumulateddepreciation - depreciation < salvage.value) { // don't go lower than salvage value
         depreciation = cost.value - accumulateddepreciation - salvage.value;
      }
      accumulateddepreciation += depreciation;
   }

   scf.PushOperand(operand, 'n$', depreciation); // Return as currency type

   return;
};

SocialCalc.Formula.FunctionList["DDB"] = [SocialCalc.Formula.DDBFunction, -4, "ddb", "", "financial"];

/**
 * @function SLNFunction
 * @memberof SocialCalc.Formula
 * @description Implements the SLN function for straight-line depreciation
 * 
 * SLN(cost, salvage, lifetime) calculates the depreciation for each period
 * using the straight-line method, where the same amount is depreciated each period.
 * 
 * @param {string} fname - The function name ("SLN")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * Arguments:
 * - cost: Initial cost of the asset
 * - salvage: Salvage value at end of useful life
 * - lifetime: Number of periods over which asset is depreciated
 * 
 * Formula: (cost - salvage) / lifetime
 * 
 * @see {@link http://en.wikipedia.org/wiki/Depreciation} Wikipedia article on depreciation
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =SLN(10000, 1000, 5) returns 1800 (depreciation per period)
 * // Asset depreciates $1800 each year for 5 years
 * SocialCalc.Formula.SLNFunction("SLN", operand, foperand, sheet);
 */
SocialCalc.Formula.SLNFunction = function(fname, operand, foperand, sheet) {
   let depreciation;
   const scf = SocialCalc.Formula;

   // Get required arguments
   const cost = scf.OperandAsNumber(sheet, foperand);
   const salvage = scf.OperandAsNumber(sheet, foperand);
   const lifetime = scf.OperandAsNumber(sheet, foperand);

   // Check for errors in arguments
   if (scf.CheckForErrorValue(operand, cost)) return;
   if (scf.CheckForErrorValue(operand, salvage)) return;
   if (scf.CheckForErrorValue(operand, lifetime)) return;

   if (lifetime.value < 1) {
      scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfuncslnlife);
      return 0;
   }

   // Calculate straight-line depreciation: (cost - salvage) / lifetime
   depreciation = (cost.value - salvage.value) / lifetime.value;

   scf.PushOperand(operand, 'n$', depreciation); // Return as currency type

   return;
};

SocialCalc.Formula.FunctionList["SLN"] = [SocialCalc.Formula.SLNFunction, 3, "csl", "", "financial"];
/**
 * @function SYDFunction
 * @memberof SocialCalc.Formula
 * @description Implements the SYD function for Sum of Year's Digits depreciation
 * 
 * SYD(cost, salvage, lifetime, period) calculates depreciation using the Sum of Year's Digits
 * method, which provides accelerated depreciation with higher amounts in earlier years.
 * 
 * The formula uses the sum of all years (1 + 2 + ... + lifetime) as the denominator,
 * and (lifetime - period + 1) as the numerator for the current period.
 * 
 * @param {string} fname - The function name ("SYD")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * Arguments:
 * - cost: Initial cost of the asset
 * - salvage: Salvage value at end of useful life
 * - lifetime: Number of periods over which asset is depreciated
 * - period: Period for which to calculate depreciation (must be > 0)
 * 
 * Formula: (cost - salvage) × (lifetime - period + 1) ÷ sum_of_years
 * Where sum_of_years = lifetime × (lifetime + 1) ÷ 2
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =SYD(10000, 1000, 5, 1) - First year gets 5/15 of total depreciation
 * // =SYD(10000, 1000, 5, 5) - Fifth year gets 1/15 of total depreciation
 * // Sum of years for 5-year life: 1+2+3+4+5 = 15
 * SocialCalc.Formula.SYDFunction("SYD", operand, foperand, sheet);
 */
SocialCalc.Formula.SYDFunction = function(fname, operand, foperand, sheet) {
   let depreciation, sumperiods;
   const scf = SocialCalc.Formula;

   // Get required arguments
   const cost = scf.OperandAsNumber(sheet, foperand);
   const salvage = scf.OperandAsNumber(sheet, foperand);
   const lifetime = scf.OperandAsNumber(sheet, foperand);
   const period = scf.OperandAsNumber(sheet, foperand);

   // Check for errors in arguments
   if (scf.CheckForErrorValue(operand, cost)) return;
   if (scf.CheckForErrorValue(operand, salvage)) return;
   if (scf.CheckForErrorValue(operand, lifetime)) return;
   if (scf.CheckForErrorValue(operand, period)) return;

   // Validate argument values
   if (lifetime.value < 1 || period.value <= 0) {
      scf.PushOperand(operand, "e#NUM!", 0);
      return 0;
   }

   // Calculate sum of years: 1 + 2 + ... + lifetime = lifetime × (lifetime + 1) ÷ 2
   sumperiods = ((lifetime.value + 1) * lifetime.value)/2; // add up 1 through lifetime
   
   // Calculate depreciation for the specified period
   // Higher depreciation in early years: (lifetime - period + 1) gives decreasing weights
   depreciation = (cost.value - salvage.value) * (lifetime.value - period.value + 1) / sumperiods; // calc depreciation

   scf.PushOperand(operand, 'n$', depreciation); // Return as currency type

   return;
};

SocialCalc.Formula.FunctionList["SYD"] = [SocialCalc.Formula.SYDFunction, 4, "cslp", "", "financial"];

/**
 * @function InterestFunctions
 * @memberof SocialCalc.Formula
 * @description Implements financial functions for time value of money calculations
 * 
 * Handles five related financial functions based on the fundamental time value of money equation:
 * - FV(rate, n, payment, [pv, [paytype]]): Future Value
 * - NPER(rate, payment, pv, [fv, [paytype]]): Number of Periods
 * - PMT(rate, n, pv, [fv, [paytype]]): Payment Amount
 * - PV(rate, n, payment, [fv, [paytype]]): Present Value
 * - RATE(n, payment, pv, [fv, [paytype, [guess]]]): Interest Rate
 * 
 * Following the Open Document Format formula specification:
 * - PV = -Fv - (Payment × Nper) [if rate equals 0]
 * - Pv×(1+Rate)^Nper + Payment × (1 + Rate×PaymentType) × ((1+Rate)^nper - 1)/Rate + Fv = 0
 * 
 * For each function, the formulas are solved for the appropriate value using algebra.
 * RATE function uses iterative approximation (Newton-Raphson method).
 * 
 * @param {string} fname - The function name (FV, NPER, PMT, PV, or RATE)
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * Common Arguments:
 * - rate: Interest rate per period
 * - n/nper: Number of periods
 * - payment/pmt: Payment amount per period
 * - pv: Present value (principal)
 * - fv: Future value (optional, defaults to 0)
 * - paytype: Payment timing (0 = end of period, 1 = beginning of period)
 * - guess: Initial guess for RATE calculation (optional, defaults to 0.1)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =FV(0.05/12, 60, -200, -1000) - Future value of loan
 * // =PMT(0.08/12, 360, 100000) - Monthly payment on mortgage
 * // =PV(0.06, 20, -2000, 0) - Present value of annuity
 * // =RATE(60, -200, 10000, 0) - Interest rate calculation
 * SocialCalc.Formula.InterestFunctions("PMT", operand, foperand, sheet);
 */
SocialCalc.Formula.InterestFunctions = function(fname, operand, foperand, sheet) {
   let resulttype, result, dval, eval, fval;
   let pv, fv, rate, n, payment, paytype, guess, part1, part2, part3, part4, part5;
   let olddelta, maxloop, tries, deltaepsilon, oldrate, m;

   const scf = SocialCalc.Formula;

   // Get the first three required arguments
   const aval = scf.OperandAsNumber(sheet, foperand);
   const bval = scf.OperandAsNumber(sheet, foperand);
   const cval = scf.OperandAsNumber(sheet, foperand);

   // Determine result type based on argument types
   resulttype = scf.LookupResultType(aval.type, bval.type, scf.TypeLookupTable.twoargnumeric);
   resulttype = scf.LookupResultType(resulttype, cval.type, scf.TypeLookupTable.twoargnumeric);
   
   // Process optional arguments
   if (foperand.length) { // optional arguments
      dval = scf.OperandAsNumber(sheet, foperand);
      resulttype = scf.LookupResultType(resulttype, dval.type, scf.TypeLookupTable.twoargnumeric);
      if (foperand.length) { // optional arguments
         eval = scf.OperandAsNumber(sheet, foperand);
         resulttype = scf.LookupResultType(resulttype, eval.type, scf.TypeLookupTable.twoargnumeric);
         if (foperand.length) { // optional arguments
            if (fname != "RATE") { // only rate has 6 possible args
               scf.FunctionArgsError(fname, operand);
               return 0;
            }
            fval = scf.OperandAsNumber(sheet, foperand);
            resulttype = scf.LookupResultType(resulttype, fval.type, scf.TypeLookupTable.twoargnumeric);
         }
      }
   }

   if (resulttype == "n") {
      switch (fname) {
         case "FV": // FV(rate, n, payment, [pv, [paytype]])
            rate = aval.value;
            n = bval.value;
            payment = cval.value;
            pv = dval != null ? dval.value : 0; // get value if present, or use default
            paytype = eval != null ? (eval.value ? 1 : 0) : 0;
            
            if (rate == 0) { // simple calculation if no interest
               fv = -pv - (payment * n);
            } else {
               fv = -(pv*Math.pow(1+rate,n) + payment * (1 + rate*paytype) * ( Math.pow(1+rate,n) -1)/rate);
            }
            result = fv;
            resulttype = 'n$'; // Currency type
            break;

         case "NPER": // NPER(rate, payment, pv, [fv, [paytype]])
            rate = aval.value;
            payment = bval.value;
            pv = cval.value;
            fv = dval != null ? dval.value : 0;
            paytype = eval != null ? (eval.value ? 1 : 0) : 0;
            
            if (rate == 0) { // simple calculation if no interest
               if (payment == 0) {
                  scf.PushOperand(operand, "e#NUM!", 0);
                  return;
               }
               n = (pv + fv)/(-payment);
            } else {
               part1 = payment * (1 + rate * paytype) / rate;
               part2 = pv + part1;
               if (part2 == 0 || rate <= -1) {
                  scf.PushOperand(operand, "e#NUM!", 0);
                  return;
               }
               part3 = (part1 - fv) / part2;
               if (part3 <= 0) {
                  scf.PushOperand(operand, "e#NUM!", 0);
                  return;
               }
               part4 = Math.log(part3);
               part5 = Math.log(1 + rate); // rate > -1
               n = part4/part5;
            }
            result = n;
            resulttype = 'n'; // Numeric type
            break;

         case "PMT": // PMT(rate, n, pv, [fv, [paytype]])
            rate = aval.value;
            n = bval.value;
            pv = cval.value;
            fv = dval != null ? dval.value : 0;
            paytype = eval != null ? (eval.value ? 1 : 0) : 0;
            
            if (n == 0) {
               scf.PushOperand(operand, "e#NUM!", 0);
               return;
            } else if (rate == 0) { // simple calculation if no interest
               payment = (fv - pv)/n;
            } else {
               payment = (0 - fv - pv*Math.pow(1+rate,n))/((1 + rate*paytype) * ( Math.pow(1+rate,n) -1)/rate);
            }
            result = payment;
            resulttype = 'n$'; // Currency type
            break;

         case "PV": // PV(rate, n, payment, [fv, [paytype]])
            rate = aval.value;
            n = bval.value;
            payment = cval.value;
            fv = dval != null ? dval.value : 0;
            paytype = eval != null ? (eval.value ? 1 : 0) : 0;
            
            if (rate == -1) {
               scf.PushOperand(operand, "e#DIV/0!", 0);
               return;
            } else if (rate == 0) { // simple calculation if no interest
               pv = -fv - (payment * n);
            } else {
               pv = (-fv - payment * (1 + rate*paytype) * ( Math.pow(1+rate,n) -1)/rate)/(Math.pow(1+rate,n));
            }
            result = pv;
            resulttype = 'n$'; // Currency type
            break;

         case "RATE": // RATE(n, payment, pv, [fv, [paytype, [guess]]])
            n = aval.value;
            payment = bval.value;
            pv = cval.value;
            fv = dval != null ? dval.value : 0;
            paytype = eval != null ? (eval.value ? 1 : 0) : 0;
            guess = fval != null ? fval.value : 0.1;

            // Rate is calculated by repeated approximations using Newton-Raphson method
            // The deltas are used to calculate new guesses
            maxloop = 100;
            tries = 0;
            let delta = 1;
            const epsilon = 0.0000001; // convergence threshold
            rate = guess || 0.00000001; // zero is not allowed
            
            while ((delta >= 0 ? delta : -delta) > epsilon && (rate != oldrate)) {
               delta = fv + pv*Math.pow(1+rate,n) + payment * (1 + rate*paytype) * ( Math.pow(1+rate,n) -1)/rate;
               if (olddelta != null) {
                  m = (delta - olddelta)/(rate - oldrate) || .001; // get slope (not zero)
                  oldrate = rate;
                  rate = rate - delta / m; // look for zero crossing
                  olddelta = delta;
               } else { // first time - no old values
                  oldrate = rate;
                  rate = 1.1 * rate;
                  olddelta = delta;
               }
               tries++;
               if (tries >= maxloop) { // didn't converge yet
                  scf.PushOperand(operand, "e#NUM!", 0);
                  return;
               }
            }
            result = rate;
            resulttype = 'n%'; // Percentage type
            break;
      }
   }
 
   scf.PushOperand(operand, resulttype, result);

   return;
};

/**
 * @description Function list registrations for financial interest/time-value-of-money functions
 * All functions take at least 3 arguments (indicated by -3) and belong to the "financial" category
 * - Different argument definition keys correspond to help text for each function's parameters
 */
SocialCalc.Formula.FunctionList["FV"] = [SocialCalc.Formula.InterestFunctions, -3, "fv", "", "financial"];
SocialCalc.Formula.FunctionList["NPER"] = [SocialCalc.Formula.InterestFunctions, -3, "nper", "", "financial"];
SocialCalc.Formula.FunctionList["PMT"] = [SocialCalc.Formula.InterestFunctions, -3, "pmt", "", "financial"];
SocialCalc.Formula.FunctionList["PV"] = [SocialCalc.Formula.InterestFunctions, -3, "pv", "", "financial"];
SocialCalc.Formula.FunctionList["RATE"] = [SocialCalc.Formula.InterestFunctions, -3, "rate", "", "financial"];
/**
 * @function NPVFunction
 * @memberof SocialCalc.Formula
 * @description Implements the NPV function for Net Present Value calculation
 * 
 * NPV(rate, value1, value2, ...) calculates the net present value of a series of cash flows
 * occurring at regular intervals, discounted at a specified rate.
 * 
 * The function applies the formula: NPV = Σ(CF_t / (1 + rate)^t) where:
 * - CF_t is the cash flow at time t
 * - rate is the discount rate per period
 * - t is the time period (1, 2, 3, ...)
 * 
 * @param {string} fname - The function name ("NPV")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * Arguments:
 * - rate: Discount rate per period
 * - value1, value2, ...: Series of cash flows (can include ranges)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =NPV(0.1, -1000, 500, 500, 500) - Initial investment followed by returns
 * // =NPV(0.08, A1:A5) - NPV of cash flows in range A1:A5
 * SocialCalc.Formula.NPVFunction("NPV", operand, foperand, sheet);
 */
SocialCalc.Formula.NPVFunction = function(fname, operand, foperand, sheet) {
   let resulttypenpv, rate, sum, factor, value1;

   const scf = SocialCalc.Formula;

   // Get the discount rate (first argument)
   rate = scf.OperandAsNumber(sheet, foperand);
   if (scf.CheckForErrorValue(operand, rate)) return;

   sum = 0;
   resulttypenpv = "n";
   factor = 1; // Accumulates (1 + rate)^t for each period

   // Process each cash flow value
   while (foperand.length) {
      value1 = scf.OperandValueAndType(sheet, foperand);
      if (value1.type.charAt(0) == "n") {
         factor *= (1 + rate.value); // Compound the discount factor
         if (factor == 0) {
            scf.PushOperand(operand, "e#DIV/0!", 0);
            return;
         }
         sum += value1.value / factor; // Discount the cash flow to present value
         resulttypenpv = scf.LookupResultType(value1.type, resulttypenpv || value1.type, scf.TypeLookupTable.plus);
      } else if (value1.type.charAt(0) == "e" && resulttypenpv.charAt(0) != "e") {
         resulttypenpv = value1.type; // Propagate error
         break;
      }
   }

   if (resulttypenpv.charAt(0) == "n") {
      resulttypenpv = 'n$'; // Return as currency type
   }

   scf.PushOperand(operand, resulttypenpv, sum);

   return;
};

SocialCalc.Formula.FunctionList["NPV"] = [SocialCalc.Formula.NPVFunction, -2, "npv", "", "financial"];

/**
 * @function IRRFunction
 * @memberof SocialCalc.Formula
 * @description Implements the IRR function for Internal Rate of Return calculation
 * 
 * IRR(values, [guess]) calculates the internal rate of return for a series of cash flows
 * represented by the values array. IRR is the rate at which the net present value
 * of the cash flows equals zero.
 * 
 * Uses iterative approximation (Newton-Raphson method) to find the rate where:
 * NPV = Σ(CF_t / (1 + IRR)^t) = 0
 * 
 * @param {string} fname - The function name ("IRR")
 * @param {Array<Object>} operand - The main operand stack
 * @param {Array<Object>} foperand - The function operand stack containing arguments
 * @param {Object} sheet - The spreadsheet object containing cell data
 * @returns {undefined} Returns undefined (results/errors are pushed to operand stack)
 * 
 * Arguments:
 * - values: Range or array of cash flows (must contain at least one positive and one negative value)
 * - guess: Optional initial guess for the rate (defaults to 0.1 or 10%)
 * 
 * @example
 * // Called internally by CalculateFunction
 * // =IRR(A1:A5) - IRR for cash flows in range A1:A5
 * // =IRR(A1:A5, 0.15) - IRR with initial guess of 15%
 * // Typical: =IRR({-1000, 300, 400, 500}) - Investment followed by returns
 * SocialCalc.Formula.IRRFunction("IRR", operand, foperand, sheet);
 */
SocialCalc.Formula.IRRFunction = function(fname, operand, foperand, sheet) {
   let value1, guess, oldsum, maxloop, tries, epsilon, rate, oldrate, m, sum, factor, i;
   const rangeoperand = [];
   const cashflows = [];

   const scf = SocialCalc.Formula;

   // Get the range of cash flows (first argument)
   rangeoperand.push(foperand.pop()); // first operand is a range

   // Extract all cash flow values from the range
   while (rangeoperand.length) { // get values from range so we can do iterative approximations
      value1 = scf.OperandValueAndType(sheet, rangeoperand);
      if (value1.type.charAt(0) == "n") {
         cashflows.push(value1.value);
      } else if (value1.type.charAt(0) == "e") {
         scf.PushOperand(operand, "e#VALUE!", 0);
         return;
      }
   }

   if (!cashflows.length) {
      scf.PushOperand(operand, "e#NUM!", 0);
      return;
   }

   // Get optional guess argument
   guess = {value: 0};

   if (foperand.length) { // guess is provided
      guess = scf.OperandAsNumber(sheet, foperand);
      if (guess.type.charAt(0) != "n" && guess.type.charAt(0) != "b") {
         scf.PushOperand(operand, "e#VALUE!", 0);
         return;
      }
      if (foperand.length) { // should be no more args
         scf.FunctionArgsError(fname, operand);
         return;
      }
   }

   guess.value = guess.value || 0.1; // Default guess is 10%

   // Rate is calculated by repeated approximations using Newton-Raphson method
   // We're looking for the rate where NPV = 0
   maxloop = 20;
   tries = 0;
   epsilon = 0.0000001; // convergence threshold
   rate = guess.value;
   sum = 1;

   while ((sum >= 0 ? sum : -sum) > epsilon && (rate != oldrate)) {
      sum = 0;
      factor = 1;
      
      // Calculate NPV at current rate
      for (i = 0; i < cashflows.length; i++) {
         factor *= (1 + rate);
         if (factor == 0) {
            scf.PushOperand(operand, "e#DIV/0!", 0);
            return;
         }
         sum += cashflows[i] / factor;
      }

      // Use Newton-Raphson method to find better rate guess
      if (oldsum != null) {
         m = (sum - oldsum)/(rate - oldrate); // get slope (derivative approximation)
         oldrate = rate;
         rate = rate - sum / m; // look for zero crossing
         oldsum = sum;
      } else { // first time - no old values
         oldrate = rate;
         rate = 1.1 * rate; // perturb rate to get started
         oldsum = sum;
      }
      
      tries++;
      if (tries >= maxloop) { // didn't converge yet
         scf.PushOperand(operand, "e#NUM!", 0);
         return;
      }
   }

   scf.PushOperand(operand, 'n%', rate); // Return as percentage type

   return;
};

SocialCalc.Formula.FunctionList["IRR"] = [SocialCalc.Formula.IRRFunction, -1, "irr", "", "financial"];

/**
 * @namespace SocialCalc.Formula.SheetCache
 * @description Sheet caching system for cross-sheet references
 * 
 * Manages loading and caching of external sheets referenced in formulas.
 * Each sheet is loaded only once and stored with recalculation state tracking.
 */
SocialCalc.Formula.SheetCache = {
   /**
    * @type {Object<string, Object>}
    * @description Sheet data cache
    * 
    * Attributes are each sheet in the cache with values of an object containing:
    * - sheet: SocialCalc.Sheet object (or null if not found)
    * - recalcstate: Current recalculation state (asloaded/recalcing/recalcdone)
    * - name: name of sheet (for reference when you only have the object)
    */
   sheets: {},

   /**
    * @type {string|null}
    * @description Waiting for loading indicator
    * 
    * If sheet is not in cache, this is set to the sheetname being loaded
    * so it can be tested in the recalc loop to start load and then wait until restarted.
    * Reset to null before restarting.
    */
   waitingForLoading: null,

   /**
    * @type {Object<string, number>}
    * @description Constants for sheet recalculation states
    */
   constants: {
      asloaded: 0,    // Sheet loaded but not recalculated
      recalcing: 1,   // Sheet currently being recalculated
      recalcdone: 2   // Sheet recalculation completed
   },

   /**
    * @type {Function|null}
    * @description Deprecated sheet loading function
    * @deprecated Use SocialCalc.RecalcInfo.LoadSheet instead
    */
   loadsheet: null
};

/**
 * @function FindInSheetCache
 * @memberof SocialCalc.Formula
 * @description Returns a SocialCalc.Sheet object corresponding to the given sheet name
 * 
 * Each sheet is loaded only once and then stored in a cache.
 * Loading is handled elsewhere, e.g., in the recalc loop.
 * 
 * @param {string} sheetname - Name of the sheet to find
 * @returns {SocialCalc.Sheet|null} Sheet object if found/cached, null if not available or loading
 * 
 * @example
 * const sheet = SocialCalc.Formula.FindInSheetCache("Sheet2");
 * if (sheet) {
 *    // Use the cached sheet
 * } else {
 *    // Sheet not available or loading
 * }
 */
SocialCalc.Formula.FindInSheetCache = function(sheetname) {
   let str;
   const sfsc = SocialCalc.Formula.SheetCache;

   const nsheetname = SocialCalc.Formula.NormalizeSheetName(sheetname); // normalize different versions

   if (sfsc.sheets[nsheetname]) { // a sheet by that name is in the cache already
      return sfsc.sheets[nsheetname].sheet; // return it
   }

   if (sfsc.waitingForLoading) { // waiting already - only queue up one
      return null; // return not found
   }

   sfsc.waitingForLoading = nsheetname; // let recalc loop know that we have a sheet to load

   return null; // return not found
};

/**
 * @function AddSheetToCache
 * @memberof SocialCalc.Formula
 * @description Adds a new sheet to the sheet cache
 * 
 * Creates a new SocialCalc.Sheet object from the provided saved sheet string
 * and adds it to the cache with appropriate state tracking.
 * 
 * @param {string} sheetname - Name of the sheet to add
 * @param {string} str - Saved sheet data string to parse
 * @returns {SocialCalc.Sheet|null} The newly created and cached sheet object
 * 
 * @example
 * const newSheet = SocialCalc.Formula.AddSheetToCache("Sheet2", savedSheetData);
 * // Sheet is now available in cache for formula references
 */
SocialCalc.Formula.AddSheetToCache = function(sheetname, str) {
   let newsheet = null;
   const sfsc = SocialCalc.Formula.SheetCache;
   const sfscc = sfsc.constants;
   const newsheetname = SocialCalc.Formula.NormalizeSheetName(sheetname);

   if (str) {
      newsheet = new SocialCalc.Sheet();
      newsheet.ParseSheetSave(str);
   }

   sfsc.sheets[newsheetname] = {
      sheet: newsheet, 
      recalcstate: sfscc.asloaded, 
      name: newsheetname
   };

   SocialCalc.Formula.FreshnessInfo.sheets[newsheetname] = true;

   return newsheet;
};

/**
 * @function NormalizeSheetName
 * @memberof SocialCalc.Formula
 * @description Normalizes sheet names for consistent cache lookup
 * 
 * Uses a callback function if available, otherwise converts to lowercase.
 * This ensures that sheet names are handled consistently regardless of case variations.
 * 
 * @param {string} sheetname - The sheet name to normalize
 * @returns {string} Normalized sheet name
 * 
 * @example
 * const normalized = SocialCalc.Formula.NormalizeSheetName("MySheet");
 * // Returns "mysheet" (or result of custom callback)
 */
SocialCalc.Formula.NormalizeSheetName = function(sheetname) {
   if (SocialCalc.Callbacks.NormalizeSheetName) {
      return SocialCalc.Callbacks.NormalizeSheetName(sheetname);
   } else {
      return sheetname.toLowerCase();
   }
};

/**
 * @namespace SocialCalc.Formula.RemoteFunctionInfo
 * @description Remote function execution information for server-side calculations
 * 
 * Tracks the state of remote function calls that require server communication.
 */
SocialCalc.Formula.RemoteFunctionInfo = {
   /**
    * @type {string|null}
    * @description Server communication status
    * 
    * If waiting for an XHR response from the server, this is set to some non-blank status text
    * so it can be tested in the recalc loop to start load and then wait until restarted.
    * Reset to null before restarting.
    */
   waitingForServer: null
};

/**
 * @namespace SocialCalc.Formula.FreshnessInfo
 * @description Freshness tracking for spreadsheet dependencies
 * 
 * This information is generated during recalc and may be used to help determine
 * when the recalc data in a spreadsheet may be out of date.
 * 
 * For example, it may be used to display messages like:
 * "Dependent on sheet 'FOO' which was updated more recently than this printout"
 */
SocialCalc.Formula.FreshnessInfo = {
   /**
    * @type {Object<string, boolean>}
    * @description External sheet dependencies
    * 
    * For each external sheet referenced successfully, an attribute of that name with value true.
    */
   sheets: {},

   /**
    * @type {Object<string, boolean>}
    * @description Volatile function usage
    * 
    * For each volatile function that is called, an attribute of that name with value true.
    * Volatile functions include NOW(), TODAY(), and RAND().
    */
   volatile: {},

   /**
    * @type {boolean}
    * @description Recalculation completion status
    * 
    * Set to false when recalc is started and true when recalc completes successfully.
    */
   recalc_completed: false
};

/**
 * @function FreshnessInfoReset
 * @memberof SocialCalc.Formula
 * @description Resets all freshness tracking information
 * 
 * Called at the start of recalculation to clear dependency and volatility tracking
 * from the previous calculation cycle.
 * 
 * @example
 * SocialCalc.Formula.FreshnessInfoReset();
 * // All freshness tracking is now cleared for new recalc cycle
 */
SocialCalc.Formula.FreshnessInfoReset = function() {
   const scffi = SocialCalc.Formula.FreshnessInfo;

   scffi.sheets = {};
   scffi.volatile = {};
   scffi.recalc_completed = false;
};
/**
 * @namespace SocialCalc.Formula Miscellaneous Routines
 * @description Utility functions for coordinate manipulation, range processing, and criteria testing
 */

/**
 * @function PlainCoord
 * @memberof SocialCalc.Formula
 * @description Removes all dollar signs from a coordinate reference
 * 
 * Strips absolute reference markers ($) from cell coordinates, converting
 * absolute and mixed references to relative format.
 * 
 * @param {string} coord - Cell coordinate that may contain $ symbols
 * @returns {string} Coordinate without any $ symbols
 * 
 * @example
 * SocialCalc.Formula.PlainCoord("$A$1") // Returns "A1"
 * SocialCalc.Formula.PlainCoord("$A1")  // Returns "A1"
 * SocialCalc.Formula.PlainCoord("A$1")  // Returns "A1"
 * SocialCalc.Formula.PlainCoord("A1")   // Returns "A1"
 */
SocialCalc.Formula.PlainCoord = function(coord) {
   if (coord.indexOf("$") == -1) return coord;

   return coord.replace(/\$/g, ""); // remove any $'s
};

/**
 * @function OrderRangeParts
 * @memberof SocialCalc.Formula
 * @description Orders two coordinates to ensure proper range boundaries
 * 
 * Takes two cell coordinates and returns the range boundaries with
 * the upper-left corner as c1/r1 and lower-right corner as c2/r2.
 * This ensures ranges are always specified in canonical form regardless
 * of the order the coordinates were provided.
 * 
 * @param {string} coord1 - First cell coordinate (e.g., "A1")
 * @param {string} coord2 - Second cell coordinate (e.g., "C3")
 * @returns {Object} Range boundaries object containing:
 *   - {number} c1 - Left column number
 *   - {number} r1 - Top row number
 *   - {number} c2 - Right column number
 *   - {number} r2 - Bottom row number
 * 
 * @example
 * SocialCalc.Formula.OrderRangeParts("C3", "A1")
 * // Returns: {c1: 1, r1: 1, c2: 3, r2: 3} (A1 to C3)
 * 
 * SocialCalc.Formula.OrderRangeParts("A1", "C3")  
 * // Returns: {c1: 1, r1: 1, c2: 3, r2: 3} (same result)
 */
SocialCalc.Formula.OrderRangeParts = function(coord1, coord2) {
   let cr1, cr2;
   const result = {};

   cr1 = SocialCalc.coordToCr(coord1);
   cr2 = SocialCalc.coordToCr(coord2);
   
   // Order columns (c1 <= c2)
   if (cr1.col > cr2.col) { 
      result.c1 = cr2.col; 
      result.c2 = cr1.col; 
   } else { 
      result.c1 = cr1.col; 
      result.c2 = cr2.col; 
   }
   
   // Order rows (r1 <= r2)
   if (cr1.row > cr2.row) { 
      result.r1 = cr2.row; 
      result.r2 = cr1.row; 
   } else { 
      result.r1 = cr1.row; 
      result.r2 = cr2.row; 
   }

   return result;
};

/**
 * @function TestCriteria
 * @memberof SocialCalc.Formula
 * @description Tests whether a value meets specified criteria
 * 
 * Determines whether a value/type combination meets the given criteria.
 * Used extensively by database functions (DSUM, DCOUNT, etc.) and conditional
 * functions (COUNTIF, SUMIF, etc.) for filtering data.
 * 
 * Criteria formats supported:
 * - Numeric value: exact match
 * - Text: partial match from beginning (unless with comparator)
 * - "=value": exact match
 * - "<value", "<=value": less than (or equal)
 * - ">value", ">=value": greater than (or equal)  
 * - "<>value": not equal to
 * - Empty criteria: matches nothing
 * 
 * @param {*} value - The value to test
 * @param {string} type - SocialCalc type of the value ("n", "t", "b", "e", etc.)
 * @param {*} criteria - The criteria to test against (can be null/undefined)
 * @returns {boolean} True if value meets criteria, false otherwise
 * 
 * @example
 * // Numeric comparisons
 * SocialCalc.Formula.TestCriteria(10, "n", ">5")     // Returns true
 * SocialCalc.Formula.TestCriteria(3, "n", ">=5")     // Returns false
 * SocialCalc.Formula.TestCriteria(5, "n", "5")       // Returns true
 * 
 * // Text comparisons  
 * SocialCalc.Formula.TestCriteria("Apple", "t", "App")      // Returns true (starts with)
 * SocialCalc.Formula.TestCriteria("Apple", "t", "=Apple")   // Returns true (exact)
 * SocialCalc.Formula.TestCriteria("Apple", "t", "<>Banana") // Returns true (not equal)
 * 
 * // Edge cases
 * SocialCalc.Formula.TestCriteria("", "b", "")        // Returns false (blank vs empty)
 * SocialCalc.Formula.TestCriteria("text", "t", null)  // Returns false (null criteria)
 */
SocialCalc.Formula.TestCriteria = function(value, type, criteria) {
   let comparitor, basestring, basevalue, cond, testvalue;

   if (criteria == null) { // undefined (e.g., error value) is always false
      return false;
   }

   criteria = criteria + "";
   
   // Parse comparator from criteria string
   comparitor = criteria.charAt(0); // look for comparitor
   if (comparitor == "=" || comparitor == "<" || comparitor == ">") {
      basestring = criteria.substring(1);
   } else {
      comparitor = criteria.substring(0,2);
      if (comparitor == "<=" || comparitor == "<>" || comparitor == ">=") {
         basestring = criteria.substring(2);
      } else {
         comparitor = "none"; // no explicit comparator
         basestring = criteria;
      }
   }

   // Determine the type and value of the criteria
   basevalue = SocialCalc.DetermineValueType(basestring); // get type of value being compared
   
   if (!basevalue.type) { // no criteria base value given
      if (comparitor == "none") { // blank criteria matches nothing
         return false;
      }
      if (type.charAt(0) == "b") { // comparing to empty cell
         if (comparitor == "=") { // empty equals empty
            return true;
         }
      } else {
         if (comparitor == "<>") { // "something" does not equal empty
            return true;
         }
      }
      return false; // otherwise false
   }

   cond = false;

   // Handle case where criteria is number but value is text (try to convert)
   if (basevalue.type.charAt(0) == "n" && type.charAt(0) == "t") { // criteria is number, but value is text
      testvalue = SocialCalc.DetermineValueType(value);
      if (testvalue.type.charAt(0) == "n") { // could be number - make it one
         value = testvalue.value;
         type = testvalue.type;
      }
   }

   // Numeric comparison
   if (type.charAt(0) == "n" && basevalue.type.charAt(0) == "n") { // compare two numbers
      value = value - 0; // ensure numeric
      basevalue.value = basevalue.value - 0;
      
      switch (comparitor) {
         case "<":
            cond = value < basevalue.value;
            break;
         case "<=":
            cond = value <= basevalue.value;
            break;
         case "=":
         case "none":
            cond = value == basevalue.value;
            break;
         case ">=":
            cond = value >= basevalue.value;
            break;
         case ">":
            cond = value > basevalue.value;
            break;
         case "<>":
            cond = value != basevalue.value;
            break;
      }
   } 
   // Error handling
   else if (type.charAt(0) == "e") { // error on left
      cond = false;
   } else if (basevalue.type.charAt(0) == "e") { // error on right
      cond = false;
   } 
   // Text comparison
   else { // text, maybe mixed with number or blank
      if (type.charAt(0) == "n") {
         value = SocialCalc.format_number_for_display(value, "n", "");
      }
      if (basevalue.type.charAt(0) == "n") {
         return false; // if number and didn't match already, isn't a match
      }

      // Convert to lowercase for case-insensitive comparison
      value = value ? value.toLowerCase() : "";
      basevalue.value = basevalue.value ? basevalue.value.toLowerCase() : "";

      switch (comparitor) {
         case "<":
            cond = value < basevalue.value;
            break;
         case "<=":
            cond = value <= basevalue.value;
            break;
         case "=":
            cond = value == basevalue.value;
            break;
         case "none":
            // Default text matching: starts with the criteria text
            cond = value.substring(0, basevalue.value.length) == basevalue.value;
            break;
         case ">=":
            cond = value >= basevalue.value;
            break;
         case ">":
            cond = value > basevalue.value;
            break;
         case "<>":
            cond = value != basevalue.value;
            break;
      }
   }

   return cond;
};
