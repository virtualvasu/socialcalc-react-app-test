/* ---------------------------------------------------------------
 * SocialCalc - Popup Manager  (modernized ES6+)
 * ---------------------------------------------------------------
 * This file is part of the SocialCalc package.
 * Originally written 2009 by Socialtext, Inc.
 * Modernized 2025 – functionality unchanged.
 * ------------------------------------------------------------- */

(() => {
  'use strict';

  /* ---------- GLOBAL NAMESPACE SET-UP ---------- */

  // Ensure a single global SocialCalc object
  const root = typeof window !== 'undefined' ? window : globalThis;
  const SocialCalc = root.SocialCalc ?? (root.SocialCalc = {});

  // Ensure the Popup namespace exists
  const Popup = SocialCalc.Popup ?? (SocialCalc.Popup = {});

  /* ---------- CORE DATA STRUCTURES ---------- */

  Popup.Types     = Popup.Types     ?? {};
  Popup.Controls  = Popup.Controls  ?? {};
  Popup.Current   = Popup.Current   ?? {};
  Popup.imagePrefix = '/src/images/sc-';            // image path prefix for React app

  // Override externally for i18n if needed
  Popup.LocalizeString = Popup.LocalizeString ?? (str => str);

  /* =============================================================
   * GENERAL ROUTINES
   * ===========================================================*/

  /**
   * Create a control of a given type under element `id`.
   * @param {string} type      – registered control type
   * @param {string} id        – parent element id
   * @param {object} [attribs] – type-specific attributes
   */
  Popup.Create = (type, id, attribs = {}) => {
    const handler = Popup.Types[type];
    handler?.Create?.(type, id, attribs);
  };

  /**
   * Set the value for a control and trigger its change callback.
   * @param {string} id     – control id
   * @param {*}      value  – new value
   */
  Popup.SetValue = (id, value) => {
    const control = Popup.Controls[id];
    if (!control) { alert(`Unknown control ${id}`); return; }

    const { type, data } = control;
    Popup.Types[type]?.SetValue?.(type, id, value);

    // Fire any caller-supplied change callback
    data?.attribs?.changedcallback?.(data.attribs, id, value);
  };

  /**
   * Enable or disable a control.
   * @param {string}  id        – control id
   * @param {boolean} disabled  – true ⇢ disable, false ⇢ enable
   */
  Popup.SetDisabled = (id, disabled) => {
    const control = Popup.Controls[id];
    if (!control) { alert(`Unknown control ${id}`); return; }

    const { type } = control;
    const current = Popup.Current;

    // Close if this control’s popup is open
    if (current.id && id === current.id) {
      Popup.Types[type]?.Hide?.(type, current.id);
      current.id = null;
    }
    Popup.Types[type]?.SetDisabled?.(type, id, disabled);
  };

  /**
   * Get the current value of a control.
   * @param  {string} id – control id
   * @return {*}         – current value or null
   */
  Popup.GetValue = id => {
    const control = Popup.Controls[id];
    if (!control) { alert(`Unknown control ${id}`); return null; }

    const { type } = control;
    return Popup.Types[type]?.GetValue?.(type, id) ?? null;
  };

  /**
   * Pass type-specific initialization data to a control.
   * @param {string} id    – control id
   * @param {object} data  – initialization data
   */
  Popup.Initialize = (id, data) => {
    const control = Popup.Controls[id];
    if (!control) { alert(`Unknown control ${id}`); return; }

    const { type } = control;
    Popup.Types[type]?.Initialize?.(type, id, data);
  };

  /**
   * Reset all pop-up state for a given type (e.g., when changing pages).
   * @param {string} type – registered control type
   */
  Popup.Reset = type => {
    Popup.Types[type]?.Reset?.(type);
  };

})();  

/* =============================================================

/**
 * SocialCalc.Popup.CClick(id)
 *
 * Should be called when the user clicks on a control to do the popup
 *
 * @param {string} id
 */
SocialCalc.Popup.CClick = function(id) {
   const sp  = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;

   if (!spc[id]) { alert("Unknown control " + id); return; }

   if (spc[id].data && spc[id].data.disabled) return;

   const type = spc[id].type;
   const pt   = spt[type];

   if (sp.Current.id) {
      spt[spc[sp.Current.id].type].Hide(type, sp.Current.id);
      if (id == sp.Current.id) { // same one - done
         sp.Current.id = null;
         return;
      }
   }

   if (pt && pt.Show) {
      pt.Show(type, id);
   }

   sp.Current.id = id;
};

/**
 * SocialCalc.Popup.Close()
 *
 * Used to close any open popup.
 */
SocialCalc.Popup.Close = function() {
   const sp = SocialCalc.Popup;

   if (!sp.Current.id) return;

   sp.CClick(sp.Current.id);
};

/**
 * SocialCalc.Popup.Cancel()
 *
 * Closes Popup and restores old value
 */
SocialCalc.Popup.Cancel = function() {
   const sp  = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;

   if (!sp.Current.id) return;

   const type = spc[sp.Current.id].type;
   const pt   = spt[type];

   pt.Cancel(type, sp.Current.id);

   sp.Current.id = null;
};

/**
 * ele = SocialCalc.Popup.CreatePopupDiv(id, attribs)
 *
 * Utility function to create the main popup div of width attribs.width.
 * If attribs.title, create one with that text, and optionally attribs.moveable.
 *
 * @param {string} id
 * @param {object} attribs
 * @returns {HTMLElement} main
 */
SocialCalc.Popup.CreatePopupDiv = function(id, attribs) {
   let pos;

   const sp      = SocialCalc.Popup;
   const spc     = sp.Controls;
   const spcdata = spc[id].data;

   const main = document.createElement("div");
   main.style.position       = "absolute";

   pos = SocialCalc.GetElementPositionWithScroll(spcdata.mainele);

   main.style.top            = (pos.top + spcdata.mainele.offsetHeight) + "px";
   main.style.left           = pos.left + "px";
   main.style.zIndex         = 100;
   main.style.backgroundColor = "#FFF";
   main.style.border         = "1px solid black";

   if (attribs.width) {
      main.style.width = attribs.width;
   }

   spcdata.mainele.appendChild(main);

   if (attribs.title) {
      main.innerHTML =
         '<table cellspacing="0" cellpadding="0" style="border-bottom:1px solid black;"><tr>' +
            '<td style="font-size:10px;cursor:default;width:100%;background-color:#999;color:#FFF;">' +
               attribs.title +
            '</td>' +
            '<td style="font-size:10px;cursor:default;color:#666;" onclick="SocialCalc.Popup.Cancel();">&nbsp;X&nbsp;</td>' +
         '</tr></table>';

      if (attribs.moveable) {
         spcdata.dragregistered = main.firstChild.firstChild.firstChild.firstChild;
         SocialCalc.DragRegister(
            spcdata.dragregistered,
            true,
            true,
            {
               MouseDown : SocialCalc.DragFunctionStart,
               MouseMove : SocialCalc.DragFunctionPosition,
               MouseUp   : SocialCalc.DragFunctionPosition,
               Disabled  : null,
               positionobj: main
            }
         );
      }
   }

   return main;
};

/**
 * SocialCalc.Popup.EnsurePosition(id, container)
 *
 * Utility function to make sure popup is positioned completely within container (both element objects)
 * and appropriate with respect to the main element controlling the popup.
 *
 * @param {string} id
 * @param {HTMLElement} container
 */
SocialCalc.Popup.EnsurePosition = function(id, container) {
   const sp      = SocialCalc.Popup;
   const spc     = sp.Controls;
   const spcdata = spc[id].data;

   const main = spcdata.mainele.firstChild;
   if (!main) { alert("No main popup element firstChild."); return; }

   const popup = spcdata.popupele;

   function GetLayoutValues(ele) {
      const r   = SocialCalc.GetElementPositionWithScroll(ele);
      r.height  = ele.offsetHeight;
      r.width   = ele.offsetWidth;
      r.bottom  = r.top + r.height;
      r.right   = r.left + r.width;
      return r;
   }

   const p = GetLayoutValues(popup);
   const c = GetLayoutValues(container);
   const m = GetLayoutValues(main);
   let   t = 0; // type of placement
   //addmsg("popup t/r/b/l/h/w= "+p.top+"/"+p.right+"/"+p.bottom+"/"+p.left+"/"+p.height+"/"+p.width);
   //addmsg("container t/r/b/l/h/w= "+c.top+"/"+c.right+"/"+c.bottom+"/"+c.left+"/"+c.height+"/"+c.width);
   //addmsg("main t/r/b/l/h/w= "+m.top+"/"+m.right+"/"+m.bottom+"/"+m.left+"/"+m.height+"/"+m.width);

   // Check various layout cases in priority order

   if (m.bottom + p.height < c.bottom && m.left + p.width < c.right) { // normal case: room on bottom and right
      popup.style.top  = m.bottom + "px";
      popup.style.left = m.left   + "px";
      t = 1;
   }
   else if (m.top - p.height > c.top && m.left + p.width < c.right) { // room on top and right
      popup.style.top  = (m.top - p.height) + "px";
      popup.style.left = m.left + "px";
      t = 2;
   }
   else if (m.bottom + p.height < c.bottom && m.right - p.width > c.left) { // room on bottom and left
      popup.style.top  = m.bottom + "px";
      popup.style.left = (m.right - p.width) + "px";
      t = 3;
   }
   else if (m.top - p.height > c.top && m.right - p.width > c.left) { // room on top and left
      popup.style.top  = (m.top - p.height) + "px";
      popup.style.left = (m.right - p.width) + "px";
      t = 4;
   }
   else if (m.bottom + p.height < c.bottom && p.width < c.width) { // room on bottom and middle
      popup.style.top  = m.bottom + "px";
      popup.style.left = (c.left + Math.floor((c.width - p.width) / 2)) + "px";
      t = 5;
   }
   else if (m.top - p.height > c.top && p.width < c.width) { // room on top and middle
      popup.style.top  = (m.top - p.height) + "px";
      popup.style.left = (c.left + Math.floor((c.width - p.width) / 2)) + "px";
      t = 6;
   }
   else if (p.height < c.height && m.right + p.width < c.right) { // room on middle and right
      popup.style.top  = (c.top + Math.floor((c.height - p.height) / 2)) + "px";
      popup.style.left = m.right + "px";
      t = 7;
   }
   else if (p.height < c.height && m.left - p.width > c.left) { // room on middle and left
      popup.style.top  = (c.top + Math.floor((c.height - p.height) / 2)) + "px";
      popup.style.left = (m.left - p.width) + "px";
      t = 8;
   }
   else { /* nothing works, so leave as it is */ }
   //addmsg("Popup layout "+t);
};

/**
 * ele = SocialCalc.Popup.DestroyPopupDiv(ele, dragregistered)
 *
 * Utility function to get rid of the main popup div.
 *
 * @param {HTMLElement} ele
 * @param {HTMLElement} dragregistered
 */
SocialCalc.Popup.DestroyPopupDiv = function(ele, dragregistered) {
   if (!ele) return;

   ele.innerHTML = "";

   SocialCalc.DragUnregister(dragregistered); // OK to do this even if not registered

   if (ele.parentNode) {
      ele.parentNode.removeChild(ele);
   }
};

/**
 * Color Utility Functions
 */

/**
 * Converts RGB color values to hexadecimal format
 * @param {string} val - RGB color string (e.g., "rgb(255,255,255)")
 * @returns {string} Hexadecimal color string (e.g., "FFFFFF")
 */
SocialCalc.Popup.RGBToHex = function(val) {
   const sp = SocialCalc.Popup;

   if (val == "") {
      return "000000";
   }
   const rgbvals = val.match(/(\d+)\D+(\d+)\D+(\d+)/);
   if (rgbvals) {
      return sp.ToHex(rgbvals[1]) + sp.ToHex(rgbvals[2]) + sp.ToHex(rgbvals[3]);
   }
   else {
      return "000000";
   }
};

SocialCalc.Popup.HexDigits = "0123456789ABCDEF";

/**
 * Converts a number to hexadecimal representation
 * @param {number} num - Number to convert
 * @returns {string} Two-digit hexadecimal string
 */
SocialCalc.Popup.ToHex = function(num) {
   const sp = SocialCalc.Popup;
   const first = Math.floor(num / 16);
   const second = num % 16;
   return sp.HexDigits.charAt(first) + sp.HexDigits.charAt(second);
};

/**
 * Converts hexadecimal string to decimal number
 * @param {string} str - Two-character hexadecimal string
 * @returns {number} Decimal representation
 */
SocialCalc.Popup.FromHex = function(str) {
   const sp = SocialCalc.Popup;
   const first = sp.HexDigits.indexOf(str.charAt(0).toUpperCase());
   const second = sp.HexDigits.indexOf(str.charAt(1).toUpperCase());
   return ((first >= 0) ? first : 0) * 16 + ((second >= 0) ? second : 0);
};

/**
 * Converts hexadecimal color to RGB format
 * @param {string} val - Hexadecimal color string (e.g., "#FFFFFF")
 * @returns {string} RGB color string (e.g., "rgb(255,255,255)")
 */
SocialCalc.Popup.HexToRGB = function(val) {
   const sp = SocialCalc.Popup;

   return "rgb(" + sp.FromHex(val.substring(1, 3)) + "," + sp.FromHex(val.substring(3, 5)) + "," + sp.FromHex(val.substring(5, 7)) + ")";
};

/**
 * Creates RGB color string from individual red, green, blue values
 * @param {number} r - Red value
 * @param {number} g - Green value
 * @param {number} b - Blue value
 * @returns {string} RGB color string
 */
SocialCalc.Popup.makeRGB = function(r, g, b) {
   return "rgb(" + (r > 0 ? r : 0) + "," + (g > 0 ? g : 0) + "," + (b > 0 ? b : 0) + ")";
};

/**
 * Splits RGB color string into individual color components
 * @param {string} rgb - RGB color string
 * @returns {object} Object with r, g, b properties
 */
SocialCalc.Popup.splitRGB = function(rgb) {
   const parts = rgb.match(/(\d+)\D+(\d+)\D+(\d+)\D/);
   if (!parts) {
      return { r: 0, g: 0, b: 0 };
   }
   else {
      return { r: parts[1] - 0, g: parts[2] - 0, b: parts[3] - 0 };
   }
};

// * * * * * * * * * * * * * * * *
//
// ROUTINES FOR EACH TYPE
//
// * * * * * * * * * * * * * * * *

/**
 * List
 *
 * type: List
 * value: value of control,
 * display: "value to display",
 * custom: true if custom value,
 * disabled: t/f,
 * attribs: {
 *    title: "popup title string",
 *    moveable: t/f,
 *    width: optional width, e.g., "100px",
 *    ensureWithin: optional element object to ensure popup fits within if possible
 *    changedcallback: optional function(attribs, id, newvalue),
 *    ...
 *    }
 * data: {
 *    ncols: calculated number of columns
 *    options: [
 *       {o: option-name, v: value-to-return,
 *        a: {option attribs} // optional: {skip: true, custom: true, cancel: true, newcol: true}
 *       },
 *       ...]
 *    }
 *
 * popupele: gets popup element object when created
 * contentele: gets element created with all the content
 * listdiv: gets div with list of items
 * customele: gets input element with custom value
 * dragregistered: gets element, if any, registered as draggable
 */

SocialCalc.Popup.Types.List = {};

/**
 * Creates a List control
 * @param {string} type - Control type
 * @param {string} id - Element ID
 * @param {object} attribs - Control attributes
 */
SocialCalc.Popup.Types.List.Create = function(type, id, attribs) {
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;

   const spcid = { type: type, value: "", display: "", data: {} };
   if (spc[id]) { 
       console.warn("SocialCalc List element already created:", id); 
       return; 
   }
   spc[id] = spcid;
   const spcdata = spcid.data;

   spcdata.attribs = attribs || {};

   const ele = document.getElementById(id);
   if (!ele) { alert("Missing element " + id); return; }

   spcdata.mainele = ele;

   ele.innerHTML = '<input style="cursor:pointer;width:' + (spcdata.attribs.inputWidth || '100px') + ';font-size:smaller;" onfocus="this.blur();" onclick="SocialCalc.Popup.CClick(\'' + id + '\');" value="">';

   spcdata.options = []; // set to nothing - use Initialize to fill
};

/**
 * Sets the value of a List control
 * @param {string} type - Control type
 * @param {string} id - Element ID
 * @param {*} value - New value
 */
SocialCalc.Popup.Types.List.SetValue = function(type, id, value) {
   let i;
   let o;

   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   spcdata.value = value;
   spcdata.custom = false;

   for (i = 0; i < spcdata.options.length; i++) {
      o = spcdata.options[i];
      if (o.a) {
         if (o.a.skip || o.a.custom || o.a.cancel) {
            continue;
         }
      }
      if (o.v == spcdata.value) { // matches value
         spcdata.display = o.o;
         break;
      }
   }
   if (i == spcdata.options.length) { // none found
      spcdata.display = "Custom";
      spcdata.custom = true;
   }

   if (spcdata.mainele && spcdata.mainele.firstChild) {
      spcdata.mainele.firstChild.value = spcdata.display;
   }
};

/**
 * Sets the disabled state of a List control
 * @param {string} type - Control type
 * @param {string} id - Element ID
 * @param {boolean} disabled - Whether to disable the control
 */
SocialCalc.Popup.Types.List.SetDisabled = function(type, id, disabled) {
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   spcdata.disabled = disabled;

   if (spcdata.mainele && spcdata.mainele.firstChild) {
      spcdata.mainele.firstChild.disabled = disabled;
   }
};

/**
 * Gets the current value of a List control
 * @param {string} type - Control type
 * @param {string} id - Element ID
 * @returns {*} Current value
 */
SocialCalc.Popup.Types.List.GetValue = function(type, id) {
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   return spcdata.value;
};

/**
 * Initializes a List control with data
 * data is: {value: initial value, attribs: {attribs stuff}, options: [{o: option-name, v: value-to-return, a: optional-attribs}, ...]}
 * @param {string} type - Control type
 * @param {string} id - Element ID
 * @param {object} data - Initialization data
 */
SocialCalc.Popup.Types.List.Initialize = function(type, id, data) {
   let a;

   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   for (a in data.attribs) {
      spcdata.attribs[a] = data.attribs[a];
   }

   spcdata.options = data ? data.options : [];

   if (data.value) { // if has a value, set to it
      sp.SetValue(id, data.value);
   }
};

/**
 * Resets the List popup state
 * @param {string} type - Control type
 */
SocialCalc.Popup.Types.List.Reset = function(type) {
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;

   if (sp.Current.id && spc[sp.Current.id].type == type) { // we have a popup
      spt[type].Hide(type, sp.Current.id);
      sp.Current.id = null;
   }
};

/**
 * Shows the List popup
 * @param {string} type - Control type
 * @param {string} id - Element ID
 */
SocialCalc.Popup.Types.List.Show = function(type, id) {
   let i, ele, o, bg;

   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   let str = "";

   spcdata.popupele = sp.CreatePopupDiv(id, spcdata.attribs);

   if (spcdata.custom) {
      str = SocialCalc.Popup.Types.List.MakeCustom(type, id);

      ele = document.createElement("div");
      ele.innerHTML = '<div style="cursor:default;padding:4px;background-color:#CCC;">' + str + '</div>';

      spcdata.customele = ele.firstChild.firstChild.childNodes[1];
      spcdata.listdiv = null;
      spcdata.contentele = ele;
   }
   else {
      str = SocialCalc.Popup.Types.List.MakeList(type, id);

      ele = document.createElement("div");
      ele.innerHTML = '<div style="cursor:default;padding:4px;">' + str + '</div>';

      spcdata.customele = null;
      spcdata.listdiv = ele.firstChild;
      spcdata.contentele = ele;
   }

   if (spcdata.mainele && spcdata.mainele.firstChild) {
      spcdata.mainele.firstChild.disabled = true;
   }

   spcdata.popupele.appendChild(ele);

   if (spcdata.attribs.ensureWithin) {
      SocialCalc.Popup.EnsurePosition(id, spcdata.attribs.ensureWithin);
   }
};

/**
 * Creates the HTML for the list display
 * @param {string} type - Control type
 * @param {string} id - Element ID
 * @returns {string} HTML string for the list
 */
SocialCalc.Popup.Types.List.MakeList = function(type, id) {
   let i, ele, o, bg;

   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   let str = '<table cellspacing="0" cellpadding="0"><tr>';
   const td = '<td style="vertical-align:top;">';

   str += td;

   spcdata.ncols = 1;

   for (i = 0; i < spcdata.options.length; i++) {
      o = spcdata.options[i];
      if (o.a) {
         if (o.a.newcol) {
            str += '</td>' + td + "&nbsp;&nbsp;&nbsp;&nbsp;" + '</td>' + td;
            spcdata.ncols += 1;
            continue;
         }
         if (o.a.skip) {
            str += '<div style="font-size:x-small;white-space:nowrap;">' + o.o + '</div>';
            continue;
         }
      }
      if (o.v == spcdata.value && !(o.a && (o.a.custom || o.a.cancel))) { // matches value
         bg = "background-color:#DDF;";
      }
      else {
         bg = "";
      }
      str += '<div style="font-size:x-small;white-space:nowrap;' + bg + '" onclick="SocialCalc.Popup.Types.List.ItemClicked(\'' + id + '\',\'' + i + '\');" onmousemove="SocialCalc.Popup.Types.List.MouseMove(\'' + id + '\',this);">' + o.o + '</div>';
   }

   str += "</td></tr></table>";

   return str;
};

/**
 * Creates the HTML for the custom input display
 * @param {string} type - Control type
 * @param {string} id - Element ID
 * @returns {string} HTML string for the custom input
 */
SocialCalc.Popup.Types.List.MakeCustom = function(type, id) {
   const SPLoc = SocialCalc.Popup.LocalizeString;

   let i, ele, o, bg;

   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   const style = 'style="font-size:smaller;"';

   let str = "";

   let val = spcdata.value;
   val = SocialCalc.special_chars(val);

   str = '<div style="white-space:nowrap;"><br>' +
         '<input id="customvalue" value="' + val + '"><br><br>' +
         '<input ' + style + ' type="button" value="' + SPLoc("OK") + '" onclick="SocialCalc.Popup.Types.List.CustomOK(\'' + id + '\');return false;">' +
         '<input ' + style + ' type="button" value="' + SPLoc("List") + '" onclick="SocialCalc.Popup.Types.List.CustomToList(\'' + id + '\');">' +
         '<input ' + style + ' type="button" value="' + SPLoc("Cancel") + '" onclick="SocialCalc.Popup.Close();">' +
         '<br></div>';

   return str;
};

/**
 * Handles when a list item is clicked
 * @param {string} id - Element ID
 * @param {string} num - Index of the clicked item
 */
SocialCalc.Popup.Types.List.ItemClicked = function(id, num) {
   let oele, str, nele;
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   const a = spcdata.options[num].a;

   if (a && a.custom) {
      oele = spcdata.contentele;
      str = SocialCalc.Popup.Types.List.MakeCustom("List", id);
      nele = document.createElement("div");
      nele.innerHTML = '<div style="cursor:default;padding:4px;background-color:#CCC;">' + str + '</div>';
      spcdata.customele = nele.firstChild.firstChild.childNodes[1];
      spcdata.listdiv = null;
      spcdata.contentele = nele;
      spcdata.popupele.replaceChild(nele, oele);
      if (spcdata.attribs.ensureWithin) {
         SocialCalc.Popup.EnsurePosition(id, spcdata.attribs.ensureWithin);
      }
      return;
   }

   if (a && a.cancel) {
      SocialCalc.Popup.Close();
      return;
   }

   SocialCalc.Popup.SetValue(id, spcdata.options[num].v);

   SocialCalc.Popup.Close();
};

/**
 * Switches from custom input back to list view
 * @param {string} id - Element ID
 */
SocialCalc.Popup.Types.List.CustomToList = function(id) {
   let oele, str, nele;
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   oele = spcdata.contentele;
   str = SocialCalc.Popup.Types.List.MakeList("List", id);
   nele = document.createElement("div");
   nele.innerHTML = '<div style="cursor:default;padding:4px;">' + str + '</div>';
   spcdata.customele = null;
   spcdata.listdiv = nele.firstChild;
   spcdata.contentele = nele;
   spcdata.popupele.replaceChild(nele, oele);
   
   if (spcdata.attribs.ensureWithin) {
      SocialCalc.Popup.EnsurePosition(id, spcdata.attribs.ensureWithin);
   }
};

/**
 * Handles OK button click in custom input mode
 * @param {string} id - Element ID
 */
SocialCalc.Popup.Types.List.CustomOK = function(id) {
   let i, c;
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   SocialCalc.Popup.SetValue(id, spcdata.customele.value);

   SocialCalc.Popup.Close();
};

/**
 * Handles mouse movement over list items for highlighting
 * @param {string} id - Element ID
 * @param {HTMLElement} ele - The element being hovered
 */
SocialCalc.Popup.Types.List.MouseMove = function(id, ele) {
   let col, i, c;
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   const list = spcdata.listdiv;

   if (!list) return;

   const rowele = list.firstChild.firstChild.firstChild; // div.table.tbody.tr

   for (col = 0; col < spcdata.ncols; col++) {
      for (i = 0; i < rowele.childNodes[col * 2].childNodes.length; i++) {
         rowele.childNodes[col * 2].childNodes[i].style.backgroundColor = "#FFF";
      }
   }

   ele.style.backgroundColor = "#DDF";
};

/**
 * Hides the List popup
 * @param {string} type - Control type
 * @param {string} id - Element ID
 */
SocialCalc.Popup.Types.List.Hide = function(type, id) {
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   sp.DestroyPopupDiv(spcdata.popupele, spcdata.dragregistered);
   spcdata.popupele = null;

   if (spcdata.mainele && spcdata.mainele.firstChild) {
      spcdata.mainele.firstChild.disabled = false;
   }
};

/**
 * Cancels the List popup without saving changes
 * @param {string} type - Control type
 * @param {string} id - Element ID
 */
SocialCalc.Popup.Types.List.Cancel = function(type, id) {
   SocialCalc.Popup.Types.List.Hide(type, id);
};

/**
 * ColorChooser
 *
 * type: ColorChooser
 * value: value of control as "rgb(r,g,b)" or "" if default,
 * oldvalue: starting value to reset to on close,
 * display: "value to display" as hex color value,
 * custom: true if custom value,
 * disabled: t/f,
 * attribs: {
 *    title: "popup title string",
 *    moveable: t/f,
 *    width: optional width, e.g., "100px", of popup chooser
 *    ensureWithin: optional element object to ensure popup fits within if possible
 *    sampleWidth: optional width, e.g., "20px",
 *    sampleHeight: optional height, e.g., "20px",
 *    backgroundImage: optional background image for sample (transparent where want to show current color), e.g., "colorbg.gif"
 *    backgroundImageDefault: optional background image for sample when default (transparent shows white)
 *    backgroundImageDisabled: optional background image for sample when disabled (transparent shows gray)
 *    changedcallback: optional function(attribs, id, newvalue),
 *    ...
 *    }
 * data: {
 *    }
 *
 * popupele: gets popup element object when created
 * contentele: gets element created with all the content
 * customele: gets input element with custom value
 */

SocialCalc.Popup.Types.ColorChooser = {};

/**
 * Creates a ColorChooser control
 * @param {string} type - Control type
 * @param {string} id - Element ID
 * @param {object} attribs - Control attributes
 */
SocialCalc.Popup.Types.ColorChooser.Create = function(type, id, attribs) {
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;

   const spcid = { type: type, value: "", display: "", data: {} };
   if (spc[id]) { 
       console.warn("SocialCalc ColorChooser element already created:", id); 
       return; 
   }
   spc[id] = spcid;
   const spcdata = spcid.data;

   spcdata.attribs = attribs || {};
   const spca = spcdata.attribs;

   spcdata.value = "";

   const ele = document.getElementById(id);
   if (!ele) { alert("Missing element " + id); return; }

   spcdata.mainele = ele;

   ele.innerHTML = '<div style="cursor:pointer;border:1px solid black;vertical-align:top;width:' +
                   (spca.sampleWidth || '15px') + ';height:' + (spca.sampleHeight || '15px') +
                   ';" onclick="SocialCalc.Popup.Types.ColorChooser.ControlClicked(\'' + id + '\');">&nbsp;</div>';
};

/**
 * Sets the value of a ColorChooser control
 * @param {string} type - Control type
 * @param {string} id - Element ID
 * @param {string} value - New color value
 */
SocialCalc.Popup.Types.ColorChooser.SetValue = function(type, id, value) {
   let i, img, pos;

   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;
   const spca = spcdata.attribs;

   spcdata.value = value;
   spcdata.custom = false;

   if (spcdata.mainele && spcdata.mainele.firstChild) {
      if (spcdata.value) {
         spcdata.mainele.firstChild.style.backgroundColor = spcdata.value;
         if (spca.backgroundImage) {
            img = "url(" + sp.imagePrefix + spca.backgroundImage + ")";
         }
         else {
            img = "";
         }
         pos = "center center";
      }
      else {
         spcdata.mainele.firstChild.style.backgroundColor = "#FFF";
         if (spca.backgroundImageDefault) {
            img = "url(" + sp.imagePrefix + spca.backgroundImageDefault + ")";
            pos = "center center";
         }
         else {
            img = "url(" + sp.imagePrefix + "defaultcolor.gif)";
            pos = "left top";
         }
      }
      spcdata.mainele.firstChild.style.backgroundPosition = pos;
      spcdata.mainele.firstChild.style.backgroundImage = img;
   }
};

/**
 * Sets the disabled state of a ColorChooser control
 * @param {string} type - Control type
 * @param {string} id - Element ID
 * @param {boolean} disabled - Whether to disable the control
 */
SocialCalc.Popup.Types.ColorChooser.SetDisabled = function(type, id, disabled) {
   let i, img, pos;

   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;
   const spca = spcdata.attribs;

   spcdata.disabled = disabled;

   if (spcdata.mainele && spcdata.mainele.firstChild) {
      if (disabled) {
         spcdata.mainele.firstChild.style.backgroundColor = "#DDD";
         if (spca.backgroundImageDisabled) {
            img = "url(" + sp.imagePrefix + spca.backgroundImageDisabled + ")";
            pos = "center center";
         }
         else {
            img = "url(" + sp.imagePrefix + "defaultcolor.gif)";
            pos = "left top";
         }
         spcdata.mainele.firstChild.style.backgroundPosition = pos;
         spcdata.mainele.firstChild.style.backgroundImage = img;
      }
      else {
         sp.SetValue(id, spcdata.value);
      }
   }
};

/**
 * Gets the current value of a ColorChooser control
 * @param {string} type - Control type
 * @param {string} id - Element ID
 * @returns {string} Current color value
 */
SocialCalc.Popup.Types.ColorChooser.GetValue = function(type, id) {
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   return spcdata.value;
};

/**
 * Initializes a ColorChooser control with data
 * @param {string} type - Control type
 * @param {string} id - Element ID
 * @param {object} data - Initialization data
 */
SocialCalc.Popup.Types.ColorChooser.Initialize = function(type, id, data) {
   let a;

   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   for (a in data.attribs) {
      spcdata.attribs[a] = data.attribs[a];
   }

   if (data.value) { // if has a value, set to it
      sp.SetValue(id, data.value);
   }
};

/**
 * Resets the ColorChooser popup state
 * @param {string} type - Control type
 */
SocialCalc.Popup.Types.ColorChooser.Reset = function(type) {
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;

   if (sp.Current.id && spc[sp.Current.id].type == type) { // we have a popup
      spt[type].Hide(type, sp.Current.id);
      sp.Current.id = null;
   }
};

/**
 * Shows the ColorChooser popup
 * @param {string} type - Control type
 * @param {string} id - Element ID
 */
SocialCalc.Popup.Types.ColorChooser.Show = function(type, id) {
   let i, ele, mainele;

   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   let str = "";

   spcdata.oldvalue = spcdata.value; // remember starting value

   spcdata.popupele = sp.CreatePopupDiv(id, spcdata.attribs);

   if (spcdata.custom) {
      str = SocialCalc.Popup.Types.ColorChooser.MakeCustom(type, id);

      ele = document.createElement("div");
      ele.innerHTML = '<div style="cursor:default;padding:4px;background-color:#CCC;">' + str + '</div>';

      spcdata.customele = ele.firstChild.firstChild.childNodes[2];
      spcdata.contentele = ele;
   }
   else {
      mainele = SocialCalc.Popup.Types.ColorChooser.CreateGrid(type, id);

      ele = document.createElement("div");
      ele.style.padding = "3px";
      ele.style.backgroundColor = "#CCC";
      ele.appendChild(mainele);

      spcdata.customele = null;
      spcdata.contentele = ele;
   }

   spcdata.popupele.appendChild(ele);

   if (spcdata.attribs.ensureWithin) {
      SocialCalc.Popup.EnsurePosition(id, spcdata.attribs.ensureWithin);
   }
};

/**
 * Creates the HTML for the custom color input display
 * @param {string} type - Control type
 * @param {string} id - Element ID
 * @returns {string} HTML string for the custom input
 */
SocialCalc.Popup.Types.ColorChooser.MakeCustom = function(type, id) {
   let i, ele, o, bg;

   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   const SPLoc = sp.LocalizeString;

   const style = 'style="font-size:smaller;"';

   let str = "";

   str = '<div style="white-space:nowrap;"><br>' +
         '#<input id="customvalue" style="width:75px;" value="' + spcdata.value + '"><br><br>' +
         '<input ' + style + ' type="button" value="' + SPLoc("OK") + '" onclick="SocialCalc.Popup.Types.ColorChooser.CustomOK(\'' + id + '\');return false;">' +
         '<input ' + style + ' type="button" value="' + SPLoc("Grid") + '" onclick="SocialCalc.Popup.Types.ColorChooser.CustomToGrid(\'' + id + '\');">' +
         '<br></div>';

   return str;
};

/**
 * Handles when a color grid item is clicked
 * @param {string} id - Element ID
 * @param {string} num - Index of the clicked item
 */
SocialCalc.Popup.Types.ColorChooser.ItemClicked = function(id, num) {
   let oele, str, nele;
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   SocialCalc.Popup.Close();
};

/**
 * Switches from custom input to list view (placeholder function)
 * @param {string} id - Element ID
 */
SocialCalc.Popup.Types.ColorChooser.CustomToList = function(id) {
   let oele, str, nele;
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;
};

/**
 * Handles OK button click in custom input mode
 * @param {string} id - Element ID
 */
SocialCalc.Popup.Types.ColorChooser.CustomOK = function(id) {
   let i, c;
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   sp.SetValue(id, spcdata.customele.value);

   sp.Close();
};

/**
 * Hides the ColorChooser popup
 * @param {string} type - Control type
 * @param {string} id - Element ID
 */
SocialCalc.Popup.Types.ColorChooser.Hide = function(type, id) {
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   sp.DestroyPopupDiv(spcdata.popupele, spcdata.dragregistered);
   spcdata.popupele = null;
};

/**
 * Cancels the ColorChooser popup and restores old value
 * @param {string} type - Control type
 * @param {string} id - Element ID
 */
SocialCalc.Popup.Types.ColorChooser.Cancel = function(type, id) {
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const spcdata = spc[id].data;

   sp.SetValue(id, spcdata.oldvalue); // reset to old value

   SocialCalc.Popup.Types.ColorChooser.Hide(type, id);
};

/**
 * Creates the color grid interface for the ColorChooser
 * @param {string} type - Control type
 * @param {string} id - Element ID
 * @returns {HTMLElement} Main grid element
 */
SocialCalc.Popup.Types.ColorChooser.CreateGrid = function(type, id) {
   let ele, pos, row, rowele, col, g;

   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const spc = sp.Controls;
   const SPLoc = sp.LocalizeString;
   const spcdata = spc[id].data;
   spcdata.grid = {};
   const grid = spcdata.grid;

   const mainele = document.createElement("div");

   ele = document.createElement("table");
   ele.cellSpacing = 0;
   ele.cellPadding = 0;
   ele.style.width = "100px";
   grid.table = ele;

   ele = document.createElement("tbody");
   grid.table.appendChild(ele);
   grid.tbody = ele;

   for (row = 0; row < 16; row++) {
      rowele = document.createElement("tr");
      for (col = 0; col < 5; col++) {
         g = {};
         grid[row + "," + col] = g;
         ele = document.createElement("td");
         ele.style.fontSize = "1px";
         ele.innerHTML = "&nbsp;";
         ele.style.height = "10px";
         if (col <= 1) {
            ele.style.width = "17px";
            ele.style.borderRight = "3px solid white";
         }
         else {
            ele.style.width = "20px";
            ele.style.backgroundRepeat = "no-repeat";
         }
         rowele.appendChild(ele);
         g.ele = ele;
      }
      grid.tbody.appendChild(rowele);
   }
   mainele.appendChild(grid.table);

   ele = document.createElement("div");
   ele.style.marginTop = "3px";
   ele.innerHTML = '<table cellspacing="0" cellpadding="0"><tr>' +
      '<td style="width:17px;background-color:#FFF;background-image:url(' + sp.imagePrefix + 'defaultcolor.gif);height:16px;font-size:10px;cursor:pointer;" title="' + SPLoc("Default") + '">&nbsp;</td>' +
      '<td style="width:23px;height:16px;font-size:10px;text-align:center;cursor:pointer;" title="' + SPLoc("Custom") + '">#</td>' +
      '<td style="width:60px;height:16px;font-size:10px;text-align:center;cursor:pointer;">' + SPLoc("OK") + '</td>' +
      '</tr></table>';
   grid.defaultbox = ele.firstChild.firstChild.firstChild.childNodes[0];
   grid.defaultbox.onclick = spt.ColorChooser.DefaultClicked;
   grid.custom = ele.firstChild.firstChild.firstChild.childNodes[1];
   grid.custom.onclick = spt.ColorChooser.CustomClicked;
   grid.msg = ele.firstChild.firstChild.firstChild.childNodes[2];
   grid.msg.onclick = spt.ColorChooser.CloseOK;
   mainele.appendChild(ele);

   grid.table.onmousedown = spt.ColorChooser.GridMouseDown;

   spt.ColorChooser.DetermineColors(id);
   spt.ColorChooser.SetColors(id);

   return mainele;
};

/**
 * Gets a grid cell object by row and column
 * @param {object} grid - Grid object
 * @param {number} row - Row index
 * @param {number} col - Column index
 * @returns {object} Grid cell object
 */
SocialCalc.Popup.Types.ColorChooser.gridToG = function(grid, row, col) {
   return grid[row + "," + col];
};

/**
 * Determines the colors for each cell in the color grid
 * @param {string} id - Element ID
 */
SocialCalc.Popup.Types.ColorChooser.DetermineColors = function(id) {
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const sptc = spt.ColorChooser;
   const spc = sp.Controls;
   const spcdata = spc[id].data;
   const grid = spcdata.grid;

   let col, row;
   const rgb = sp.splitRGB(spcdata.value);
   let color;

   col = 2;
   row = 16 - Math.floor((rgb.r + 16) / 16);
   grid["selectedrow" + col] = row;
   for (row = 0; row < 16; row++) {
      sptc.gridToG(grid, row, col).rgb = sp.makeRGB(17 * (15 - row), 0, 0);
   }

   col = 3;
   row = 16 - Math.floor((rgb.g + 16) / 16);
   grid["selectedrow" + col] = row;
   for (row = 0; row < 16; row++) {
      sptc.gridToG(grid, row, col).rgb = sp.makeRGB(0, 17 * (15 - row), 0);
   }

   col = 4;
   row = 16 - Math.floor((rgb.b + 16) / 16);
   grid["selectedrow" + col] = row;
   for (row = 0; row < 16; row++) {
      sptc.gridToG(grid, row, col).rgb = sp.makeRGB(0, 0, 17 * (15 - row));
   }

   col = 1;
   for (row = 0; row < 16; row++) {
      sptc.gridToG(grid, row, col).rgb = sp.makeRGB(17 * (15 - row), 17 * (15 - row), 17 * (15 - row));
   }

   col = 0;
   const steps = [0, 68, 153, 204, 255];
   const commonrgb = ["400", "310", "420", "440", "442", "340", "040", "042", "032", "044", "024", "004", "204", "314", "402", "414"];
   let x;
   for (row = 0; row < 16; row++) {
      x = commonrgb[row];
      sptc.gridToG(grid, row, col).rgb = "rgb(" + steps[x.charAt(0) - 0] + "," + steps[x.charAt(1) - 0] + "," + steps[x.charAt(2) - 0] + ")";
   }
};

/**
 * Sets the colors for all cells in the color grid
 * @param {string} id - Element ID
 */
SocialCalc.Popup.Types.ColorChooser.SetColors = function(id) {
   let row, col, g, ele, rgb;

   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const sptc = spt.ColorChooser;
   const spc = sp.Controls;
   const spcdata = spc[id].data;
   const grid = spcdata.grid;

   for (row = 0; row < 16; row++) {
      for (col = 0; col < 5; col++) {
         g = sptc.gridToG(grid, row, col);
         g.ele.style.backgroundColor = g.rgb;
         g.ele.title = sp.RGBToHex(g.rgb);
         if (grid["selectedrow" + col] == row) {
            g.ele.style.backgroundImage = "url(" + sp.imagePrefix + "chooserarrow.gif)";
         }
         else {
            g.ele.style.backgroundImage = "";
         }
      }
   }

   sp.SetValue(id, spcdata.value);

   grid.msg.style.backgroundColor = spcdata.value;
   rgb = sp.splitRGB(spcdata.value || "rgb(255,255,255)");
   if (rgb.r + rgb.g + rgb.b < 220) {
      grid.msg.style.color = "#FFF";
   }
   else {
      grid.msg.style.color = "#000";
   }
   if (!spcdata.value) { // default
      grid.msg.style.backgroundColor = "#FFF";
      grid.msg.style.backgroundImage = "url(" + sp.imagePrefix + "defaultcolor.gif)";
      grid.msg.title = "Default";
   }
   else {
      grid.msg.style.backgroundImage = "";
      grid.msg.title = sp.RGBToHex(spcdata.value);
   }
};

/**
 * Handles mouse events on the color grid
 * @param {Event} e - Mouse event
 */
SocialCalc.Popup.Types.ColorChooser.GridMouseDown = function(e) {
   const event = e || window.event;

   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const sptc = spt.ColorChooser;
   const spc = sp.Controls;

   const id = sp.Current.id;
   if (!id) return;

   const spcdata = spc[id].data;
   const grid = spcdata.grid;

   switch (event.type) {
      case "mousedown":
         grid.mousedown = true;
         break;
      case "mouseup":
         grid.mousedown = false;
         break;
      case "mousemove":
         if (!grid.mousedown) {
            return;
         }
         break;
   }

   const viewport = SocialCalc.GetViewportInfo();
   const clientX = event.clientX + viewport.horizontalScroll;
   const clientY = event.clientY + viewport.verticalScroll;
   const gpos = SocialCalc.GetElementPosition(grid.table);
   let row = Math.floor((clientY - gpos.top - 2) / 10); // -2 is to split the diff btw IE & FF
   row = row < 0 ? 0 : row;
   let col = Math.floor((clientX - gpos.left) / 20);
   row = row < 0 ? 0 : (row > 15 ? 15 : row);
   col = col < 0 ? 0 : (col > 4 ? 4 : col);
   const color = sptc.gridToG(grid, row, col).ele.style.backgroundColor;
   const newrgb = sp.splitRGB(color);
   const oldrgb = sp.splitRGB(spcdata.value);

   switch (col) {
      case 2:
         spcdata.value = sp.makeRGB(newrgb.r, oldrgb.g, oldrgb.b);
         break;
      case 3:
         spcdata.value = sp.makeRGB(oldrgb.r, newrgb.g, oldrgb.b);
         break;
      case 4:
         spcdata.value = sp.makeRGB(oldrgb.r, oldrgb.g, newrgb.b);
         break;
      case 0:
      case 1:
         spcdata.value = color;
   }

   sptc.DetermineColors(id);
   sptc.SetColors(id);
};

/**
 * Handles clicks on the color control element
 * @param {string} id - Element ID
 */
SocialCalc.Popup.Types.ColorChooser.ControlClicked = function(id) {
   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const sptc = spt.ColorChooser;
   const spc = sp.Controls;

   const cid = sp.Current.id;
   if (!cid || id != cid) {
      sp.CClick(id);
      return;
   }

   sptc.CloseOK();
};

/**
 * Handles clicks on the default color button
 * @param {Event} e - Click event
 */
SocialCalc.Popup.Types.ColorChooser.DefaultClicked = function(e) {
   const event = e || window.event;

   const sp = SocialCalc.Popup;
   const spt = sp.Types;
   const sptc = spt.ColorChooser;
   const spc = sp.Controls;

   const id = sp.Current.id;
   if (!id) return;

   const spcdata = spc[id].data;

   spcdata.value = "";
   SocialCalc.Popup.SetValue(id, spcdata.value);

   SocialCalc.Popup.Close();
};

/**
 * Handles clicks on the “Custom” (#) button to switch the chooser
 * into custom-hex input mode.
 *
 * @param {Event} e – Mouse event
 */
SocialCalc.Popup.Types.ColorChooser.CustomClicked = function(e) {
   const event = e || window.event;

   const sp   = SocialCalc.Popup;
   const spt  = sp.Types;
   const sptc = spt.ColorChooser;
   const spc  = sp.Controls;

   const id = sp.Current.id;
   if (!id) return;

   const spcdata = spc[id].data;

   // Replace the current grid with the custom-value pane
   const oldEle = spcdata.contentele;
   const html   = SocialCalc.Popup.Types.ColorChooser.MakeCustom("ColorChooser", id);
   const newEle = document.createElement("div");
   newEle.innerHTML = '<div style="cursor:default;padding:4px;background-color:#CCC;">' + html + '</div>';

   spcdata.customele  = newEle.firstChild.firstChild.childNodes[2];
   spcdata.contentele = newEle;
   spcdata.popupele.replaceChild(newEle, oldEle);

   // Pre-populate hex field with current value
   spcdata.customele.value = sp.RGBToHex(spcdata.value);

   // Keep popup on-screen if requested
   if (spcdata.attribs.ensureWithin) {
      SocialCalc.Popup.EnsurePosition(id, spcdata.attribs.ensureWithin);
   }
};

/**
 * Converts the custom hex value to RGB, stores it, and returns to grid mode.
 *
 * @param {string} id – Element ID
 */
SocialCalc.Popup.Types.ColorChooser.CustomToGrid = function(id) {
   const sp   = SocialCalc.Popup;
   const spt  = sp.Types;
   const spc  = sp.Controls;
   const spcdata = spc[id].data;

   // Apply custom value
   SocialCalc.Popup.SetValue(id, sp.HexToRGB("#" + spcdata.customele.value));

   // Swap back to grid UI
   const oldEle = spcdata.contentele;
   const gridEle = SocialCalc.Popup.Types.ColorChooser.CreateGrid("ColorChooser", id);
   const wrapper = document.createElement("div");
   wrapper.style.padding = "3px";
   wrapper.style.backgroundColor = "#CCC";
   wrapper.appendChild(gridEle);

   spcdata.customele  = null;
   spcdata.contentele = wrapper;
   spcdata.popupele.replaceChild(wrapper, oldEle);

   // Re-position if needed
   if (spcdata.attribs.ensureWithin) {
      SocialCalc.Popup.EnsurePosition(id, spcdata.attribs.ensureWithin);
   }
};

/**
 * Commits the custom hex value and closes the popup.
 *
 * @param {string} id – Element ID
 */
SocialCalc.Popup.Types.ColorChooser.CustomOK = function(id) {
   const sp   = SocialCalc.Popup;
   const spc  = sp.Controls;
   const spcdata = spc[id].data;

   SocialCalc.Popup.SetValue(id, sp.HexToRGB("#" + spcdata.customele.value));
   SocialCalc.Popup.Close();
};

/**
 * Finalizes the selection (OK button inside the grid) and closes the popup.
 *
 * @param {Event} e – Click event
 */
SocialCalc.Popup.Types.ColorChooser.CloseOK = function(e) {
   const event = e || window.event;

   const sp   = SocialCalc.Popup;
   const spc  = sp.Controls;

   const id = sp.Current.id;
   if (!id) return;

   const spcdata = spc[id].data;

   SocialCalc.Popup.SetValue(id, spcdata.value);
   SocialCalc.Popup.Close();
};

