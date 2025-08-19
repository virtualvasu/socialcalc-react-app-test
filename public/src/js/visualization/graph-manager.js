/**
 * Graph functionality for SocialCalc
 * @fileoverview Contains graph configuration and core graph handling functions
 */

/**
 * Configuration object for available graph types and their display information
 * @type {Object}
 * @property {string[]} displayorder - Order in which graph types should be displayed
 * @property {Object} verticalbar - Vertical bar chart configuration
 * @property {Object} horizontalbar - Horizontal bar chart configuration
 * @property {Object} piechart - Pie chart configuration
 * @property {Object} linechart - Line chart configuration
 * @property {Object} scatterchart - Scatter chart configuration
 */
SocialCalc.GraphTypesInfo = {
    displayorder: ["verticalbar", "horizontalbar", "piechart", "linechart", "scatterchart"],

    verticalbar: { display: "Vertical Bar", func: GraphVerticalBar },
    horizontalbar: { display: "Horizontal Bar", func: GraphHorizontalBar },
    piechart: { display: "Pie Chart", func: MakePieChart },
    linechart: { display: "Line Chart", func: MakeLineChart },
    scatterchart: { display: "Plot Points", func: MakeScatterChart }
};

/**
 * Handles the graph button click event, initializes graph UI elements
 * @param {Object} s - Spreadsheet control object
 * @param {Object} t - Target element (unused)
 * @returns {void}
 */
function GraphOnClick(s, t) {
    const nameList = [];
    const nl = document.getElementById(`${s.idPrefix}graphlist`);
    
    s.editor.RangeChangeCallback.graph = UpdateGraphRangeProposal;
    
    // Collect and sort sheet names
    for (const name in s.sheet.names) {
        nameList.push(name);
    }
    nameList.sort();
    
    // Populate range selection dropdown
    nl.length = 0;
    nl.options[0] = new Option("[select range]");
    
    for (let i = 0; i < nameList.length; i++) {
        const name = nameList[i];
        nl.options[i + 1] = new Option(name, name);
        
        if (name === s.graphrange) {
            nl.options[i + 1].selected = true;
        }
    }
    
    if (s.graphrange === "") {
        nl.options[0].selected = true;
    }
    
    UpdateGraphRangeProposal(s.editor);
    
    // Populate graph type dropdown
    const typeList = document.getElementById(`${s.idPrefix}graphtype`);
    typeList.length = 0;
    
    for (let i = 0; i < SocialCalc.GraphTypesInfo.displayorder.length; i++) {
        const name = SocialCalc.GraphTypesInfo.displayorder[i];
        typeList.options[i] = new Option(SocialCalc.GraphTypesInfo[name].display, name);
        
        if (name === s.graphtype) {
            typeList.options[i].selected = true;
        }
    }
    
    if (!s.graphtype) {
        typeList.options[0].selected = true;
        s.graphtype = typeList.options[0].value;
    }
    
    DoGraph(false, true);
}

/**
 * Updates the graph range proposal display based on current editor selection
 * @param {Object} editor - The spreadsheet editor object
 * @returns {void}
 */
function UpdateGraphRangeProposal(editor) {
    const ele = document.getElementById(`${SocialCalc.GetSpreadsheetControlObject().idPrefix}graphlist`);
    
    if (editor.range.hasrange) {
        ele.options[0].text = `${SocialCalc.crToCoord(editor.range.left, editor.range.top)}:${SocialCalc.crToCoord(editor.range.right, editor.range.bottom)}`;
    } else {
        ele.options[0].text = "[select range]";
    }
}

/**
 * Sets the selected cell range for graphing
 * @returns {void}
 */
function GraphSetCells() {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const { editor } = spreadsheet;
    
    const lele = document.getElementById(`${spreadsheet.idPrefix}graphlist`);
    
    if (lele.selectedIndex === 0) {
        if (editor.range.hasrange) {
            spreadsheet.graphrange = `${SocialCalc.crToCoord(editor.range.left, editor.range.top)}:${SocialCalc.crToCoord(editor.range.right, editor.range.bottom)}`;
        } else {
            spreadsheet.graphrange = `${editor.ecell.coord}:${editor.ecell.coord}`;
        }
    } else {
        spreadsheet.graphrange = lele.options[lele.selectedIndex].value;
    }
    
    const ele = document.getElementById(`${spreadsheet.idPrefix}graphrange`);
    ele.innerHTML = spreadsheet.graphrange;
    
    DoGraph(false, false);
}

/**
 * Main function to render the graph based on current settings
 * @param {boolean} helpflag - Whether to show help information
 * @param {boolean} isResize - Whether this is a resize operation
 * @returns {void}
 */
function DoGraph(helpflag, isResize) {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const { editor } = spreadsheet;
    const gview = spreadsheet.views.graph.element;
    const ginfo = SocialCalc.GraphTypesInfo[spreadsheet.graphtype];
    const gfunc = ginfo.func;
    
    if (!spreadsheet.graphrange) {
        if (gfunc && helpflag) {
            gfunc(spreadsheet, null, gview, spreadsheet.graphtype, helpflag, isResize);
        } else {
            gview.innerHTML = '<div style="padding:30px;font-weight:bold;">Select a range of cells with numeric values to graph and use the OK button above to set the range as the graph range.</div>';
        }
        return;
    }
    
    let grange = spreadsheet.graphrange;
    
    // Handle named ranges
    if (grange && grange.indexOf(":") === -1) {
        const nrange = SocialCalc.Formula.LookupName(spreadsheet.sheet, grange || "");
        
        if (nrange.type !== "range") {
            gview.innerHTML = `Unknown range name: ${grange}`;
            return;
        }
        
        const rparts = nrange.value.match(/^(.*)\|(.*)\|$/);
        grange = `${rparts[1]}:${rparts[2]}`;
    }
    
    const prange = SocialCalc.ParseRange(grange);
    const range = {};
    
    // Normalize range coordinates
    if (prange.cr1.col <= prange.cr2.col) {
        range.left = prange.cr1.col;
        range.right = prange.cr2.col;
    } else {
        range.left = prange.cr2.col;
        range.right = prange.cr1.col;
    }
    
    if (prange.cr1.row <= prange.cr2.row) {
        range.top = prange.cr1.row;
        range.bottom = prange.cr2.row;
    } else {
        range.top = prange.cr2.row;
        range.bottom = prange.cr1.row;
    }
    
    if (gfunc) {
        gfunc(spreadsheet, range, gview, spreadsheet.graphtype, helpflag, isResize);
    }
}

/**
 * Handles graph type selection change
 * @param {HTMLSelectElement} gtobj - The graph type select element
 * @returns {void}
 */
function GraphChanged(gtobj) {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    spreadsheet.graphtype = gtobj.options[gtobj.selectedIndex].value;
    DoGraph(false, false);
}

/**
 * Handles min/max value changes for graph axes
 * @param {HTMLInputElement} minmaxobj - The input element that changed
 * @param {number} index - Index indicating which value changed (0=minX, 1=maxX, 2=minY, 3=maxY)
 * @returns {void}
 */
function MinMaxChanged(minmaxobj, index) {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    
    switch (index) {
        case 0:
            spreadsheet.graphMinX = minmaxobj.value;
            break;
        case 1:
            spreadsheet.graphMaxX = minmaxobj.value;
            break;
        case 2:
            spreadsheet.graphMinY = minmaxobj.value;
            break;
        case 3:
            spreadsheet.graphMaxY = minmaxobj.value;
            break;
    }
    
    DoGraph(false, true);
}

/**
 * Saves graph settings to a string format for persistence
 * @param {Object} editor - The spreadsheet editor object (unused)
 * @param {Object} setting - Settings object (unused)
 * @returns {string} Encoded graph settings string
 */
function GraphSave(editor, setting) {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const gtype = spreadsheet.graphtype || "";
    
    let str = `graph:range:${SocialCalc.encodeForSave(spreadsheet.graphrange)}:type:${SocialCalc.encodeForSave(gtype)}`;
    str += `:minmax:${SocialCalc.encodeForSave(`${spreadsheet.graphMinX},${spreadsheet.graphMaxX},${spreadsheet.graphMinY},${spreadsheet.graphMaxY}`)}\n`;
    
    return str;
}

/**
 * Loads graph settings from a saved string format
 * @param {Object} editor - The spreadsheet editor object (unused)
 * @param {Object} setting - Settings object (unused)
 * @param {string} line - The saved settings line to parse
 * @param {Object} flags - Flags object (unused)
 * @returns {boolean} Always returns true
 */
function GraphLoad(editor, setting, line, flags) {
    const spreadsheet = SocialCalc.GetSpreadsheetControlObject();
    const parts = line.split(":");
    
    for (let i = 1; i < parts.length; i += 2) {
        switch (parts[i]) {
            case "range":
                spreadsheet.graphrange = SocialCalc.decodeFromSave(parts[i + 1]);
                break;
            case "type":
                spreadsheet.graphtype = SocialCalc.decodeFromSave(parts[i + 1]);
                break;
            case "minmax":
                const splitMinMax = SocialCalc.decodeFromSave(parts[i + 1]).split(",");
                spreadsheet.graphMinX = splitMinMax[0];
                document.getElementById("SocialCalc-graphMinX").value = spreadsheet.graphMinX;
                spreadsheet.graphMaxX = splitMinMax[1];
                document.getElementById("SocialCalc-graphMaxX").value = spreadsheet.graphMaxX;
                spreadsheet.graphMinY = splitMinMax[2];
                document.getElementById("SocialCalc-graphMinY").value = spreadsheet.graphMinY;
                spreadsheet.graphMaxY = splitMinMax[3];
                document.getElementById("SocialCalc-graphMaxY").value = spreadsheet.graphMaxY;
                break;
        }
    }
    
    return true;
}

/**
 * Graphing Functions for SocialCalc
 * All graphing functions are called with (spreadsheet, range, gview, gtype, helpflag, isResize)
 */

/**
 * Generates a vertical bar chart from the selected range
 * @param {Object} spreadsheet - The spreadsheet control object
 * @param {Object|null} range - Range object with left, right, top, bottom properties
 * @param {HTMLElement} gview - The graph view container element
 * @param {string} gtype - The graph type identifier
 * @param {boolean} helpflag - Whether to display help text instead of chart
 * @param {boolean} [isResize] - Whether this is a resize operation (unused)
 * @returns {void}
 */
function GraphVerticalBar(spreadsheet, range, gview, gtype, helpflag) {
    const values = [];
    const labels = [];
    
    if (helpflag || !range) {
        const str = `<input type="button" value="Hide Help" onclick="DoGraph(false,false);"><br><br>
            This is the help text for graph type: ${SocialCalc.GraphTypesInfo[gtype].display}.<br><br>
            The <b>Graph</b> tab displays a bar graph of the cells which have been selected 
            (either in a single row across or column down). 
            If the row above (or column to the left) of the selection has values, those values are used as labels. 
            Otherwise the cells value is used as a label. 
            <br><br><input type="button" value="Hide Help" onclick="DoGraph(false,false);">`;
        
        gview.innerHTML = str;
        return;
    }
    
    // Determine if we're working with a column (byrow=true) or row (byrow=false)
    const byrow = range.left === range.right;
    const nitems = byrow ? range.bottom - range.top + 1 : range.right - range.left + 1;
    
    const maxheight = (spreadsheet.height - spreadsheet.nonviewheight) - 50;
    const totalwidth = spreadsheet.width - 30;
    
    let minval = null;
    let maxval = null;
    
    // Extract values and labels from the range
    for (let i = 0; i < nitems; i++) {
        const cr = byrow ? 
            `${SocialCalc.rcColname(range.left)}${i + range.top}` : 
            `${SocialCalc.rcColname(i + range.left)}${range.top}`;
        
        const cr1 = byrow ? 
            `${SocialCalc.rcColname(range.left - 1 || 1)}${i + range.top}` : 
            `${SocialCalc.rcColname(i + range.left)}${range.top - 1 || 1}`;
        
        const cell = spreadsheet.sheet.GetAssuredCell(cr);
        
        if (cell.valuetype.charAt(0) === "n") {
            const val = cell.datavalue - 0;
            
            if (maxval === null || maxval < val) maxval = val;
            if (minval === null || minval > val) minval = val;
            
            values.push(val);
            
            const labelCell = spreadsheet.sheet.GetAssuredCell(cr1);
            
            if ((range.right === range.left || range.top === range.bottom) && 
                (labelCell.valuetype.charAt(0) === "t" || labelCell.valuetype.charAt(0) === "n")) {
                labels.push(labelCell.datavalue + "");
            } else {
                labels.push(val + "");
            }
        }
    }
    
    // Ensure zero line is visible
    if (maxval < 0) maxval = 0;
    if (minval > 0) minval = 0;
    
    const str = '<table><tr><td><canvas id="myBarCanvas" width="500px" height="400px" style="border:1px solid black;"></canvas></td><td><span id="googleBarChart"></span></td></tr></table>';
    gview.innerHTML = str;
    
    const profChartVals = [];
    const profChartLabels = [];
    
    const canv = document.getElementById("myBarCanvas");
    const ctx = canv.getContext('2d');
    
    // Note: mozTextStyle is deprecated but kept for compatibility
    ctx.mozTextStyle = "10pt bold Arial";
    
    const canvasMaxHeight = canv.height - 60;
    const canvasTotalWidth = canv.width;
    const colors = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];
    
    /**
     * Generates a random hex color
     * @returns {string} Random 6-character hex color
     */
    const generateRandomColor = () => {
        return Array.from({ length: 6 }, () => colors[Math.round(Math.random() * 14)]).join('');
    };
    
    let barColor = generateRandomColor();
    ctx.fillStyle = `#${barColor}`;
    const colorList = [barColor];
    
    const eachwidth = Math.floor(canvasTotalWidth / (values.length || 1)) - 4 || 1;
    const zeroLine = canvasMaxHeight * (maxval / (maxval - minval)) + 30;
    
    // Draw zero line
    ctx.lineWidth = 5;
    ctx.moveTo(0, zeroLine);
    ctx.lineTo(canv.width, zeroLine);
    ctx.stroke();
    
    const yScale = canvasMaxHeight / (maxval - minval);
    
    // Draw bars
    for (let i = 0; i < values.length; i++) {
        ctx.fillRect(i * eachwidth, zeroLine - yScale * values[i], eachwidth, yScale * values[i]);
        profChartVals.push(Math.floor((values[i] - minval) * yScale / 3.4));
        profChartLabels.push(labels[i]);
        
        barColor = generateRandomColor();
        ctx.fillStyle = `#${barColor}`;
        colorList.push(barColor);
    }
    
    // Draw labels
    ctx.strokeStyle = "#000000";
    ctx.fillStyle = "#000000";
    
    if (values[0] > 0) {
        ctx.translate(5, zeroLine + 22);
    } else {
        ctx.translate(5, zeroLine - 15);
    }
    
    ctx.mozDrawText(labels[0]);
    
    for (let i = 1; i < values.length; i++) {
        if ((values[i] > 0) && (values[i - 1] < 0)) {
            ctx.translate(eachwidth, 37);
        } else if ((values[i] < 0) && (values[i - 1] > 0)) {
            ctx.translate(eachwidth, -37);
        } else {
            ctx.translate(eachwidth, 0);
        }
        ctx.mozDrawText(labels[i]);
    }
    
    // Generate Google Charts API URL
    const gChart = document.getElementById("googleBarChart");
    const googleZeroLine = (-1 * minval) * yScale / 340;
    const profChartUrl = `chs=300x250&cht=bvg&chd=t:${profChartVals.join(",")}&chxt=x,y&chxl=0:|${profChartLabels.join("|")}|&chxr=1,${minval},${maxval}&chp=${googleZeroLine}&chbh=a&chm=r,000000,0,${googleZeroLine},${googleZeroLine + 0.005}&chco=${colorList.join("|")}`;
    
    gChart.innerHTML = `<iframe src="{{ static_url("urlJump.html") }}?img=${escape(profChartUrl)}" style="width:315px;height:270px;"></iframe>`;
}

/**
 * Generates a horizontal bar chart from the selected range
 * @param {Object} spreadsheet - The spreadsheet control object
 * @param {Object|null} range - Range object with left, right, top, bottom properties
 * @param {HTMLElement} gview - The graph view container element
 * @param {string} gtype - The graph type identifier
 * @param {boolean} helpflag - Whether to display help text instead of chart
 * @param {boolean} [isResize] - Whether this is a resize operation (unused)
 * @returns {void}
 */
function GraphHorizontalBar(spreadsheet, range, gview, gtype, helpflag) {
    const values = [];
    const labels = [];
    
    if (helpflag || !range) {
        const str = `<input type="button" value="Hide Help" onclick="DoGraph(false,false);"><br><br>
            This is the help text for graph type: ${SocialCalc.GraphTypesInfo[gtype].display}.<br><br>
            The <b>Graph</b> tab displays a very simple bar graph representation of the cells currently selected as a range to graph 
            (either in a single row across or column down). 
            If the range is a single row or column, and if the row above (or column to the left) has values, those values are used as labels. 
            Otherwise the cell coordinates are used (e.g., B5). 
            This is a very early, minimal implementation for demonstration purposes. 
            <br><br><input type="button" value="Hide Help" onclick="DoGraph(false,false);">`;
        
        gview.innerHTML = str;
        return;
    }
    
    // Determine if we're working with a column (byrow=true) or row (byrow=false)
    const byrow = range.left === range.right;
    const nitems = byrow ? range.bottom - range.top + 1 : range.right - range.left + 1;
    
    const maxheight = (spreadsheet.height - spreadsheet.nonviewheight) - 50;
    const totalwidth = spreadsheet.width - 30;
    
    let minval = null;
    let maxval = null;
    
    // Extract values and labels from the range
    for (let i = 0; i < nitems; i++) {
        const cr = byrow ? 
            `${SocialCalc.rcColname(range.left)}${i + range.top}` : 
            `${SocialCalc.rcColname(i + range.left)}${range.top}`;
        
        const cr1 = byrow ? 
            `${SocialCalc.rcColname(range.left - 1 || 1)}${i + range.top}` : 
            `${SocialCalc.rcColname(i + range.left)}${range.top - 1 || 1}`;
        
        const cell = spreadsheet.sheet.GetAssuredCell(cr);
        
        if (cell.valuetype.charAt(0) === "n") {
            const val = cell.datavalue - 0;
            
            if (maxval === null || maxval < val) maxval = val;
            if (minval === null || minval > val) minval = val;
            
            values.push(val);
            
            const labelCell = spreadsheet.sheet.GetAssuredCell(cr1);
            
            if ((range.right === range.left || range.top === range.bottom) && 
                (labelCell.valuetype.charAt(0) === "t" || labelCell.valuetype.charAt(0) === "n")) {
                labels.push(labelCell.datavalue + "");
            } else {
                labels.push(val + "");
            }
        }
    }
    
    // Ensure zero line is visible
    if (maxval < 0) maxval = 0;
    if (minval > 0) minval = 0;
    
    const str = '<table><tr><td><canvas id="myBarCanvas" width="500px" height="400px" style="border:1px solid black;"></canvas></td><td><span id="googleBarChart"></span></td></tr></table>';
    gview.innerHTML = str;
    
    const profChartVals = [];
    const profChartLabels = [];
    
    const canv = document.getElementById("myBarCanvas");
    const ctx = canv.getContext('2d');
    
    // Note: mozTextStyle is deprecated but kept for compatibility
    ctx.mozTextStyle = "10pt bold Arial";
    
    const canvasMaxHeight = canv.height - 60;
    const canvasTotalWidth = canv.width;
    const colors = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];
    
    /**
     * Generates a random hex color
     * @returns {string} Random 6-character hex color
     */
    const generateRandomColor = () => {
        return Array.from({ length: 6 }, () => colors[Math.round(Math.random() * 14)]).join('');
    };
    
    let barColor = generateRandomColor();
    ctx.fillStyle = `#${barColor}`;
    const colorList = [barColor];
    
    const eachwidth = Math.floor(canvasMaxHeight / (values.length || 1)) - 4 || 1;
    let zeroLine = canvasTotalWidth * (maxval / (maxval - minval)) - 5;
    zeroLine = canv.width - zeroLine + 40;
    
    // Draw zero line
    ctx.lineWidth = 5;
    ctx.moveTo(zeroLine, 0);
    ctx.lineTo(zeroLine, canv.height);
    ctx.stroke();
    
    const yScale = canvasTotalWidth / (maxval - minval) * 4.4 / 5;
    
    // Draw bars
    for (let i = 0; i < values.length; i++) {
        ctx.fillRect(zeroLine + yScale * values[i], i * eachwidth + 30, -1 * yScale * values[i], eachwidth);
        profChartVals.push(Math.floor((values[i] - minval) * yScale / 4.4));
        profChartLabels.push(labels[i]);
        
        barColor = generateRandomColor();
        ctx.fillStyle = `#${barColor}`;
        colorList.push(barColor);
    }
    
    // Draw labels
    ctx.strokeStyle = "#000000";
    ctx.fillStyle = "#000000";
    
    if (values[0] > 0) {
        ctx.translate(zeroLine - 22, 45);
    } else {
        ctx.translate(zeroLine + 15, 45);
    }
    
    ctx.mozDrawText(labels[0]);
    
    for (let i = 1; i < values.length; i++) {
        if ((values[i] > 0) && (values[i - 1] < 0)) {
            ctx.translate(-37, eachwidth);
        } else if ((values[i] < 0) && (values[i - 1] > 0)) {
            ctx.translate(37, eachwidth);
        } else {
            ctx.translate(0, eachwidth);
        }
        ctx.mozDrawText(labels[i]);
    }
    
    // Generate Google Charts API URL
    const gChart = document.getElementById("googleBarChart");
    const googleZeroLine = (-1 * minval) * yScale / canv.width;
    const profChartUrl = `chs=300x250&cht=bhs&chd=t:${profChartVals.join(",")}&chxt=x,y&chxl=1:|${profChartLabels.reverse().join("|")}|&chxr=0,${minval},${maxval}&chp=${googleZeroLine}&chbh=a&chm=r,000000,0,${googleZeroLine},${googleZeroLine + 0.005}&chco=${colorList.join("|")}`;
    
    gChart.innerHTML = `<iframe src="urlJump.html?img=${escape(profChartUrl)}" style="width:315px;height:270px;"></iframe>`;
}

/**
 * Generates a pie chart from the selected range
 * @param {Object} spreadsheet - The spreadsheet control object
 * @param {Object} range - Range object with left, right, top, bottom properties
 * @param {HTMLElement} gview - The graph view container element
 * @param {string} gtype - The graph type identifier
 * @param {boolean} helpflag - Whether to display help text instead of chart
 * @returns {void}
 */
function MakePieChart(spreadsheet, range, gview, gtype, helpflag) {
    const values = [];
    const labels = [];
    let total = 0;
    
    // Determine if we're working with a column (byrow=true) or row (byrow=false)
    const byrow = range.left === range.right;
    const nitems = byrow ? range.bottom - range.top + 1 : range.right - range.left + 1;
    
    // Collect values and calculate total for distribution over 2π radians
    for (let i = 0; i < nitems; i++) {
        const cr = byrow ? 
            `${SocialCalc.rcColname(range.left)}${i + range.top}` : 
            `${SocialCalc.rcColname(i + range.left)}${range.top}`;
        
        const cr1 = byrow ? 
            `${SocialCalc.rcColname(range.left - 1 || 1)}${i + range.top}` : 
            `${SocialCalc.rcColname(i + range.left)}${range.top - 1 || 1}`;
        
        const cell = spreadsheet.sheet.GetAssuredCell(cr);
        
        if (cell.valuetype.charAt(0) === "n") {
            const val = cell.datavalue - 0;
            total += val;
            values.push(val);
            
            const labelCell = spreadsheet.sheet.GetAssuredCell(cr1);
            
            if ((range.right === range.left || range.top === range.bottom) && 
                (labelCell.valuetype.charAt(0) === "t" || labelCell.valuetype.charAt(0) === "n")) {
                labels.push(labelCell.datavalue + "");
            } else {
                labels.push(val + "");
            }
        }
    }
    
    const str = '<table><tr><td><img id="canvImg" style="border:1px solid black;" src=""/><canvas id="myCanvas" style="display:none;" width="500px" height="400px"></canvas></td><td><span id="googleChart"></span></td></tr></table>';
    gview.innerHTML = str;
    
    let profChartUrl = "";
    let profChartLabels = "";
    
    const canv = document.getElementById("myCanvas");
    const ctx = canv.getContext('2d');
    
    // Note: mozTextStyle is deprecated but kept for compatibility
    ctx.mozTextStyle = "10pt Arial";
    
    const centerX = canv.width / 2;
    const centerY = canv.height / 2;
    const rad = centerY - 50;
    const textRad = rad * 1.1;
    let lastStart = 0;
    
    // Color palette for random colors (0-f hex values)
    const colors = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];
    
    /**
     * Generates a random 3-digit hex color
     * @returns {string} Random hex color in format #rgb
     */
    const generateRandomColor = () => {
        return `#${colors[Math.round(Math.random() * 14)]}${colors[Math.round(Math.random() * 14)]}${colors[Math.round(Math.random() * 14)]}`;
    };
    
    // Draw pie chart slices
    for (let i = 0; i < values.length; i++) {
        // Prepare to draw a slice
        ctx.beginPath();
        ctx.moveTo(centerX, centerY);
        
        // Choose a random color (note: color changes on each redraw)
        const arcColor = generateRandomColor();
        ctx.fillStyle = arcColor;
        
        // Set the size of this arc piece in radians
        const arcRads = 2 * Math.PI * (values[i] / total);
        profChartUrl += `,${values[i]}`;
        
        // Draw arc
        ctx.arc(centerX, centerY, rad, lastStart, lastStart + arcRads, false);
        ctx.closePath();
        ctx.fill();
        
        // Draw label with percentage
        ctx.fillStyle = "black";
        const centralRad = lastStart + 0.5 * arcRads;
        
        // leftBias gives text more room if it's on the left part of the circle
        let leftBias = 0;
        if ((centralRad > 1.5) && (centralRad < 4.6)) {
            leftBias = 55;
        }
        
        ctx.translate(
            centerX + Math.cos(centralRad) * textRad - leftBias, 
            centerY + Math.sin(centralRad) * textRad
        );
        
        // Note: mozDrawText is deprecated but needed for compatibility
        ctx.mozDrawText(`${labels[i]} (${Math.round(values[i] / total * 100)}%)`);
        
        // Reset translation to continue drawing
        ctx.translate(
            -1 * centerX - Math.cos(centralRad) * textRad + leftBias,
            -1 * centerY - Math.sin(centralRad) * textRad
        );
        
        ctx.fillRect(1, 1, 1, 1);
        ctx.closePath();
        profChartLabels += `|${labels[i]}`;
        
        // Prepare for next arc
        lastStart += arcRads;
    }
    
    // Replace HTML canvas with its PNG image
    const realCanv = document.getElementById("canvImg");
    realCanv.src = canv.toDataURL();
    
    // Generate Google Charts API image
    const gChart = document.getElementById("googleChart");
    profChartUrl = `chs=300x145&cht=p&chd=t:${profChartUrl.substring(1)}&chl=${profChartLabels.substring(1)}`;
    gChart.innerHTML = `<iframe src="urlJump.html?img=${escape(profChartUrl)}" style="width:315px;height:270px;"></iframe>`;
}

/**
 * Generates a line chart from the selected range
 * @param {Object} spreadsheet - The spreadsheet control object
 * @param {Object} range - Range object with left, right, top, bottom properties
 * @param {HTMLElement} gview - The graph view container element
 * @param {string} gtype - The graph type identifier
 * @param {boolean} helpflag - Whether to display help text instead of chart
 * @param {boolean} isResize - Whether this is a resize operation affecting min/max values
 * @returns {void}
 */
function MakeLineChart(spreadsheet, range, gview, gtype, helpflag, isResize) {
    const values = [];
    const labels = [];
    let total = 0;
    
    // Color and shape palettes for chart styling
    const colors = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];
    const shapes = ["s", "o", "c"];
    
    // Determine if we're working with a column (byrow=true) or row (byrow=false)
    const byrow = range.left === range.right;
    const nitems = byrow ? range.bottom - range.top + 1 : range.right - range.left + 1;
    
    // Initialize min/max values - use user-set values if this is a resize operation
    let minX = null, maxX = null, minval = null, maxval = null;
    
    if (isResize) {
        try { 
            minX = 1 * document.getElementById("SocialCalc-graphMinX").value; 
        } catch (e) { 
            minX = null; 
        }
        
        try { 
            maxX = 1 * document.getElementById("SocialCalc-graphMaxX").value; 
        } catch (e) { 
            maxX = null; 
        }
        
        try { 
            minval = 1 * document.getElementById("SocialCalc-graphMinY").value; 
        } catch (e) { 
            minval = null; 
        }
        
        try { 
            maxval = 1 * document.getElementById("SocialCalc-graphMaxY").value; 
        } catch (e) { 
            maxval = null; 
        }
    }
    
    let evenlySpaced = false;
    
    // Collect values and labels from the range
    for (let i = 0; i < nitems; i++) {
        const cr = byrow ? 
            `${SocialCalc.rcColname(range.left)}${i + range.top}` : 
            `${SocialCalc.rcColname(i + range.left)}${range.top}`;
        
        const cr1 = byrow ? 
            `${SocialCalc.rcColname(range.left - 1 || 1)}${i + range.top}` : 
            `${SocialCalc.rcColname(i + range.left)}${range.top - 1 || 1}`;
        
        const cell = spreadsheet.sheet.GetAssuredCell(cr);
        
        if (cell.valuetype.charAt(0) === "n") {
            const val = cell.datavalue - 0;
            
            // Update Y-axis min/max if not in resize mode
            if ((maxval === null || maxval < val) && !isResize) {
                maxval = val;
            }
            if ((minval === null || minval > val) && !isResize) {
                minval = val;
            }
            
            values.push(val);
            
            const labelCell = spreadsheet.sheet.GetAssuredCell(cr1);
            
            if ((range.right === range.left || range.top === range.bottom) && 
                (labelCell.valuetype.charAt(0) === "t" || labelCell.valuetype.charAt(0) === "n")) {
                
                labels.push(labelCell.datavalue + "");
                
                // Update X-axis min/max if not in resize mode
                if ((maxX === null || maxX < labelCell.datavalue) && !isResize) {
                    maxX = labelCell.datavalue;
                }
                if ((minX === null || minX > labelCell.datavalue) && !isResize) {
                    minX = labelCell.datavalue;
                }
            } else {
                labels.push(cr);
                evenlySpaced = true;
            }
        }
    }
    
    // Create evenly-spaced X values if none were provided
    if (evenlySpaced) {
        for (let i = 0; i < values.length; i++) {
            labels[i] = i;
        }
        
        if (!isResize) {
            minX = 0;
            maxX = values.length - 1;
        }
    }
    
    const str = '<canvas id="myLineCanvas" style="border:1px solid black;" width="500px" height="400px"></canvas><span id="googleLineChart"></span>';
    gview.innerHTML = str;
    
    // Set min/max values in UI controls if not in resize mode
    if (!isResize) {
        document.getElementById("SocialCalc-graphMinX").value = minX;
        spreadsheet.graphMinX = minX;
        document.getElementById("SocialCalc-graphMaxX").value = maxX;
        spreadsheet.graphMaxX = maxX;
        document.getElementById("SocialCalc-graphMinY").value = minval;
        spreadsheet.graphMinY = minval;
        document.getElementById("SocialCalc-graphMaxY").value = maxval;
        spreadsheet.graphMaxY = maxval;
    }
    
    // Note: Function appears to be incomplete - likely continues in next chunk
    // Continuation of MakeLineChart function
    const canv = document.getElementById("myLineCanvas");
    const ctx = canv.getContext('2d');
    
    // Calculate scaling factors for chart dimensions
    const scaleFactorX = (canv.width - 40) / (maxX - minX);
    const scaleFactorY = (canv.height - 40) / (maxval - minval);
    
    let lastX = scaleFactorX * (labels[0] - minX) + 20;
    let lastY = scaleFactorY * (values[0] - minval) + 20;
    
    const profChart = [
        Math.floor(lastX / canv.width * 100),
        Math.floor(lastY / canv.height * 100)
    ];
    
    const topY = canv.height;
    
    /**
     * Generates a random 6-digit hex color
     * @returns {string} Random hex color in format #rrggbb
     */
    const generateRandomColor = () => {
        return `#${Array.from({ length: 6 }, () => colors[Math.round(Math.random() * 14)]).join('')}`;
    };
    
    let drawColor = generateRandomColor();
    const colorArray = [drawColor.replace("#", "")];
    
    ctx.strokeStyle = drawColor;
    ctx.fillStyle = drawColor;
    ctx.fillRect(lastX - 3, topY - lastY - 3, 6, 6);
    ctx.beginPath();
    
    // Draw line chart connecting points
    for (let i = 1; i < values.length; i++) {
        // Determine if next X is part of the same line (greater than the last X value)
        if ((labels[i] * 1) > (labels[i - 1] * 1)) {
            // Draw line to the next point
            ctx.moveTo(lastX, topY - lastY);
            ctx.lineTo(
                (scaleFactorX * (labels[i] - minX)) + 20, 
                topY - (scaleFactorY * (values[i] - minval) + 20)
            );
            ctx.stroke();
        } else {
            // Start a new line with new color
            drawColor = generateRandomColor();
            ctx.strokeStyle = drawColor;
            ctx.fillStyle = drawColor;
            colorArray.push(drawColor.replace("#", ""));
            ctx.beginPath();
        }
        
        // Calculate canvas coordinates for next point
        lastX = scaleFactorX * (labels[i] - minX) + 20;
        lastY = scaleFactorY * (values[i] - minval) + 20;
        
        // Draw different shapes based on line index
        const shapeIndex = (colorArray.length - 1) % 3;
        
        if (shapeIndex === 0) {
            // Square
            ctx.fillRect(lastX - 3, topY - lastY - 3, 6, 6);
        } else if (shapeIndex === 1) {
            // Circle
            ctx.beginPath();
            ctx.arc(lastX, topY - lastY, 3, 0, Math.PI * 2, false);
            ctx.fill();
        } else {
            // Plus sign
            ctx.fillRect(lastX, topY - lastY - 3, 2, 8);
            ctx.fillRect(lastX - 3, topY - lastY, 8, 2);
        }
        
        // Update Google chart data
        if ((labels[i] * 1) > (labels[i - 1] * 1)) {
            // Add a point to the current line
            profChart[profChart.length - 2] += `,${Math.floor(lastX / canv.width * 100)}`;
            profChart[profChart.length - 1] += `,${Math.floor(lastY / canv.height * 100)}`;
        } else {
            // Add a new line
            const newIndex = profChart.length;
            profChart[newIndex] = Math.floor(lastX / canv.width * 100);
            profChart[newIndex + 1] = Math.floor(lastY / canv.height * 100);
        }
    }
    
    ctx.stroke();
    
    // Prepare color markings for Google Charts API
    let colorMarkings = `&chco=${colorArray.join(",")}&chm=`;
    
    for (let i = 0; i < colorArray.length; i++) {
        const shapeIndex = i % 3;
        
        if (shapeIndex === 0) {
            // Square
            colorArray[i] = `s,${colorArray[i]},${i},-1,6`;
        } else if (shapeIndex === 1) {
            // Circle
            colorArray[i] = `o,${colorArray[i]},${i},-1,6`;
        } else {
            // Plus sign
            colorArray[i] = `c,${colorArray[i]},${i},-1,10`;
        }
    }
    
    colorMarkings += colorArray.join("|");
    
    // Draw X=0 axis if data spans zero
    if (minval <= 0 && maxval >= 0) {
        ctx.beginPath();
        ctx.strokeStyle = "#000000";
        ctx.moveTo(0, canv.height - (scaleFactorY * -1 * minval + 20));
        ctx.lineTo(canv.width, canv.height - (scaleFactorY * -1 * minval + 20));
        ctx.stroke();
        
        const graphPlace = 1 - ((canv.height - (scaleFactorY * -1 * minval + 20)) / canv.height);
        colorMarkings += `|r,000000,0,${graphPlace},${graphPlace + 0.005}`;
    }
    
    // Draw Y=0 axis if data spans zero
    if (minX <= 0 && maxX >= 0) {
        ctx.beginPath();
        ctx.strokeStyle = "#000000";
        ctx.moveTo(scaleFactorX * -1 * minX + 20, 0);
        ctx.lineTo(scaleFactorX * -1 * minX + 20, canv.height);
        ctx.stroke();
        
        const graphPlace = (scaleFactorX * -1 * minX + 20) / canv.width;
        colorMarkings += `|R,000000,0,${graphPlace},${graphPlace + 0.005}`;
    }
    
    // Generate Google Charts API URL
    const gChart = document.getElementById("googleLineChart");
    
    // Add margin to sides of Google chart
    minX -= (maxX - minX) / 23;
    maxX += (maxX - minX) / 23;
    minval -= (maxval - minval) / 18;
    maxval += (maxval - minval) / 18;
    
    const profChartUrl = `chs=300x250${colorMarkings}&cht=lxy&chxt=x,y&chxr=0,${minX},${maxX}|1,${minval},${maxval}&chd=t:${profChart.join("|")}`;
    gChart.innerHTML = `<iframe src="urlJump.html?img=${escape(profChartUrl)}" style="width:315px;height:270px;"></iframe>`;
}

/**
 * Generates a scatter chart from the selected range with optional dot sizing
 * @param {Object} spreadsheet - The spreadsheet control object
 * @param {Object} range - Range object with left, right, top, bottom properties
 * @param {HTMLElement} gview - The graph view container element
 * @param {string} gtype - The graph type identifier
 * @param {boolean} helpflag - Whether to display help text instead of chart
 * @param {boolean} isResize - Whether this is a resize operation affecting min/max values
 * @returns {void}
 */
function MakeScatterChart(spreadsheet, range, gview, gtype, helpflag, isResize) {
    const values = [];
    const labels = [];
    let total = 0;
    
    // Color palette for chart styling
    const colors = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f"];
    
    // Determine if we're working with a column (byrow=true) or row (byrow=false)
    const byrow = range.left === range.right;
    const nitems = byrow ? range.bottom - range.top + 1 : range.right - range.left + 1;
    
    // Initialize min/max values - use user-set values if this is a resize operation
    let minX = null, maxX = null, minval = null, maxval = null;
    
    if (isResize) {
        try { 
            minX = 1 * document.getElementById("SocialCalc-graphMinX").value; 
        } catch (e) { 
            minX = null; 
        }
        
        try { 
            maxX = 1 * document.getElementById("SocialCalc-graphMaxX").value; 
        } catch (e) { 
            maxX = null; 
        }
        
        try { 
            minval = 1 * document.getElementById("SocialCalc-graphMinY").value; 
        } catch (e) { 
            minval = null; 
        }
        
        try { 
            maxval = 1 * document.getElementById("SocialCalc-graphMaxY").value; 
        } catch (e) { 
            maxval = null; 
        }
    }
    
    let evenlySpaced = false;
    const dotSizes = [];
    
    // Collect values, labels, and dot sizes from the range
    for (let i = 0; i < nitems; i++) {
        const cr = byrow ? 
            `${SocialCalc.rcColname(range.left)}${i + range.top}` : 
            `${SocialCalc.rcColname(i + range.left)}${range.top}`;
        
        const cr1 = byrow ? 
            `${SocialCalc.rcColname(range.left - 1 || 1)}${i + range.top}` : 
            `${SocialCalc.rcColname(i + range.left)}${range.top - 1 || 1}`;
        
        const cr2 = byrow ? 
            `${SocialCalc.rcColname(range.left + 1 || 2)}${i + range.top}` : 
            `${SocialCalc.rcColname(i + range.left)}${range.top + 1 || 2}`;
        
        const cell = spreadsheet.sheet.GetAssuredCell(cr);
        
        if (cell.valuetype.charAt(0) === "n") {
            const val = cell.datavalue - 0;
            
            // Update Y-axis min/max if not in resize mode
            if ((maxval === null || maxval < val) && !isResize) {
                maxval = val;
            }
            if ((minval === null || minval > val) && !isResize) {
                minval = val;
            }
            
            values.push(val);
            
            const labelCell = spreadsheet.sheet.GetAssuredCell(cr1);
            
            if ((range.right === range.left || range.top === range.bottom) && 
                (labelCell.valuetype.charAt(0) === "t" || labelCell.valuetype.charAt(0) === "n")) {
                
                labels.push(labelCell.datavalue + "");
                
                // Update X-axis min/max if not in resize mode
                if ((maxX === null || maxX < labelCell.datavalue) && !isResize) {
                    maxX = labelCell.datavalue;
                }
                if ((minX === null || minX > labelCell.datavalue) && !isResize) {
                    minX = labelCell.datavalue;
                }
            } else {
                labels.push(cr);
                evenlySpaced = true;
            }
            
            // Get dot size from adjacent column/row
            const sizeCell = spreadsheet.sheet.GetAssuredCell(cr2);
            
            if ((range.right === range.left || range.top === range.bottom) && 
                (sizeCell.valuetype.charAt(0) === "t" || sizeCell.valuetype.charAt(0) === "n")) {
                dotSizes.push(sizeCell.datavalue + "");
            } else {
                dotSizes.push("5");
            }
        }
    }
    
    // Create evenly-spaced X values if none were provided
    if (evenlySpaced) {
        for (let i = 0; i < values.length; i++) {
            labels[i] = i;
        }
        
        if (!isResize) {
            minX = 0;
            maxX = values.length - 1;
        }
    }
    
    let str = '<canvas id="myScatterCanvas" style="border:1px solid black;" width="500px" height="400px"></canvas><span id="googleScatterChart"></span>';
    str += '<div id="scatterChartScales"><input type="button" id="autoScaleButton" value="Reset" onclick=""/>X-min:<input id="minPlotX" onchange="" size=5/>X-max:<input id="maxPlotX" onchange="" size=5/>Y-min:<input id="minPlotY" onchange="" size=5/>Y-max:<input id="maxPlotY" onchange="" size=5/></div>';
    gview.innerHTML = str;
    
    // Set min/max values in UI controls if not in resize mode
    if (!isResize) {
        document.getElementById("SocialCalc-graphMinX").value = minX;
        spreadsheet.graphMinX = minX;
        document.getElementById("SocialCalc-graphMaxX").value = maxX;
        spreadsheet.graphMaxX = maxX;
        document.getElementById("SocialCalc-graphMinY").value = minval;
        spreadsheet.graphMinY = minval;
        document.getElementById("SocialCalc-graphMaxY").value = maxval;
        spreadsheet.graphMaxY = maxval;
    }
    
    const canv = document.getElementById("myScatterCanvas");
    const ctx = canv.getContext('2d');
    
    // Calculate scaling factors for chart dimensions
    const scaleFactorX = (canv.width - 40) / (maxX - minX);
    const scaleFactorY = (canv.height - 40) / (maxval - minval);
    
    let lastX = scaleFactorX * (labels[0] - minX) + 20;
    let lastY = scaleFactorY * (values[0] - minval) + 20;
    
    const profChart = [
        Math.floor(lastX / canv.width * 100),
        Math.floor(lastY / canv.height * 100), 
        dotSizes[0] * 10
    ];
    
    const topY = canv.height;
    
    /**
     * Generates a random 6-digit hex color
     * @returns {string} Random hex color in format #rrggbb
     */
    const generateRandomColor = () => {
        return `#${Array.from({ length: 6 }, () => colors[Math.round(Math.random() * 14)]).join('')}`;
    };
    
    const drawColor = generateRandomColor();
    ctx.fillStyle = drawColor;
    ctx.beginPath();
    ctx.arc(lastX, topY - lastY, dotSizes[0], 0, 2 * Math.PI, false);
    ctx.fill();
    
    // Draw scatter points
    for (let i = 1; i < values.length; i++) {
        // Draw next point
        ctx.moveTo(lastX, topY - lastY);
        lastX = scaleFactorX * (labels[i] - minX) + 20;
        lastY = scaleFactorY * (values[i] - minval) + 20;
        
        ctx.beginPath();
        ctx.arc(lastX, topY - lastY, dotSizes[i], 0, 2 * Math.PI, false);
        ctx.fill();
        
        // Update Google chart data
        profChart[profChart.length - 3] += `,${Math.floor(lastX / canv.width * 100)}`;
        profChart[profChart.length - 2] += `,${Math.floor(lastY / canv.height * 100)}`;
        profChart[profChart.length - 1] += `,${dotSizes[i] * 10}`;
    }
    
    // Prepare color markings for Google Charts API
    let colorMarkings = `&chm=o,${drawColor.replace("#", "")},0,-1,10`;
    
    // Draw X=0 axis if data spans zero
    if (minval <= 0 && maxval >= 0) {
        ctx.beginPath();
        ctx.strokeStyle = "#000000";
        ctx.moveTo(0, canv.height - (scaleFactorY * -1 * minval + 20));
        ctx.lineTo(canv.width, canv.height - (scaleFactorY * -1 * minval + 20));
        ctx.stroke();
        
        const graphPlace = 1 - ((canv.height - (scaleFactorY * -1 * minval + 20)) / canv.height);
        colorMarkings += `|r,000000,0,${graphPlace},${graphPlace + 0.005}`;
    }
    
    // Draw Y=0 axis if data spans zero
    if (minX <= 0 && maxX >= 0) {
        ctx.beginPath();
        ctx.strokeStyle = "#000000";
        ctx.moveTo(scaleFactorX * -1 * minX + 20, 0);
        ctx.lineTo(scaleFactorX * -1 * minX + 20, canv.height);
        ctx.stroke();
        
        const graphPlace = (scaleFactorX * -1 * minX + 20) / canv.width;
        colorMarkings += `|R,000000,0,${graphPlace},${graphPlace + 0.005}`;
    }
    
    // Generate Google Charts API URL
    const gChart = document.getElementById("googleScatterChart");
    
    // Add margin to sides of Google chart
    minX -= (maxX - minX) / 23;
    maxX += (maxX - minX) / 23;
    minval -= (maxval - minval) / 18;
    maxval += (maxval - minval) / 18;
    
    const profChartUrl = `chs=300x250${colorMarkings}&cht=s&chxt=x,y&chxr=0,${minX},${maxX}|1,${minval},${maxval}&chd=t:${profChart.join("|")}`;
    gChart.innerHTML = `<iframe src="urlJump.html?img=${escape(profChartUrl)}" style="width:315px;height:270px;"></iframe>`;
}

