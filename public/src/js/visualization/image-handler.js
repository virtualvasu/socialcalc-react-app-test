/**
 * Image Handler Module for SocialCalc
 * Handles image embedding functionality including file uploads and URL images
 * Provides utilities for cell parsing and image insertion into spreadsheet cells
 */

// Initialize SocialCalc namespace if not exists
const SocialCalc = window.SocialCalc || (() => {
	alert("Main SocialCalc code module needed");
	return {};
})();

if (!SocialCalc.TableEditor) {
	alert("SocialCalc TableEditor code module needed");
}

/**
 * Parses cell reference string and extracts column and row details
 * @param {string} str - Cell reference string (e.g., 'A1', 'B2', 'AA10')
 * @returns {Array} Array containing [columnNumber, rowNumber, columnLetter, rowString]
 * @example
 * SocialCalc.getCellDetails('A1') // returns [1, 1, 'A', '1']
 * SocialCalc.getCellDetails('B2') // returns [2, 2, 'B', '2']
 */
SocialCalc.getCellDetails = function (str) {
	let col = "";
	let i = -1;
	let char;
	let code = 0;
	while(true) {
		char = str.charCodeAt(++i);
		if(char >=65 && char <=90) {
			code = 26 * code + (char - 64);
			col = col + str.charAt(i);
			}
		else {
			break;
		}
	}
	const row = str.substring(i);
	const arr = [];
	arr.push(code);
	arr.push(parseInt(row));
	arr.push(col);
	arr.push(row);
	return arr;
};

/**
 * Calculates range dimensions from cell range string
 * @param {string} str - Range string in format 'A1:B2'
 * @returns {Array} Array containing [columnCount, rowCount]
 * @example
 * SocialCalc.getRange('A1:C3') // returns [3, 3]
 * SocialCalc.getRange('B2:D5') // returns [3, 4]
 */
SocialCalc.getRange = function (str) {
	const strarr = str.split(":");
	const arr = [];
	arr.push(SocialCalc.getCellDetails(strarr[0]));
	arr.push(SocialCalc.getCellDetails(strarr[1]));
	const range = [];
	range.push(arr[1][0] - arr[0][0] + 1);
	range.push(arr[1][1] - arr[0][1] + 1);
	return range;
};

/**
 * Image handling constructor
 * Initializes image dimensions properties
 * @constructor
 */
SocialCalc.Images = function() {
	this.height = 0;
	this.width = 0;
};

/**
 * Shows the image embedding form dialog
 * Displays the embed image interface and focuses on the range input
 */
SocialCalc.Images.Insert = function() {
	const embedImageEl = document.getElementById("EmbedImage");
	const rangeInputEl = document.getElementById('image-embed-range');
	
	if (embedImageEl) {
		embedImageEl.style.display = "inline";
	}
	if (rangeInputEl) {
		SocialCalc.CmdGotFocus(rangeInputEl);
	}
};

/**
 * Shows the appropriate image form (local file or URL)
 * @param {string} value - Type of image form to show ('local' or 'url')
 */
SocialCalc.Images.showImgForm = function(value) {
	const formId = `${value}ImgForm`;
	const localForm = document.getElementById("localImgForm");
	const urlForm = document.getElementById("urlImgForm");
	const targetForm = document.getElementById(formId);
	const submitBtn = document.getElementById("image-embed-submit-button");
	const cancelBtn = document.getElementById("image-embed-cancel-button");
	
	// Hide all forms first
	if (localForm) localForm.style.display = "none";
	if (urlForm) urlForm.style.display = "none";
	
	// Show target form
	if (targetForm) {
		targetForm.style.display = "inline";
	}
	
	// Focus URL input if URL form is selected
	if (value === 'url') {
		const txtBox = document.getElementById("imgurl");
		if (txtBox) {
			SocialCalc.CmdGotFocus(txtBox);
		}
	}
	
	// Show control buttons
	if (submitBtn) submitBtn.style.display = "inline";
	if (cancelBtn) cancelBtn.style.display = "inline";
};

/**
 * Handles file selection for local image uploads
 * Processes selected files and displays image preview
 * @param {Event} evt - File input change event
 */
SocialCalc.Images.handleFileSelect = function(evt) {
	const files = evt.target.files;
	
	if (!files || files.length === 0) {
		console.warn('No files selected');
		return;
	}
	
	for (const file of files) {
		console.log('Processing file:', file.name);
		
		// Validate file type
		if (!file.type.match('image.*')) {
			console.warn('Skipping non-image file:', file.name);
			continue;
		}
		
		const reader = new FileReader();
		
		// Handle file read completion
		reader.onload = ((theFile) => {
			return (e) => {
				const fileImageHolder = document.getElementById("file-image-holder");
				const fileTextHolder = document.getElementById("file-text-holder");
				
				if (!fileImageHolder || !fileTextHolder) {
					console.error('Required DOM elements not found');
					return;
				}
				
				// Update image display
				fileImageHolder.src = e.target.result;
				fileImageHolder.style.display = "inline";
				fileTextHolder.style.display = "none";
				
				// Set image attributes
				fileImageHolder.className = "thumb";
				fileImageHolder.setAttribute("title", theFile.name || 'Uploaded image');
			};
		})(file);
		
		// Handle read errors
		reader.onerror = () => {
			console.error('Error reading file:', file.name);
		};
		
		reader.readAsDataURL(file);
	}
};
	
/**
 * Loads and displays image from URL
 * Validates URL and shows image preview
 */
SocialCalc.Images.getUrlImage = function () {
	const urlInput = document.getElementById("imgurl");
	const urlImageHolder = document.getElementById("url-image-holder");
	const urlTextHolder = document.getElementById("url-text-holder");
	
	if (!urlInput || !urlImageHolder || !urlTextHolder) {
		console.error('Required DOM elements not found');
		return;
	}
	
	const imageUrl = urlInput.value.trim();
	
	if (!imageUrl) {
		console.warn('No URL provided');
		return;
	}
	
	// Basic URL validation
	try {
		new URL(imageUrl);
	} catch (e) {
		console.error('Invalid URL:', imageUrl);
		return;
	}
	
	urlImageHolder.src = imageUrl;
	urlImageHolder.style.display = "inline";
	urlTextHolder.style.display = "none";
};

/**
 * Embeds the selected image into the specified spreadsheet range
 * Merges cells and inserts HTML img tag with calculated dimensions
 */
SocialCalc.Images.Embed = function () {
	const ifi = document.getElementById("img-file-inp");
	const fih = document.getElementById("file-image-holder");
	const uih = document.getElementById("url-image-holder");
	const ier = document.getElementById("image-embed-range");
	
	if (!ifi || !fih || !uih || !ier) {
		console.error('Required DOM elements not found');
		return;
	}
	
	if (!ier.value.trim()) {
		console.error('No range specified');
		return;
	}
	
	const imgsource = uih.width !== 0 ? uih : fih; 
	const width = imgsource.width;
	const height = imgsource.height;
	
	if (!imgsource.src) {
		console.error('No image source available');
		return;
	}
	
	try {
		const colrange = width / SocialCalc.Constants.defaultColWidth;
		const rowrange = height / 20;
		
		const arr = SocialCalc.getRange(ier.value);
		const calculatedHeight = 20 * arr[1];
		const imgHtml = `<img src="${imgsource.src}" style="height:${calculatedHeight}px;"/>`;
		
		const spl = ier.value.split(":")[0];
		const cmdk = `merge ${ier.value}\nset ${spl} textvalueformat text-html\nset ${spl} text t ${imgHtml}`;
		
		console.log('Executing command:', cmdk);
		SocialCalc.ScheduleSheetCommands(
			workbook.sheetArr[SocialCalc.GetCurrentWorkBookControl().currentSheetButton.id].sheet,
			cmdk,
			true,
			true
		);
		
		// Cleanup form state
		SocialCalc.Images._cleanupForm(ifi, fih, uih, ier);
		
	} catch (error) {
		console.error('Error embedding image:', error);
	}
};

/**
 * Internal cleanup function for form elements
 * @private
 * @param {HTMLElement} ifi - File input element
 * @param {HTMLElement} fih - File image holder element
 * @param {HTMLElement} uih - URL image holder element  
 * @param {HTMLElement} ier - Image embed range element
 */
SocialCalc.Images._cleanupForm = function(ifi, fih, uih, ier) {
	// Reset input values
	if (ifi) ifi.value = "";
	if (ier) ier.value = "";
	
	// Reset image holders
	if (fih) {
		fih.src = "";
		fih.style.display = "none";
	}
	if (uih) {
		uih.src = "";
		uih.style.display = "none";
	}
	
	// Reset form elements
	const elementsToHide = [
		"imgurl", "localImgForm", "urlImgForm", 
		"embed-button", "EmbedImage"
	];
	
	elementsToHide.forEach(id => {
		const element = document.getElementById(id);
		if (element) {
			if (id === "imgurl") {
				element.value = "";
			} else {
				element.style.display = "none";
			}
		}
	});
};

/**
 * Hides the image embedding form
 * @param {string} value - Form type to hide (unused parameter)
 */
SocialCalc.Images.hideImgForm = function(value) {
	const embedImageEl = document.getElementById("EmbedImage");
	if (embedImageEl) {
		embedImageEl.style.display = "none";
	}
};