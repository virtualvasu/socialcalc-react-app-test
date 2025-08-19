/**
 * EtherCalc Real-time Collaboration Module
 * Handles collaborative editing features including broadcasting changes,
 * polling for updates, and synchronizing state across multiple clients.
 */

// Initialize console polyfill for older browsers
if (!window.console) window.console = {};
if (!window.console.log) window.console.log = () => {};

// Disable auto-initialization in React environment
// Comment out the following lines to prevent automatic polling
// 
// Initialize collaboration system when DOM is ready
// $(() => {
//     player.initialize();
//     updater.poll();
// });


/**
 * Check if the current session supports multiple sheets
 * @returns {boolean} True if workbook control object exists
 */
const isMultipleSheet = () => {
    return !!SocialCalc.CurrentWorkbookControlObject;
};

/**
 * Get cookie value by name
 * @param {string} name - Cookie name to retrieve
 * @returns {string|undefined} Cookie value or undefined if not found
 */
const getCookie = (name) => {
    const match = document.cookie.match(`\\b${name}=([^;]*)\\b`);
    return match ? match[1] : undefined;
};

/**
 * jQuery extension for JSON POST requests
 * @param {string} url - Request URL
 * @param {Object} args - Request parameters
 * @param {Function} [callback] - Success callback function
 */
jQuery.postJSON = (url, args, callback) => {
    const paramString = $.param(args);
    
    $.ajax({
        url,
        data: paramString,
        dataType: 'text',
        type: 'POST',
        success(response) {
            if (callback) {
                try {
                    callback(JSON.parse(response));
                } catch (e) {
                    console.error('Failed to parse JSON response:', e);
                }
            }
        },
        error(response) {
            console.error('AJAX Error:', response);
        }
    });
};

/**
 * Broadcast collaboration events to other clients
 * @param {string} type - Event type
 * @param {Object} data - Event data
 */
/**
 * Safe broadcast function that prevents circular structure errors
 * @param {string} type - Event type
 * @param {*} data - Data to broadcast (will be safely serialized)
 */
SocialCalc.Callbacks.broadcast = (type, data) => {
    // Skip certain event types that might cause issues
    if (type === 'ask.ecell' || type === 'ecell' || type === 'step') {
        return;
    }
    
    // Check if we're in a React environment without backend
    if (window.location.hostname === 'localhost' && 
        (window.location.port === '3000' || window.location.port === '3001')) {
        console.log('Broadcasting disabled in React development environment');
        return;
    }
    
    try {
        // Create a safe copy of data without circular references
        let safeData = data;
        if (typeof data === 'object' && data !== null) {
            // Create a simplified version of the data
            safeData = {
                type: type,
                timestamp: Date.now()
            };
            
            // Only include simple properties
            if (data.cmdstr) safeData.cmdstr = data.cmdstr;
            if (data.savestr) safeData.savestr = data.savestr;
            if (typeof data === 'string') safeData = data;
        }
        
        const message = {
            from: player.idInSession || 'react-client',
            type,
            data: encodeURIComponent(JSON.stringify(safeData))
        };
        
        // Only broadcast if we have a proper backend
        if (typeof $.postJSON === 'function') {
            $.postJSON('/broadcast', message, (response) => {
                if (response && updater.showMessage) {
                    updater.showMessage(response);
                }
            });
        }
    } catch (error) {
        console.warn('Broadcast failed:', error.message);
        // Silently fail - this is normal in React environment
    }
};


/**
 * Real-time update polling system
 */
const updater = {
    errorSleepTime: 500,
    cursor: null,

    /**
     * Poll server for new updates
     * Only runs if we're in a proper backend environment
     */
    poll() {
        // Check if we're in a React environment without backend
        if (window.location.hostname === 'localhost' && 
            (window.location.port === '3000' || window.location.port === '3001')) {
            console.log('Polling disabled in React development environment');
            return;
        }
        
        const args = { '_xsrf': getCookie('_xsrf') };
        if (this.cursor) args.cursor = this.cursor;
        
        $.ajax({
            url: '/updates',
            type: 'POST',
            dataType: 'text',
            data: $.param(args),
            success: this.onSuccess.bind(this),
            error: this.onError.bind(this)
        });
    },

    /**
     * Handle successful polling response
     * @param {string} response - JSON response string
     */
    onSuccess(response) {
        try {
            this.newMessages(JSON.parse(response));
        } catch (e) {
            console.error('Failed to parse polling response:', e);
            this.onError();
            return;
        }
        this.errorSleepTime = 500;
        setTimeout(() => this.poll(), 0);
    },

    /**
     * Handle polling errors with exponential backoff
     */
    onError(response) {
        this.errorSleepTime *= 2;
        console.warn(`Poll error; retrying in ${this.errorSleepTime}ms`);
        setTimeout(() => this.poll(), this.errorSleepTime);
    },

    /**
     * Process new messages from server
     * @param {Object} response - Server response containing messages
     */
    newMessages(response) {
        if (!response.messages) return;
        
        const messages = response.messages;
        this.cursor = messages[messages.length - 1].id;
        
        // Process each message
        messages.forEach(message => {
            player.onNewEvent(message);
        });
    },

    /**
     * Display message to user (placeholder)
     * @param {Object} message - Message to display
     */
    showMessage(message) {
        // TODO: Implement user messaging
    }
};

/**
 * Collaboration player system for handling multi-user sessions
 */
const player = {
    
    /**
     * Initialize the collaboration player
     */
    initialize() {
        this.isConnected = true;
        this.idInSession = getCookie('idinsession');
        
        // Request snapshot if not the first user
        if (this.idInSession !== '1') {
            SocialCalc.Callbacks.broadcast('ask.snapshot', { arbit: 'arbit' });
        }
    },
    
    /**
     * Handle new collaboration events
     * @param {Object} data - Event data from server
     */
    onNewEvent(data) {
        if (!this.isConnected) return;
        if (data.from === this.idInSession) return;
        if (typeof SocialCalc === 'undefined') return;
        
        const decodedData = decodeURIComponent(data.data);
        const msgData = JSON.parse(decodedData);

        const editor = SocialCalc.CurrentSpreadsheetControlObject.editor;
        
        switch (data.type) {
        case 'ecell': {
            // Currently disabled peer highlighting
            break;
            const peerClass = ' defaultPeer';
            const find = new RegExp(peerClass, 'g');
	    
            if (msgData.original) {
                const origCR = SocialCalc.coordToCr(msgData.original);
                const origCell = SocialCalc.GetEditorCellElement(editor, origCR.row, origCR.col);
                origCell.element.className = origCell.element.className.replace(find, '');
            }
	    
            const cr = SocialCalc.coordToCr(msgData.ecell);
            const cell = SocialCalc.GetEditorCellElement(editor, cr.row, cr.col);
            if (cell.element.className.search(find) === -1) {
                cell.element.className += peerClass;
            }
            break;
        }
        case 'ask.snapshot': {
            if (this.idInSession === '1') {
                const snapshotData = {
                    to: data.idinsession,
                    snapshot: isMultipleSheet() 
                        ? SocialCalc.WorkBookControlSaveSheet()
                        : SocialCalc.CurrentSpreadsheetControlObject.CreateSpreadsheetSave()
                };
                SocialCalc.Callbacks.broadcast('snapshot', snapshotData);
            }
            break;
        }
        case 'ask.ecell': {
            // Currently disabled
            break;
            
            // Future implementation:
            // SocialCalc.Callbacks.broadcast('ecell', {
            //     to: data.idinsession,
            //     ecell: editor.ecell.coord
            // });
        }
        case 'snapshot': {
            // Prevent duplicate snapshot loading
            if (this._hadSnapshot) break;
            this._hadSnapshot = true;
            
            if (isMultipleSheet()) {
                SocialCalc.WorkBookControlLoad(msgData.snapshot);
            } else {
                const spreadsheet = SocialCalc.CurrentSpreadsheetControlObject;
                const parts = spreadsheet.DecodeSpreadsheetSave(msgData.snapshot);
                
                if (parts && parts.sheet) {
                    spreadsheet.sheet.ResetSheet();
                    spreadsheet.ParseSheetSave(
                        msgData.snapshot.substring(parts.sheet.start, parts.sheet.end)
                    );
                }
                
                // Refresh display based on recalc setting
                const isRecalcOff = spreadsheet.editor.context.sheetobj.attribs.recalc === 'off';
                const command = isRecalcOff ? 'redisplay' : 'recalc';
                spreadsheet.ExecuteCommand(command, '');
            }
            break;
        }
        case 'execute': {
            if (isMultipleSheet()) {
                const control = SocialCalc.GetCurrentWorkBookControl();
                control.ExecuteWorkBookControlCommand(msgData, true); // remote command
            } else {
                SocialCalc.CurrentSpreadsheetControlObject.context.sheetobj.ScheduleSheetCommands(
                    msgData.cmdstr,
                    msgData.saveundo,
                    true // isRemote = true
                );
            }
            break;
        }
        }
    }
};