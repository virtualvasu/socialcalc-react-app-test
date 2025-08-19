/**
 * SocialCalc Touch Interface
 * 
 * Provides comprehensive touch gesture support for SocialCalc spreadsheet,
 * including tap, double-tap, swipe, and drag interactions. Touch gestures
 * are mapped to corresponding mouse events for consistency.
 *
 * Features:
 * - Single/double tap for cell selection and editing
 * - Swipe gestures for scrolling (up/down, left/right)
 * - Touch-and-drag for range selection
 * - Mobile device detection and optimization
 *
 * @author Ramu Ramamurthy
 * @module TouchInterface
 */

//
//
// To initialize, SocialCalc.CreateTableEditor must call 
//    SocialCalc.TouchRegister(editor.toplevel, {Swipe: SocialCalc.EditorProcessSwipe, editor: editor});
//    Also, any element create inside the grid needs to be touch registered for it to work
//        this includes buttons, etc
//
// TestCases:
//     1) tap to point to cell, and to move edit-cell
//     2) double-tap on a cell to open edit-box, then tap on edit box to pullup keyboard
//          -- when editing, tap on any other cell to be done with edit
//              *** figure out escape to cancel edit versus return to accept edit
//          -- when editing, and if = seen than tapping on cell puts the cell name into the edit box
//     3) tap on input box on top opens edit on cell
//        tap on any cell accepts and finishes this edit
//        *** need a way to cancel the edit versus accept the edit
//
//     4) tap on a cell and move finger starts a range
//        removing finger completes the range
//        tap on cell cancels range
//       
//     5) a swipe action scrolls the sheet
//        swipe up/dn scrolls up/dn
//        swipe lt/rt scrolls lt/rt
// 
//
// Wishlist:
//     1) smooth scroll will be nice (in addition to swipe) -- using a two-touch pan ?
//


// Initialize SocialCalc namespace
const SocialCalc = window.SocialCalc || (() => {
    console.error('Main SocialCalc code module needed');
    return window.SocialCalc = {};
})();

/**
 * Touch capability detection and initialization
 */

/**
 * Indicates whether the current device supports touch events
 * @type {boolean}
 */
SocialCalc.HasTouch = (() => {
    // Modern touch detection using multiple methods
    if ('ontouchstart' in window || 
        window.navigator.maxTouchPoints > 0 || 
        window.navigator.msMaxTouchPoints > 0) {
        return true;
    }
    
    // Fallback: User agent detection for legacy devices
    const agent = navigator.userAgent.toLowerCase();
    return agent.includes('iphone') || 
           agent.includes('ipad') || 
           agent.includes('android');
})();



/**
 * Touch interface configuration and state management
 * 
 * Manages touch gesture recognition, timing thresholds, coordinate tracking,
 * and registered elements for touch event handling.
 * 
 * @namespace TouchInfo
 */
SocialCalc.TouchInfo = {
    /**
     * Array of registered elements that respond to touch events
     * Each element is an object with: { element, functionobj }
     * @type {Array<{element: HTMLElement, functionobj: Object}>}
     */
    registeredElements: [],
    
    // Swipe detection thresholds
    /** @type {number} Minimum horizontal pixel movement to trigger swipe */
    threshold_x: 20,
    /** @type {number} Minimum vertical pixel movement to trigger swipe */
    threshold_y: 20,
    
    // Coordinate tracking for touch events
    /** @type {number} Starting X coordinate of touch */
    orig_coord_x: 0,
    /** @type {number} Starting Y coordinate of touch */
    orig_coord_y: 0,
    /** @type {number} Final X coordinate of touch */
    final_coord_x: 0,
    /** @type {number} Final Y coordinate of touch */
    final_coord_y: 0,
    
    // Scroll sensitivity settings
    /** @type {number} Pixels per row for scrolling (20 pixels = 1 row) */
    px_to_rows: 20,
    /** @type {number} Pixels per column for scrolling (20 pixels = 1 col) */
    px_to_cols: 20,
    
    // Touch timing and state
    /** @type {number} Timestamp when touch started */
    touch_start: 0,
    /** @type {boolean} Whether currently in range selection mode */
    ranging: false,
    /** @type {number} Milliseconds before touch movement triggers ranging */
    ranging_threshold: 100,
    /** @type {number} Timestamp when movement started */
    move_start: 0,
    
    // Double-tap detection
    /** @type {number} Timestamp of last touch for double-tap detection */
    last_touch: 0,
    /** @type {number|null} Timeout handle for single tap delay */
    timeout_handle: null,
    /** @type {number} Maximum milliseconds between taps for double-tap */
    doubletap_threshold: 500
};


/**
 * Registers an element to respond to touch events
 * 
 * Adds touch event listeners to the specified element and registers it
 * in the touch system for gesture recognition.
 * 
 * @param {HTMLElement} element - DOM element to register for touch events
 * @param {Object} functionobj - Object containing touch event handlers
 * @param {Function} [functionobj.Swipe] - Handler for swipe gestures
 * @param {Function} [functionobj.SingleTap] - Handler for single tap
 * @param {Function} [functionobj.DoubleTap] - Handler for double tap
 */
SocialCalc.TouchRegister = (element, functionobj) => {
    // Skip registration if touch is not supported
    if (!SocialCalc.HasTouch) {
        return;
    }

    const touchinfo = SocialCalc.TouchInfo;

    // Check if element is already registered
    if (SocialCalc.LookupElement(element, touchinfo.registeredElements)) {
        return; // Already registered
    }

    // Add element to registered elements array
    touchinfo.registeredElements.push({
        element,
        functionobj
    });
    
    // Add touch event listeners if supported
    if (SocialCalc.HasTouch && element.addEventListener) {
        element.addEventListener('touchstart', SocialCalc.ProcessTouchStart, { passive: false });
        element.addEventListener('touchmove', SocialCalc.ProcessTouchMove, { passive: false });
        element.addEventListener('touchend', SocialCalc.ProcessTouchEnd, { passive: false });
        element.addEventListener('touchcancel', SocialCalc.ProcessTouchCancel, { passive: false });
        
        // Future: orientation change support
        // element.addEventListener('orientationchange', SocialCalc.ProcessOrientationChange, false);
    }
};

/**
 * Finds the registered touch element from a touch event
 * 
 * Traverses up the DOM tree from the event target to find the nearest
 * element that has been registered for touch events.
 * 
 * @param {TouchEvent} event - Touch event object
 * @returns {Object|null} Registered element object or null if not found
 */
SocialCalc.FindTouchElement = (event) => {
    const touchinfo = SocialCalc.TouchInfo;
    const actualEvent = event || window.event;
    
    let ele = actualEvent.target || actualEvent.srcElement;
    let wobj = null;
    
    // Traverse up the DOM tree to find registered element
    while (!wobj && ele) {
        wobj = SocialCalc.LookupElement(ele, touchinfo.registeredElements);
        ele = ele.parentNode;
    }
    
    return wobj;
};
    
/**
 * Processes touch start events
 * 
 * Captures initial touch coordinates and timestamp for gesture recognition.
 * 
 * @param {TouchEvent} event - Touch start event
 */
SocialCalc.ProcessTouchStart = (event) => {
    const touchinfo = SocialCalc.TouchInfo;
    
    // Capture initial touch coordinates
    if (event.targetTouches && event.targetTouches.length > 0) {
        const touch = event.targetTouches[0];
        touchinfo.orig_coord_x = touch.pageX;
        touchinfo.orig_coord_y = touch.pageY;
        
        // Initialize final coordinates
        touchinfo.final_coord_x = touchinfo.orig_coord_x;
        touchinfo.final_coord_y = touchinfo.orig_coord_y;
    }

    // Record touch start timestamp
    touchinfo.touch_start = Date.now();
    
    // Prevent default browser behavior
    event.preventDefault();
};

/**
 * Creates a simulated mouse event from a touch event
 * 
 * Converts touch coordinates and properties into a corresponding
 * mouse event for compatibility with existing mouse handlers.
 * 
 * @param {TouchEvent} event - Original touch event
 * @param {string} mouseEventName - Type of mouse event to create
 * @returns {MouseEvent} Simulated mouse event
 */
SocialCalc.TouchGetSimulatedMouseEvent = (event, mouseEventName) => {
    const touches = event.changedTouches;
    
    if (!touches || touches.length === 0) {
        throw new Error('No touch data available for mouse event simulation');
    }
    
    const touch = touches[0];
    
    // Create and initialize mouse event with touch coordinates
    const simulatedEvent = document.createEvent('MouseEvent');
    simulatedEvent.initMouseEvent(
        mouseEventName,    // event type
        true,              // bubbles
        true,              // cancelable
        window,            // view
        1,                 // detail (click count)
        touch.screenX,     // screenX
        touch.screenY,     // screenY
        touch.clientX,     // clientX
        touch.clientY,     // clientY
        false,             // ctrlKey
        false,             // altKey
        false,             // shiftKey
        false,             // metaKey
        0,                 // button (left mouse button)
        null               // relatedTarget
    );
    
    return simulatedEvent;
};

/**
 * Processes touch move events
 * 
 * Handles touch movement for range selection and drag operations.
 * Simulates mouse events when appropriate thresholds are exceeded.
 * 
 * @param {TouchEvent} event - Touch move event
 */
SocialCalc.ProcessTouchMove = (event) => {
    const touchinfo = SocialCalc.TouchInfo;
    
    // Update final coordinates with current touch position
    if (event.targetTouches && event.targetTouches.length > 0) {
        const touch = event.targetTouches[0];
        touchinfo.final_coord_x = touch.pageX;
        touchinfo.final_coord_y = touch.pageY;
    }

    const wobj = SocialCalc.FindTouchElement(event);
    if (!wobj) return; // Not one of our registered elements

    // Initialize movement tracking
    if (touchinfo.move_start === 0) {
        touchinfo.move_start = Date.now();
        
        // Check if this is a delayed move (for range selection)
        const moveDelay = touchinfo.move_start - touchinfo.touch_start;
        if (moveDelay > touchinfo.ranging_threshold) {
            // Start range selection mode
            touchinfo.ranging = true;
            
            try {
                const mouseDn = SocialCalc.TouchGetSimulatedMouseEvent(event, 'mousedown');
                if (wobj.functionobj?.editor?.fullgrid) {
                    wobj.functionobj.editor.fullgrid.dispatchEvent(mouseDn);
                }
            } catch (error) {
                console.error('Error simulating mousedown event:', error);
            }
        }
    } else if (touchinfo.ranging) {
        // Continue range selection with mouse move simulation
        try {
            const mouseMv = SocialCalc.TouchGetSimulatedMouseEvent(event, 'mousemove');
            if (wobj.functionobj?.editor?.fullgrid) {
                wobj.functionobj.editor.fullgrid.dispatchEvent(mouseMv);
            }
        } catch (error) {
            console.error('Error simulating mousemove event:', error);
        }
    }

    event.preventDefault();
};

/**
 * Processes touch end events
 * 
 * Determines the type of gesture performed (tap, double-tap, swipe, or drag)
 * and dispatches appropriate handlers or simulated mouse events.
 * 
 * @param {TouchEvent} e - Touch end event
 */
SocialCalc.ProcessTouchEnd = (e) => {
    const touchinfo = SocialCalc.TouchInfo;
    const event = e || window.event;

    // Calculate movement distance
    const changeX = touchinfo.orig_coord_x - touchinfo.final_coord_x;
    const changeY = touchinfo.orig_coord_y - touchinfo.final_coord_y;
    
    const wobj = SocialCalc.FindTouchElement(event);
    if (!wobj) return; // Not one of our registered elements

    // Reset touch state
    touchinfo.move_start = 0;
    touchinfo.touch_start = 0;
    
    if (touchinfo.ranging) {
        // End range selection with mouseup simulation
        touchinfo.ranging = false;
        
        try {
            const mouseUp = SocialCalc.TouchGetSimulatedMouseEvent(event, 'mouseup');
            if (wobj.functionobj?.editor?.fullgrid) {
                wobj.functionobj.editor.fullgrid.dispatchEvent(mouseUp);
            }
        } catch (error) {
            console.error('Error simulating mouseup event:', error);
        }
        
    } else if (Math.abs(changeY) > touchinfo.threshold_y || 
               Math.abs(changeX) > touchinfo.threshold_x) {
        // Handle swipe gesture
        const amount_y = Math.floor(changeY / touchinfo.px_to_rows);
        const amount_x = Math.floor(changeX / touchinfo.px_to_cols);
        
        if (wobj.functionobj?.Swipe) {
            wobj.functionobj.Swipe(event, touchinfo, wobj, amount_y, amount_x);
        }
        
    } else {
        // Handle tap gestures (single or double)
        const now = Date.now();
        const lastTouch = touchinfo.last_touch || (now + 1);
        const delta = now - lastTouch;
        
        // Clear any existing timeout
        if (touchinfo.timeout_handle) {
            clearTimeout(touchinfo.timeout_handle);
            touchinfo.timeout_handle = null;
        }

        if (delta < touchinfo.doubletap_threshold && delta > 0) {
            // Double-tap detected
            if (wobj.functionobj?.DoubleTap) {
                wobj.functionobj.DoubleTap(event, touchinfo, wobj);
            }
        } else {
            // Single tap - use timeout to distinguish from potential double-tap
            touchinfo.last_touch = now;
            
            const timeoutFn = () => {
                if (wobj.functionobj?.SingleTap) {
                    wobj.functionobj.SingleTap(event, touchinfo, wobj);
                }
            };
            
            touchinfo.timeout_handle = setTimeout(timeoutFn, touchinfo.doubletap_threshold);
        }
        
        touchinfo.last_touch = now;
    }
    
    event.preventDefault();
};

/**
 * Processes touch cancel events
 * 
 * Resets all touch state when a touch operation is cancelled
 * (e.g., when interrupted by system UI or notifications).
 * 
 * @param {TouchEvent} event - Touch cancel event
 */
SocialCalc.ProcessTouchCancel = (event) => {
    const wobj = SocialCalc.FindTouchElement(event);

    if (!wobj) {
        return; // Use default behavior
    } 

    // Reset all touch state
    const touchinfo = SocialCalc.TouchInfo;
    Object.assign(touchinfo, {
        orig_coord_x: 0,
        orig_coord_y: 0,
        final_coord_x: 0,    
        final_coord_y: 0,
        move_start: 0,
        touch_start: 0,
        ranging: false
    });
};

/**
 * Processes swipe gestures for editor scrolling
 * 
 * Handles swipe gestures to scroll the spreadsheet in both
 * vertical and horizontal directions based on swipe velocity.
 * 
 * @param {TouchEvent} event - Touch event
 * @param {Object} touchinfo - Touch information object
 * @param {Object} wobj - Wrapped element object
 * @param {number} swipevert - Vertical swipe amount (negative = up, positive = down)
 * @param {number} swipehoriz - Horizontal swipe amount (negative = left, positive = right)
 */
SocialCalc.EditorProcessSwipe = (event, touchinfo, wobj, swipevert, swipehoriz) => {
    if (wobj.functionobj.editor.busy) {
        return; // Ignore if editor is busy
    }

    if (swipevert !== 0 || swipehoriz !== 0) {
        wobj.functionobj.editor.ScrollRelativeBoth(swipevert, swipehoriz);
    }
};


/**
 * Processes single tap gestures for cell selection
 * 
 * Simulates a mouse click by dispatching mousedown and mouseup events
 * to the editor grid for cell selection and focus.
 * 
 * @param {TouchEvent} event - Touch event
 * @param {Object} touchinfo - Touch information object
 * @param {Object} wobj - Wrapped element object
 */
SocialCalc.EditorProcessSingleTap = (event, touchinfo, wobj) => {
    if (wobj.functionobj.editor.busy) {
        return; // Ignore if editor is busy
    }    

    try {
        // Simulate mouse click with mousedown and mouseup events
        const mouseDn = SocialCalc.TouchGetSimulatedMouseEvent(event, "mousedown");
        wobj.functionobj.editor.fullgrid.dispatchEvent(mouseDn);

        const mouseUp = SocialCalc.TouchGetSimulatedMouseEvent(event, "mouseup");
        wobj.functionobj.editor.fullgrid.dispatchEvent(mouseUp);
    } catch (error) {
        console.error('Error processing single tap:', error);
    }
};


/**
 * Processes double tap gestures for cell editing
 * 
 * Simulates a mouse double-click event to trigger cell editing mode,
 * allowing users to directly edit cell content.
 * 
 * @param {TouchEvent} event - Touch event
 * @param {Object} touchinfo - Touch information object
 * @param {Object} wobj - Wrapped element object
 */
SocialCalc.EditorProcessDoubleTap = (event, touchinfo, wobj) => {
    if (wobj.functionobj.editor.busy) {
        return; // Ignore if editor is busy
    }
    
    try {
        // Simulate mouse double-click to enter edit mode
        const mouseDblClick = SocialCalc.TouchGetSimulatedMouseEvent(event, "dblclick");
        wobj.functionobj.editor.fullgrid.dispatchEvent(mouseDblClick);
    } catch (error) {
        console.error('Error processing double tap:', error);
    }
};


/**
 * Processes device orientation change events
 * 
 * Handles device rotation events for mobile devices. Currently shows
 * an alert with orientation info - could be extended to adapt UI layout.
 * 
 * @param {Event} event - Orientation change event
 */
SocialCalc.ProcessOrientationChange = (event) => {
    // TODO: Replace alert with proper UI adaptation
    // Consider using console.log or proper UI feedback instead
    if (typeof window.orientation !== 'undefined') {
        console.log('Device orientation changed to:', window.orientation);
        // alert(window.orientation); // Commented out as alerts are disruptive
    }
};
