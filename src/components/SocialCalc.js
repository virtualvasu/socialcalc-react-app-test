import React, { useEffect, useRef, useState } from 'react';
import { saveAs } from 'file-saver';
import './SocialCalc.css';

const SocialCalc = () => {
  const spreadsheetRef = useRef(null);
  const workbookRef = useRef(null);
  const [spreadsheet, setSpreadsheet] = useState(null);
  const [isLoaded, setIsLoaded] = useState(false);
  const [error, setError] = useState(null);

  useEffect(() => {
    let mounted = true;
    
    // Disable SocialCalc broadcasting and networking to prevent errors
    const disableNetworking = () => {
      if (window.SocialCalc && window.SocialCalc.Callbacks) {
        // Disable broadcast function
        window.SocialCalc.Callbacks.broadcast = () => {
          console.log('Broadcasting disabled in React environment');
        };
      }
      
      // Disable any auto-polling that might be set up
      if (window.updater && window.updater.poll) {
        window.updater.poll = () => {
          console.log('Polling disabled in React environment');
        };
      }
      
      // Disable player initialization if it exists
      if (window.player && window.player.initialize) {
        window.player.initialize = () => {
          console.log('Player initialization disabled in React environment');
        };
      }
    };

    // Wait for SocialCalc to be available
    const initializeSocialCalc = () => {
      if (!mounted) return;
      
      console.log('SocialCalc check:', {
        SocialCalc: !!window.SocialCalc,
        SpreadsheetControl: !!window.SocialCalc?.SpreadsheetControl,
        jQuery: !!window.$
      });
      
      if (window.SocialCalc && window.SocialCalc.SpreadsheetControl) {
        try {
          // Disable networking first
          disableNetworking();
          
          // Ensure image paths are properly set
          if (window.SocialCalc.Constants) {
            window.SocialCalc.Constants.defaultImagePrefix = "/src/images/sc-";
          }
          if (window.SocialCalc.Popup) {
            window.SocialCalc.Popup.imagePrefix = "/src/images/sc-";
          }
          
          // Check if already initialized to prevent duplicates
          if (spreadsheetRef.current && spreadsheetRef.current.innerHTML.trim() !== '') {
            console.log('Spreadsheet already initialized, skipping...');
            return;
          }
          
          setError(null);
          console.log('Creating SocialCalc SpreadsheetControl...');
          
          // Create spreadsheet control
          const spreadsheetControl = new window.SocialCalc.SpreadsheetControl();
          
          // Initialize the spreadsheet in the container
          if (spreadsheetRef.current && mounted) {
            console.log('Initializing spreadsheet in container...');
            
            // Clear any existing content
            spreadsheetRef.current.innerHTML = '';
            
            // Get container dimensions
            const container = spreadsheetRef.current;
            const containerWidth = container.offsetWidth || window.innerWidth - 40;
            const containerHeight = container.offsetHeight || 500;
            
            console.log('Container dimensions:', containerWidth, 'x', containerHeight);
            
            // Initialize with container dimensions
            spreadsheetControl.InitializeSpreadsheetControl(
              spreadsheetRef.current, 
              containerHeight - 10, // height (leave some margin)
              containerWidth - 10,   // width (leave some margin)
              0    // spacebelow
            );
            
            console.log('Spreadsheet control initialized');
            
            // Set a timeout to check if initialization was successful
            setTimeout(() => {
              if (mounted && spreadsheetControl.sheet) {
                setSpreadsheet(spreadsheetControl);
                setIsLoaded(true);
                console.log('SocialCalc initialized successfully');
                
                // Force the spreadsheet to recalculate its layout and expand to full width
                setTimeout(() => {
                  if (spreadsheetControl.editor && spreadsheetControl.editor.context) {
                    try {
                      // Force a redraw to ensure proper sizing
                      spreadsheetControl.editor.context.showDisplayCellsAndStatus();
                      
                      // Ensure the grid uses the full available width
                      const gridElement = spreadsheetRef.current.querySelector('.SocialCalc-grid-control');
                      if (gridElement) {
                        gridElement.style.width = '100%';
                        gridElement.style.minWidth = '100%';
                      }
                      
                      // Ensure editing pane uses full width
                      const editingPane = spreadsheetRef.current.querySelector('.SocialCalc-editingpane');
                      if (editingPane) {
                        editingPane.style.width = '100%';
                        editingPane.style.minWidth = '100%';
                      }
                      
                      console.log('Forced spreadsheet layout recalculation');
                    } catch (e) {
                      console.warn('Could not force layout recalculation:', e);
                    }
                  }
                }, 500);
              } else {
                setError('SocialCalc failed to create sheet object');
              }
            }, 1000);
          }
        } catch (error) {
          console.error('Error initializing SocialCalc:', error);
          if (mounted) {
            setError('Failed to initialize SocialCalc: ' + error.message);
          }
        }
      } else {
        // Retry after a delay if SocialCalc is not loaded yet
        console.log('SocialCalc not ready, retrying in 500ms...');
        setTimeout(initializeSocialCalc, 500);
      }
    };

    // Start initialization after a brief delay
    setTimeout(initializeSocialCalc, 1000);

    // Add resize handler to make spreadsheet responsive
    const handleResize = () => {
      if (spreadsheet && spreadsheetRef.current) {
        // Force recalculation of the spreadsheet layout
        if (spreadsheet.editor && spreadsheet.editor.context) {
          spreadsheet.editor.context.showDisplayCellsAndStatus();
        }
      }
    };

    // Add event listener for window resize
    window.addEventListener('resize', handleResize);

    // Cleanup function
    return () => {
      mounted = false;
      window.removeEventListener('resize', handleResize);
      if (spreadsheetRef.current) {
        spreadsheetRef.current.innerHTML = '';
      }
      if (workbookRef.current) {
        workbookRef.current.innerHTML = '';
      }
    };
  }, []); // Empty dependency array is intentional

  const exportAsCSV = () => {
    if (spreadsheet && spreadsheet.sheet) {
      try {
        // Simple CSV export - convert cell data to CSV format
        const sheet = spreadsheet.sheet;
        const cells = sheet.cells || {};
        
        // Find the range of data
        let maxRow = 0;
        let maxCol = 0;
        const cellKeys = Object.keys(cells);
        
        cellKeys.forEach(cellId => {
          const coords = window.SocialCalc.coordToCr(cellId);
          if (coords) {
            maxRow = Math.max(maxRow, coords.row);
            maxCol = Math.max(maxCol, coords.col);
          }
        });
        
        // Generate CSV content
        let csvContent = '';
        for (let row = 1; row <= maxRow; row++) {
          let rowData = [];
          for (let col = 1; col <= maxCol; col++) {
            const cellId = window.SocialCalc.crToCoord(col, row);
            const cell = cells[cellId];
            let value = '';
            
            if (cell) {
              if (cell.datavalue !== undefined) {
                value = cell.datavalue;
              } else if (cell.displaystring !== undefined) {
                value = cell.displaystring;
              }
              
              // Escape commas and quotes in CSV
              if (typeof value === 'string' && (value.includes(',') || value.includes('"') || value.includes('\n'))) {
                value = '"' + value.replace(/"/g, '""') + '"';
              }
            }
            rowData.push(value);
          }
          csvContent += rowData.join(',') + '\n';
        }
        
        // Use file-saver to download
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8' });
        saveAs(blob, 'spreadsheet.csv');
        
        console.log('CSV export successful');
      } catch (error) {
        console.error('Error exporting CSV:', error);
        alert('Error exporting CSV. Please try again.');
      }
    } else {
      alert('No spreadsheet data available to export.');
    }
  };

  const importCSV = (event) => {
    const file = event.target.files[0];
    if (file && spreadsheet) {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const csvData = e.target.result;
          const lines = csvData.split('\n');
          const sheet = spreadsheet.sheet;
          
          // Clear existing data
          sheet.cells = {};
          
          // Parse CSV and populate cells
          lines.forEach((line, rowIndex) => {
            if (line.trim()) {
              const values = parseCSVLine(line);
              values.forEach((value, colIndex) => {
                if (value.trim()) {
                  const cellId = window.SocialCalc.crToCoord(colIndex + 1, rowIndex + 1);
                  const numValue = parseFloat(value);
                  
                  sheet.cells[cellId] = {
                    datavalue: isNaN(numValue) ? value : numValue,
                    datatype: isNaN(numValue) ? 't' : 'n'
                  };
                }
              });
            }
          });
          
          spreadsheet.ScheduleRender();
          console.log('CSV import successful');
        } catch (error) {
          console.error('Error importing CSV:', error);
          alert('Error importing CSV. Please check the file format.');
        }
      };
      reader.readAsText(file);
    }
    // Reset the file input
    event.target.value = '';
  };

  // Helper function to parse CSV line with proper comma/quote handling
  const parseCSVLine = (line) => {
    const result = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      
      if (char === '"') {
        if (inQuotes && line[i + 1] === '"') {
          current += '"';
          i++; // Skip next quote
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === ',' && !inQuotes) {
        result.push(current);
        current = '';
      } else {
        current += char;
      }
    }
    
    result.push(current);
    return result;
  };

  const clearSpreadsheet = () => {
    if (spreadsheet && spreadsheet.sheet) {
      try {
        // Clear all cells
        spreadsheet.sheet.cells = {};
        spreadsheet.sheet.changes = {};
        spreadsheet.ScheduleRender();
        console.log('Spreadsheet cleared');
      } catch (error) {
        console.error('Error clearing spreadsheet:', error);
      }
    }
  };

  const addSampleData = () => {
    if (spreadsheet && spreadsheet.sheet) {
      try {
        // Add some sample data
        const sheet = spreadsheet.sheet;
        
        // Headers
        sheet.cells.A1 = { datavalue: 'Name', datatype: 't' };
        sheet.cells.B1 = { datavalue: 'Age', datatype: 't' };
        sheet.cells.C1 = { datavalue: 'Score', datatype: 't' };
        
        // Data rows
        sheet.cells.A2 = { datavalue: 'John Doe', datatype: 't' };
        sheet.cells.B2 = { datavalue: 25, datatype: 'n' };
        sheet.cells.C2 = { datavalue: 95, datatype: 'n' };
        
        sheet.cells.A3 = { datavalue: 'Jane Smith', datatype: 't' };
        sheet.cells.B3 = { datavalue: 30, datatype: 'n' };
        sheet.cells.C3 = { datavalue: 87, datatype: 'n' };
        
        sheet.cells.A4 = { datavalue: 'Bob Johnson', datatype: 't' };
        sheet.cells.B4 = { datavalue: 28, datatype: 'n' };
        sheet.cells.C4 = { datavalue: 92, datatype: 'n' };
        
        // Formula
        sheet.cells.C5 = { datavalue: '=AVERAGE(C2:C4)', datatype: 'f' };
        
        spreadsheet.ScheduleRender();
        console.log('Sample data added');
      } catch (error) {
        console.error('Error adding sample data:', error);
      }
    }
  };

  const showDiagnostics = () => {
    console.log('=== SocialCalc Diagnostics ===');
    console.log('window.SocialCalc:', window.SocialCalc);
    console.log('SpreadsheetControl:', window.SocialCalc?.SpreadsheetControl);
    console.log('spreadsheet instance:', spreadsheet);
    console.log('spreadsheet.sheet:', spreadsheet?.sheet);
    console.log('Container element:', spreadsheetRef.current);
    console.log('Container innerHTML length:', spreadsheetRef.current?.innerHTML?.length);
    console.log('isLoaded:', isLoaded);
    console.log('error:', error);
  };

  return (
    <div className="socialcalc-container">
      <div className="socialcalc-header">
        <h2>SocialCalc Spreadsheet Editor</h2>
        <div className="socialcalc-controls">
          <button onClick={addSampleData} disabled={!isLoaded || !!error}>
            Add Sample Data
          </button>
          <button onClick={clearSpreadsheet} disabled={!isLoaded || !!error}>
            Clear All
          </button>
          <button onClick={exportAsCSV} disabled={!isLoaded || !!error}>
            Export CSV
          </button>
          <button onClick={showDiagnostics}>
            Show Diagnostics
          </button>
          <label className="file-input-label">
            Import CSV
            <input
              type="file"
              accept=".csv"
              onChange={importCSV}
              disabled={!isLoaded || !!error}
              style={{ display: 'none' }}
            />
          </label>
        </div>
      </div>
      
      {error && (
        <div className="error-message">
          <strong>Error:</strong> {error}
          <button onClick={() => window.location.reload()}>Reload Page</button>
        </div>
      )}
      
      {!isLoaded && !error && (
        <div className="loading-message">
          <div className="loading-spinner"></div>
          <p>Loading SocialCalc engine...</p>
          <small>This may take a few moments</small>
        </div>
      )}
      
      <div className="socialcalc-workbook" style={{ display: isLoaded && !error ? 'block' : 'none' }}>
        <div ref={workbookRef} id="workbookControl" className="workbook-tabs"></div>
        <div ref={spreadsheetRef} id="spreadsheetControl" className="spreadsheet-container"></div>
      </div>
      
      {isLoaded && !error && (
        <div className="socialcalc-footer">
          <p>
            ✅ SocialCalc loaded successfully! Click on cells to edit, use formulas like =SUM(A1:A10), and try the buttons above.
          </p>
        </div>
      )}
    </div>
  );
};

export default SocialCalc;
