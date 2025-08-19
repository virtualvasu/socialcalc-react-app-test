# React SocialCalc Application

A modern React application that integrates the modernized SocialCalc spreadsheet engine, providing a powerful web-based spreadsheet editor with collaborative features.

## Features

- **Modern React Interface**: Clean, responsive UI built with React
- **SocialCalc Integration**: Full integration with the modernized SocialCalc engine
- **CSV Import/Export**: Easy data import and export functionality
- **Sample Data**: Quick sample data insertion for testing
- **Collaborative Ready**: Built-in support for real-time collaboration
- **Responsive Design**: Works on desktop and mobile devices

## Quick Start

### Prerequisites

- Node.js (v14 or higher)
- npm or yarn

### Installation

1. Navigate to the React app directory:
   ```bash
   cd socialcalc-react-app
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Start the development server:
   ```bash
   npm start
   ```

4. Open your browser to `http://localhost:3000` (or the port shown in terminal)

## Usage

### Basic Spreadsheet Operations

1. **Cell Editing**: Click on any cell to start editing
2. **Formulas**: Enter formulas starting with `=` (e.g., `=SUM(A1:A10)`, `=AVERAGE(B1:B5)`)
3. **Navigation**: Use arrow keys or click to navigate between cells
4. **Data Entry**: Type text or numbers directly into cells

### Available Actions

- **Add Sample Data**: Populate the spreadsheet with example data
- **Clear All**: Remove all data from the spreadsheet
- **Export CSV**: Download the current spreadsheet as a CSV file
- **Import CSV**: Upload and load a CSV file into the spreadsheet

### Supported Formulas

The SocialCalc engine supports spreadsheet functions like:
- **Math**: `SUM`, `AVERAGE`, `COUNT`, `MIN`, `MAX`
- **Logical**: `IF`, `AND`, `OR`, `NOT`
- **Text**: `CONCATENATE`, `LEFT`, `RIGHT`, `MID`, `LEN`
- **Date**: `TODAY`, `NOW`, `DATE`, `YEAR`, `MONTH`, `DAY`

## Project Structure

```
socialcalc-react-app/
├── public/
│   ├── src/              # SocialCalc source files
│   │   ├── js/           # JavaScript modules
│   │   ├── css/          # Stylesheets
│   │   └── images/       # UI assets
│   └── lib/              # Third-party libraries
├── src/
│   ├── components/
│   │   ├── SocialCalc.js # Main spreadsheet component
│   │   └── SocialCalc.css # Component styles
│   ├── App.js            # Main React application
│   └── App.css           # Application styles
└── package.json
```

## Integration with SocialCalc

This React application integrates with the modernized SocialCalc project located in the parent directory (`../socialcalc-modernized-v2`). The SocialCalc source files are copied to the `public/` directory to be accessible by the browser.

## Available Scripts

### `npm start`

Runs the app in development mode. Open [http://localhost:3000](http://localhost:3000) to view it in your browser.

### `npm test`

Launches the test runner in interactive watch mode.

### `npm run build`

Builds the app for production to the `build` folder. The build is optimized for best performance.

### `npm run eject`

**Note: this is a one-way operation. Once you `eject`, you can't go back!**

## Browser Compatibility

- Chrome 60+
- Firefox 55+
- Safari 12+
- Edge 79+

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project integrates with SocialCalc (MIT License) and the React wrapper is also MIT licensed.

---

Built with React and the modernized SocialCalc engine.
