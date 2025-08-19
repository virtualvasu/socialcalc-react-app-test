import React from 'react';
import './App.css';
import SocialCalc from './components/SocialCalc';

function App() {
  return (
    <div className="App">
      <header className="App-header">
        <h1>React SocialCalc Application</h1>
        <p>A modern React wrapper for the SocialCalc spreadsheet engine</p>
      </header>
      
      <main className="App-main">
        <SocialCalc />
      </main>
      
      <footer className="App-footer">
        <p>Built with React and SocialCalc • Supports CSV import/export and collaborative editing</p>
      </footer>
    </div>
  );
}

export default App;
