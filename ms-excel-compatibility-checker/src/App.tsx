import { useState } from 'react'
import './App.css'
import FileUploader from './components/FileUploader'
import ResultViewer from './components/ResultViewer'
import { CompatibilityReport } from './types'

function App() {
  const [report, setReport] = useState<CompatibilityReport | null>(null);
  const [originalFile, setOriginalFile] = useState<File | undefined>(undefined);

  const handleReportGenerated = (newReport: CompatibilityReport, file: File) => {
    setReport(newReport);
    setOriginalFile(file);
  };

  return (
    <div className="app-container">
      <header>
        <h1>Excel Compatibility Checker</h1>
        <p>Upload your Excel file to check for compatibility issues with Microsoft Excel</p>
      </header>

      <main>
        <FileUploader onReportGenerated={handleReportGenerated} />
        {report && <ResultViewer report={report} originalFile={originalFile} />}
      </main>

      <footer>
        <p>
          Built with React, Vite and TypeScript. 
          Uses <a href="https://sheetjs.com/" target="_blank" rel="noopener noreferrer">SheetJS</a> and 
          <a href="https://github.com/exceljs/exceljs" target="_blank" rel="noopener noreferrer"> ExcelJS</a> for Excel file processing.
        </p>
      </footer>
    </div>
  )
}

export default App
