import React, { useState } from 'react';

interface RepairToolsProps {
  fileName: string;
  errorType?: string;
}

export default function RepairTools({ fileName, errorType }: RepairToolsProps) {
  const [showTools, setShowTools] = useState(false);
  
  const toggleTools = () => {
    setShowTools(!showTools);
  };
  
  // Determine file type
  const isXlsx = fileName.toLowerCase().endsWith('.xlsx');
  const isXls = fileName.toLowerCase().endsWith('.xls');
  const isXlsm = fileName.toLowerCase().endsWith('.xlsm');
  
  // Determine repair options based on error type
  const isZipError = errorType?.includes('ZIP') || errorType?.includes('compression');
  const isPasswordError = errorType?.includes('password') || errorType?.includes('encrypted');
  const isCorruptionError = errorType?.includes('corrupt') || errorType?.includes('invalid');
  
  // Generate Excel repair command instructions
  const excelRepairCommand = `
    1. Open Microsoft Excel (don't open the file directly)
    2. Click on File > Open
    3. Navigate to your file: ${fileName}
    4. Instead of clicking "Open", click the small arrow next to the Open button
    5. Select "Open and Repair" from the dropdown menu
  `;
  
  return (
    <div className="repair-tools">
      <button 
        className="repair-tools-toggle"
        onClick={toggleTools}
      >
        {showTools ? 'Hide Repair Tools' : 'Show Repair Tools & Resources'}
      </button>
      
      {showTools && (
        <div className="repair-tools-content">
          <h3>Recommended Repair Methods</h3>
          
          <div className="repair-method">
            <h4>Excel's Built-in Repair</h4>
            <pre>{excelRepairCommand}</pre>
          </div>
          
          {isZipError && isXlsx && (
            <div className="repair-method">
              <h4>ZIP Repair for XLSX Files</h4>
              <p>XLSX files are actually ZIP archives containing XML files. If the ZIP structure is corrupted, try:</p>
              <ol>
                <li>Rename your file from .xlsx to .zip</li>
                <li>Use a ZIP repair tool to fix the file</li>
                <li>After repair, rename it back to .xlsx</li>
              </ol>
              <p>Recommended ZIP repair tools:</p>
              <ul>
                <li><a href="https://www.7-zip.org/" target="_blank" rel="noopener noreferrer">7-Zip</a></li>
                <li><a href="https://www.diskinternals.com/zip-repair/" target="_blank" rel="noopener noreferrer">DiskInternals ZIP Repair</a></li>
              </ul>
            </div>
          )}
          
          {isPasswordError && (
            <div className="repair-method">
              <h4>Password Protected File</h4>
              <p>This file appears to be password protected. Options:</p>
              <ol>
                <li>Contact the file creator for the password</li>
                <li>If you know the password, open in Excel and save as a new file</li>
                <li>For forgotten passwords, commercial recovery tools may help, but success is not guaranteed</li>
              </ol>
            </div>
          )}
          
          {isCorruptionError && (
            <div className="repair-method">
              <h4>Excel File Recovery</h4>
              <p>For severely corrupted files, try these options:</p>
              <ol>
                <li>Use Excel's "Open and Repair" option (as described above)</li>
                <li>Check if you have an earlier version of the file or a backup</li>
                <li>If the file is critically important, commercial recovery tools may help:</li>
              </ol>
              <ul>
                <li><a href="https://www.stellarinfo.com/excel-recovery.php" target="_blank" rel="noopener noreferrer">Stellar Repair for Excel</a></li>
                <li><a href="https://www.nucleustechnologies.com/excel-repair.html" target="_blank" rel="noopener noreferrer">Kernel for Excel Repair</a></li>
              </ul>
            </div>
          )}
          
          <div className="repair-method">
            <h4>General Excel File Repair Tips</h4>
            <ul>
              <li>Try opening the file in Google Sheets or LibreOffice Calc as an alternative</li>
              <li>If only some parts of the workbook are needed, try the "Import" feature in Excel</li>
              <li>Check for auto-recovery files in Excel's auto-recovery location</li>
              <li>For recurrent corruption issues, check your storage device for errors</li>
            </ul>
          </div>
          
          <div className="disclaimer">
            <p><strong>Disclaimer:</strong> These third-party tools are not affiliated with this application. Use them at your own risk.</p>
          </div>
        </div>
      )}
    </div>
  );
} 