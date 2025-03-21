import { useState, useRef } from 'react';
import { CompatibilityReport } from '../types';
import { checkExcelCompatibility } from '../utils/excelChecker';

interface FileUploaderProps {
  onReportGenerated: (report: CompatibilityReport, originalFile: File) => void;
}

export default function FileUploader({ onReportGenerated }: FileUploaderProps) {
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleFileDrop = async (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      await processFile(files[0]);
    }
  };

  const handleFileSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      await processFile(files[0]);
    }
  };

  const handleButtonClick = () => {
    fileInputRef.current?.click();
  };

  const processFile = async (file: File) => {
    setError(null);
    setIsProcessing(true);
    
    try {
      // Check if it's an Excel file
      if (!file.name.match(/\.(xlsx|xls|xlsm|xlsb)$/i)) {
        throw new Error('Please upload a valid Excel file (.xlsx, .xls, .xlsm, or .xlsb)');
      }
      
      // Additional check for potentially corrupted ZIP
      if (file.name.endsWith('.xlsx') && file.size < 128) {
        throw new Error('This Excel file appears to be corrupted or empty.');
      }
      
      // Process the file
      try {
        const report = await checkExcelCompatibility(file);
        onReportGenerated(report, file);
      } catch (err) {
        const errorMessage = (err instanceof Error) ? err.message : String(err);
        
        // Provide a more helpful message for common errors
        if (errorMessage.includes('ZIP') || errorMessage.includes('NaN') || errorMessage.includes('compression')) {
          throw new Error('This Excel file uses an unsupported compression method or is corrupted. Please try a different file or resave it in a newer version of Excel.');
        } else {
          throw err;
        }
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An unknown error occurred');
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="file-uploader">
      <div
        className={`upload-area ${isDragging ? 'dragging' : ''}`}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleFileDrop}
        onClick={handleButtonClick}
      >
        <input
          type="file"
          accept=".xlsx,.xls,.xlsm,.xlsb"
          onChange={handleFileSelect}
          ref={fileInputRef}
          style={{ display: 'none' }}
        />
        
        {isProcessing ? (
          <div className="processing">
            <div className="spinner"></div>
            <p>Analyzing Excel file...</p>
          </div>
        ) : (
          <div className="upload-prompt">
            <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="17 8 12 3 7 8"></polyline>
              <line x1="12" y1="3" x2="12" y2="15"></line>
            </svg>
            <h3>Drop your Excel file here</h3>
            <p>or click to browse</p>
          </div>
        )}
      </div>
      
      {error && (
        <div className="error-message">
          <p>{error}</p>
        </div>
      )}
    </div>
  );
} 