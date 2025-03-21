import { CompatibilityReport, CheckResult } from '../types';
import RepairTools from './RepairTools';
import FileRecoveryPreview from './FileRecoveryPreview';

interface ResultViewerProps {
  report: CompatibilityReport | null;
  originalFile?: File;
}

export default function ResultViewer({ report, originalFile }: ResultViewerProps) {
  if (!report) {
    return null;
  }

  // Format file size in readable format
  const formatFileSize = (bytes: number): string => {
    if (bytes < 1024) return bytes + ' bytes';
    else if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    else return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  };

  // Format date
  const formatDate = (date: Date): string => {
    return date.toLocaleString();
  };

  // Group results by status
  const errorResults = report.results.filter(r => r.status === 'error');
  const warningResults = report.results.filter(r => r.status === 'warning');
  const successResults = report.results.filter(r => r.status === 'success');
  
  // Extract potential error types for the repair tools
  let errorTypes = '';
  if (errorResults.length > 0) {
    errorTypes = errorResults.map(result => result.details || '').join(' ');
  }
  
  // Check if this is likely a corrupted file
  const isCorrupted = errorTypes.includes('corrupt') || 
                     errorTypes.includes('ZIP') || 
                     errorTypes.includes('compression') ||
                     errorTypes.toLowerCase().includes('invalid');

  return (
    <div className="result-viewer">
      <div className="report-header">
        <h2>Excel Compatibility Report</h2>
        <div className="file-info">
          <p><strong>File:</strong> {report.fileName}</p>
          <p><strong>Size:</strong> {formatFileSize(report.fileSize)}</p>
          <p><strong>Last Modified:</strong> {formatDate(report.lastModified)}</p>
        </div>
        
        <div className={`compatibility-status ${report.isCompatible ? 'compatible' : 'incompatible'}`}>
          {report.isCompatible ? 
            <span>✓ Excel Compatible</span> : 
            <span>✕ Compatibility Issues Found</span>
          }
        </div>
      </div>

      <div className="results-summary">
        <div className="summary-item errors">
          <div className="count">{errorResults.length}</div>
          <div className="label">Errors</div>
        </div>
        <div className="summary-item warnings">
          <div className="count">{warningResults.length}</div>
          <div className="label">Warnings</div>
        </div>
        <div className="summary-item success">
          <div className="count">{successResults.length}</div>
          <div className="label">Passed Checks</div>
        </div>
      </div>

      <div className="results-details">
        {errorResults.length > 0 && (
          <div className="result-section">
            <h3>Errors</h3>
            <ResultList results={errorResults} />
          </div>
        )}
        
        {warningResults.length > 0 && (
          <div className="result-section">
            <h3>Warnings</h3>
            <ResultList results={warningResults} />
          </div>
        )}
        
        {successResults.length > 0 && (
          <div className="result-section">
            <h3>Passed Checks</h3>
            <ResultList results={successResults} />
          </div>
        )}
      </div>
      
      {/* Show content recovery for corrupted files */}
      {isCorrupted && originalFile && (
        <FileRecoveryPreview file={originalFile} />
      )}
      
      {/* Show repair tools if there are errors */}
      {errorResults.length > 0 && (
        <RepairTools 
          fileName={report.fileName}
          errorType={errorTypes}
        />
      )}
    </div>
  );
}

// Helper component to render a list of results
function ResultList({ results }: { results: CheckResult[] }) {
  return (
    <ul className="result-list">
      {results.map((result, index) => (
        <li key={index} className={`result-item ${result.status}`}>
          <div className="result-message">{result.message}</div>
          {result.details && (
            <div className="result-details">
              {result.details.split('\n').map((line, i) => (
                <p key={i}>{line}</p>
              ))}
            </div>
          )}
          {result.location && <div className="result-location">Location: {result.location}</div>}
        </li>
      ))}
    </ul>
  );
} 