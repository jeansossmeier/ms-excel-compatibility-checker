import * as XLSX from 'xlsx';
import * as ExcelJS from 'exceljs';
import { CheckResult, CompatibilityReport } from '../types';

/**
 * More detailed examination of internal Excel file structure
 * Provides advanced diagnostics for corrupted files
 */
async function performDeepExcelAnalysis(file: File): Promise<string[]> {
  const issues: string[] = [];
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  
  // Check file signature
  const hasValidXlsxSignature = bytes.length >= 4 && 
    bytes[0] === 0x50 && bytes[1] === 0x4B && 
    bytes[2] === 0x03 && bytes[3] === 0x04;
    
  const hasValidXlsSignature = bytes.length >= 8 && 
    bytes[0] === 0xD0 && bytes[1] === 0xCF && 
    bytes[2] === 0x11 && bytes[3] === 0xE0;
  
  // Format detection
  if (!hasValidXlsxSignature && !hasValidXlsSignature) {
    issues.push('File does not have a valid Excel signature.');
    
    // Check if it might be another format masquerading as Excel
    if (bytes.length >= 4) {
      // Check for PDF signature
      if (bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46) {
        issues.push('The file appears to be a PDF file with an Excel extension.');
      }
      // Check for JPEG signature
      else if (bytes[0] === 0xFF && bytes[1] === 0xD8 && bytes[2] === 0xFF) {
        issues.push('The file appears to be a JPEG image with an Excel extension.');
      }
      // Check for PNG signature
      else if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4E && bytes[3] === 0x47) {
        issues.push('The file appears to be a PNG image with an Excel extension.');
      }
      // Check for ZIP but not Office Open XML
      else if (hasValidXlsxSignature && bytes.length > 100) {
        const textSample = new TextDecoder().decode(bytes.slice(0, 1000));
        if (!textSample.includes('workbook.xml') && !textSample.includes('[Content_Types].xml')) {
          issues.push('The file is a ZIP archive but does not appear to be a valid Office Open XML document.');
        }
      }
    }
  }
  
  // Excel format identification
  if (hasValidXlsxSignature) {
    issues.push('File has a valid XLSX/ZIP signature.');
    
    // Analyze ZIP structure
    try {
      // Simple ZIP central directory check
      let hasCentralDir = false;
      // Search for End of Central Directory signature
      for (let i = bytes.length - 22; i >= 0; i--) {
        if (bytes[i] === 0x50 && bytes[i + 1] === 0x4B && 
            bytes[i + 2] === 0x05 && bytes[i + 3] === 0x06) {
          hasCentralDir = true;
          break;
        }
      }
      
      if (!hasCentralDir) {
        issues.push('ZIP Central Directory is missing or corrupt.');
        issues.push('This indicates the file was not properly closed when saved or was truncated.');
      }
      
      // Check for common XLSX components in text sample
      const textSample = new TextDecoder().decode(bytes.slice(0, Math.min(bytes.length, 3000)));
      const requiredComponents = [
        '[Content_Types].xml',
        'workbook.xml',
        'sheet',
        'styles.xml',
        'sharedStrings'
      ];
      
      const missingComponents = requiredComponents.filter(component => 
        !textSample.includes(component)
      );
      
      if (missingComponents.length > 0) {
        issues.push(`Missing essential XLSX components: ${missingComponents.join(', ')}`);
        if (missingComponents.includes('[Content_Types].xml')) {
          issues.push('Missing [Content_Types].xml indicates severe corruption.');
        }
      }
    } catch (zipError) {
      issues.push(`Error analyzing ZIP structure: ${zipError instanceof Error ? zipError.message : String(zipError)}`);
    }
  } else if (hasValidXlsSignature) {
    issues.push('File has a valid binary XLS (BIFF) format signature.');
    if (file.name.endsWith('.xlsx')) {
      issues.push('The file is in the older XLS format but has an .xlsx extension. This mismatch can cause compatibility issues.');
    }
  }
  
  // Check file size consistency
  if (bytes.length < 2000 && bytes.length > 100) {
    issues.push('File size is suspiciously small for a functional Excel workbook.');
    
    // Check if it might be a truncated file
    if (hasValidXlsxSignature && !bytes.includes(0x50, bytes.length - 100)) {
      issues.push('File appears to be truncated (premature end of file).');
    }
  }
  
  // Examine for password protection
  if (bytes.length > 500) {
    const textSample = new TextDecoder().decode(bytes.slice(0, Math.min(bytes.length, 2000)));
    
    // Check for encryption markers
    if (textSample.includes('EncryptedPackage') || 
        textSample.includes('encryption') ||
        textSample.includes('EncryptionInfo') ||
        textSample.includes('StrongEncryptionDataSpace')) {
      issues.push('File appears to be password protected or encrypted.');
      issues.push('Password-protected files cannot be analyzed without the correct password.');
    }
  }
  
  return issues;
}

/**
 * Analyzes a corrupted Excel file to determine the specific issue
 * Attempts to analyze the file structure to identify corruption points
 */
async function analyzeCorruptedExcel(file: File, errorMsg: string): Promise<string> {
  // Detailed analysis results
  const issues: string[] = [];
  const recommendations: string[] = [];
  
  // Check for specific error patterns
  if (errorMsg.includes('Unsupported ZIP Compression method NaN')) {
    issues.push('The file uses an unsupported compression method (NaN).');
    issues.push('This typically happens with Excel files saved in very old versions or with non-standard tools.');
    recommendations.push('Try opening the file in Microsoft Excel and using "Save As" to create a new XLSX file.');
    recommendations.push('If the original file cannot be opened, try a file recovery tool like "Office File Recovery".');
  } else if (errorMsg.includes('Local File Header signature not found')) {
    issues.push('ZIP file structure is invalid - Local File Header signature missing.');
    issues.push('This suggests the file has been truncated or corrupted during transfer.');
    recommendations.push('Request the original file again if it was shared with you.');
    recommendations.push('Check if your antivirus software might be corrupting the file during download.');
  } else if (errorMsg.includes('End of Central Directory Record not found')) {
    issues.push('ZIP file structure is invalid - End of Central Directory Record missing.');
    issues.push('This suggests the file is incomplete or was not properly finalized when saved.');
    recommendations.push('Try using a ZIP repair tool to fix the file structure.');
    recommendations.push('If possible, try to recreate the Excel file from the source.');
  } else if (errorMsg.includes('incorrect header check')) {
    issues.push('ZIP compression header is invalid or corrupted.');
    issues.push('This often happens when file transfers are interrupted or when files are corrupted in storage.');
    recommendations.push('Try downloading the file again or requesting a new copy.');
  }
  
  // Perform advanced file structure analysis
  try {
    const deepAnalysisResults = await performDeepExcelAnalysis(file);
    issues.push(...deepAnalysisResults);
    
    // Extract more repair recommendations based on deep analysis
    if (deepAnalysisResults.some(result => result.includes('XLS format'))) {
      recommendations.push('This appears to be a binary XLS file. Try renaming it with .xls extension and opening in Excel.');
    }
    
    if (deepAnalysisResults.some(result => result.includes('password protected'))) {
      recommendations.push('You need the original password to open this file. Contact the file creator for the password.');
    }
    
    if (deepAnalysisResults.some(result => result.includes('truncated'))) {
      recommendations.push('The file appears to be incomplete. Try to obtain a complete copy of the file.');
    }
    
    if (deepAnalysisResults.some(result => result.includes('ZIP Central Directory is missing'))) {
      recommendations.push('Try using a ZIP repair tool like "ZipRepair" or "DiskInternals ZIP Repair" to rebuild the central directory.');
    }
    
    if (deepAnalysisResults.some(result => result.includes('Missing essential XLSX components'))) {
      recommendations.push('The internal structure of the Excel file is damaged. Try Excel\'s built-in repair feature by opening Excel first, then using File > Open and selecting "Open and Repair".');
    }
  } catch (analysisError) {
    issues.push(`Error during detailed analysis: ${analysisError instanceof Error ? analysisError.message : String(analysisError)}`);
  }
  
  // If we couldn't determine anything specific
  if (issues.length === 0) {
    issues.push('Unable to determine the exact corruption issue.');
    issues.push(`Raw error: ${errorMsg}`);
    issues.push('The file might be corrupted, password protected, or use an unsupported format.');
    recommendations.push('Try using Microsoft Excel\'s built-in repair feature.');
    recommendations.push('If the file is crucial, consider using a commercial Excel recovery tool.');
  }
  
  // Build the complete diagnostics report
  const diagnosticsReport = [
    '=== ISSUE DETAILS ===',
    ...issues,
    '',
    '=== REPAIR RECOMMENDATIONS ===',
    ...recommendations,
    '',
    '=== TECHNICAL DETAILS ===',
    `File name: ${file.name}`,
    `File size: ${(file.size / 1024).toFixed(2)} KB`,
    `Last modified: ${new Date(file.lastModified).toLocaleString()}`
  ];
  
  return diagnosticsReport.join('\n');
}

/**
 * Attempts to recover basic information from a corrupted Excel file
 * Returns whatever metadata could be extracted
 */
async function extractBasicMetadata(file: File): Promise<Record<string, string>> {
  const metadata: Record<string, string> = {};
  
  try {
    const buffer = await file.arrayBuffer();
    const bytes = new Uint8Array(buffer);
    const textSample = new TextDecoder().decode(bytes.slice(0, Math.min(bytes.length, 10000)));
    
    // Extract creator info if available
    const creatorMatch = textSample.match(/dc:creator>(.*?)<\/dc:creator/);
    if (creatorMatch && creatorMatch[1]) {
      metadata.creator = creatorMatch[1];
    }
    
    // Extract last modified by info
    const modifiedByMatch = textSample.match(/cp:lastModifiedBy>(.*?)<\/cp:lastModifiedBy/);
    if (modifiedByMatch && modifiedByMatch[1]) {
      metadata.lastModifiedBy = modifiedByMatch[1];
    }
    
    // Extract creation date if available
    const creationDateMatch = textSample.match(/dcterms:created>(.*?)<\/dcterms:created/);
    if (creationDateMatch && creationDateMatch[1]) {
      metadata.creationDate = creationDateMatch[1];
    }
    
    // Extract application info if available
    const applicationMatch = textSample.match(/Application>(.*?)<\/Application/);
    if (applicationMatch && applicationMatch[1]) {
      metadata.application = applicationMatch[1];
    }
    
    // Try to extract sheet names
    const sheetNameMatches = textSample.match(/<sheet name="(.*?)"/g);
    if (sheetNameMatches && sheetNameMatches.length > 0) {
      const sheetNames = sheetNameMatches
        .map(match => {
          const nameMatch = match.match(/<sheet name="(.*?)"/);
          return nameMatch ? nameMatch[1] : null;
        })
        .filter(name => name !== null);
      
      if (sheetNames.length > 0) {
        metadata.sheets = sheetNames.join(', ');
        metadata.sheetCount = sheetNames.length.toString();
      }
    }
    
  } catch (error) {
    metadata.extractionError = error instanceof Error ? error.message : String(error);
  }
  
  return metadata;
}

/**
 * Checks an Excel file for compatibility issues with Microsoft Excel
 * Performs a comprehensive analysis of various Excel features
 */
export async function checkExcelCompatibility(file: File): Promise<CompatibilityReport> {
  // Start with empty results array
  const results: CheckResult[] = [];
  
  // Basic file information
  const fileName = file.name;
  const fileSize = file.size;
  const lastModified = new Date(file.lastModified);
  
  try {
    // Try to extract basic metadata even if file is corrupted
    const metadata = await extractBasicMetadata(file);
    let metadataText = '';
    
    if (Object.keys(metadata).length > 0) {
      metadataText = 'Recovered Metadata:\n' + 
        Object.entries(metadata)
          .map(([key, value]) => `${key}: ${value}`)
          .join('\n');
      
      results.push({
        status: 'success',
        message: 'Successfully extracted some file metadata',
        details: metadataText
      });
    }
    
    // Validate file structure first
    if (!isValidExcelFile(file)) {
      throw new Error('Invalid or corrupted Excel file format');
    }

    // Read file using SheetJS
    const data = await file.arrayBuffer();
    
    // Initialize workbooks with safe defaults
    let workbook;
    let excelJsWorkbook;
    
    try {
      // For XLSX.js, set up more options to handle different compression methods
      const options = { 
        type: 'array' as const,
        cellFormula: true,
        cellStyles: true,
        WTF: true // "What The Format" - more verbose error handling
      };
      workbook = XLSX.read(data, options);
    } catch (sheetJsError) {
      const errorMsg = sheetJsError instanceof Error ? sheetJsError.message : String(sheetJsError);
      
      // Special handling for ZIP compression errors
      if (errorMsg.includes('ZIP') || errorMsg.includes('NaN') || errorMsg.includes('compression')) {
        // Get detailed diagnostics
        const detailedDiagnostics = await analyzeCorruptedExcel(file, errorMsg);
        
        // Add recovered metadata to the diagnostics if available
        const fullDiagnostics = metadataText 
          ? `${detailedDiagnostics}\n\n=== RECOVERED METADATA ===\n${metadataText}`
          : detailedDiagnostics;
          
        throw new Error(`Unsupported ZIP compression method or corrupted Excel file. Please try resaving the file in a newer version of Excel.\n\nDetailed diagnostics:\n${fullDiagnostics}`);
      }
      
      results.push({
        status: 'error',
        message: 'Failed to parse Excel file with SheetJS',
        details: errorMsg
      });
      workbook = { SheetNames: [], Sheets: {} } as XLSX.WorkBook;
    }
    
    try {
      excelJsWorkbook = new ExcelJS.Workbook();
      // Load the workbook without extra options that weren't compatible
      await excelJsWorkbook.xlsx.load(data);
    } catch (excelJsError) {
      const errorMsg = excelJsError instanceof Error ? excelJsError.message : String(excelJsError);
      
      // Special handling for ZIP compression errors
      if (errorMsg.includes('ZIP') || errorMsg.includes('NaN') || errorMsg.includes('compression')) {
        // Only throw if SheetJS didn't already fail
        if (workbook.SheetNames.length === 0) {
          // Get detailed diagnostics
          const detailedDiagnostics = await analyzeCorruptedExcel(file, errorMsg);
          
          // Add recovered metadata to the diagnostics if available
          const fullDiagnostics = metadataText 
            ? `${detailedDiagnostics}\n\n=== RECOVERED METADATA ===\n${metadataText}`
            : detailedDiagnostics;
            
          throw new Error(`Unsupported ZIP compression method or corrupted Excel file. Please try resaving the file in a newer version of Excel.\n\nDetailed diagnostics:\n${fullDiagnostics}`);
        }
      }
      
      results.push({
        status: 'error',
        message: 'Failed to parse Excel file with ExcelJS',
        details: errorMsg
      });
      excelJsWorkbook = new ExcelJS.Workbook();
    }
    
    // Run compatibility checks even if some parsers failed
    checkFileFormat(file, results);
    checkWorkbookStructure(workbook, results);
    checkCellFormats(workbook, results);
    checkFormulas(workbook, results);
    checkCharts(excelJsWorkbook, results);
    checkMacros(workbook, results);
    checkNamedRanges(workbook, results);
    checkConditionalFormatting(excelJsWorkbook, results);
    
    // Calculate summary statistics
    const totalIssues = results.filter(r => r.status !== 'success').length;
    const isCompatible = !results.some(r => r.status === 'error');
    
    return {
      fileName,
      fileSize,
      lastModified,
      results,
      totalIssues,
      isCompatible
    };
  } catch (error) {
    // Handle errors parsing the file
    results.push({
      status: 'error',
      message: 'Failed to process Excel file',
      details: error instanceof Error ? error.message : String(error)
    });
    
    return {
      fileName,
      fileSize,
      lastModified,
      results,
      totalIssues: results.length,
      isCompatible: false
    };
  }
}

/**
 * Basic validation to check if file appears to be a valid Excel file
 */
function isValidExcelFile(file: File): boolean {
  // Check file extension
  if (!file.name.match(/\.(xlsx|xls|xlsm|xlsb)$/i)) {
    return false;
  }
  
  // Check minimum file size (to avoid empty/corrupt files)
  if (file.size < 128) {
    return false;
  }
  
  return true;
}

/**
 * Checks basic file format compatibility
 */
function checkFileFormat(file: File, results: CheckResult[]): void {
  // Check file extension
  if (!file.name.match(/\.(xlsx|xls|xlsm|xlsb)$/i)) {
    results.push({
      status: 'error',
      message: 'File does not have a valid Excel extension',
      details: 'Valid extensions are .xlsx, .xls, .xlsm, and .xlsb'
    });
    return;
  }
  
  // Check file size
  const maxSizeBytes = 100 * 1024 * 1024; // 100MB (Excel's approximate limit)
  if (file.size > maxSizeBytes) {
    results.push({
      status: 'warning',
      message: 'File size exceeds recommended limit',
      details: `The file is ${(file.size / (1024 * 1024)).toFixed(2)}MB. Excel works best with files under 100MB.`
    });
  } else {
    results.push({
      status: 'success',
      message: 'File size is within recommended limits'
    });
  }
}

/**
 * Checks workbook structure for compatibility issues
 */
function checkWorkbookStructure(workbook: XLSX.WorkBook, results: CheckResult[]): void {
  // Check if SheetNames exists and is an array
  if (!workbook.SheetNames || !Array.isArray(workbook.SheetNames)) {
    results.push({
      status: 'error',
      message: 'Invalid workbook structure',
      details: 'Unable to detect worksheet information'
    });
    return;
  }

  // Check number of worksheets
  const sheetCount = workbook.SheetNames.length;
  if (sheetCount > 255) {
    results.push({
      status: 'error',
      message: 'Too many worksheets',
      details: `The workbook contains ${sheetCount} worksheets. Excel supports a maximum of 255.`
    });
  } else {
    results.push({
      status: 'success',
      message: 'Worksheet count is within Excel limits'
    });
  }
  
  // Check worksheet names
  for (const sheetName of workbook.SheetNames) {
    if (typeof sheetName !== 'string') {
      continue; // Skip non-string sheet names
    }
    
    if (sheetName.length > 31) {
      results.push({
        status: 'error',
        message: 'Sheet name too long',
        details: `The sheet "${sheetName}" has a name longer than 31 characters, which is not supported by Excel.`,
        location: sheetName
      });
    }
    
    if (/[\\\/\*\?\[\]:]/.test(sheetName)) {
      results.push({
        status: 'error',
        message: 'Invalid characters in sheet name',
        details: `The sheet "${sheetName}" contains characters not supported by Excel: \\ / * ? [ ] :`,
        location: sheetName
      });
    }
  }
}

/**
 * Checks cell formats for compatibility
 */
function checkCellFormats(workbook: XLSX.WorkBook, results: CheckResult[]): void {
  // Return early if SheetNames is missing or not an array
  if (!workbook.SheetNames || !Array.isArray(workbook.SheetNames)) {
    return;
  }

  for (const sheetName of workbook.SheetNames) {
    // Check if workbook.Sheets and the specific sheet exist
    if (!workbook.Sheets || !workbook.Sheets[sheetName]) {
      continue;
    }
    
    const worksheet = workbook.Sheets[sheetName];
    
    // Check if worksheet has a reference range defined
    if (!worksheet['!ref']) {
      continue;
    }
    
    try {
      // Check for cells beyond Excel's limits
      const range = XLSX.utils.decode_range(worksheet['!ref']);
      
      if (range.e.r > 1048575) {
        results.push({
          status: 'error',
          message: 'Too many rows',
          details: `Sheet "${sheetName}" contains data beyond row 1,048,576, which is Excel's limit.`,
          location: sheetName
        });
      }
      
      if (range.e.c > 16383) {
        results.push({
          status: 'error',
          message: 'Too many columns',
          details: `Sheet "${sheetName}" contains data beyond column XFD (16,384), which is Excel's limit.`,
          location: sheetName
        });
      }
    } catch (error) {
      // Handle errors when decoding range
      results.push({
        status: 'warning',
        message: 'Unable to check cell limits',
        details: `Could not analyze cell range for sheet "${sheetName}"`,
        location: sheetName
      });
    }
  }
}

/**
 * Checks formula compatibility
 */
function checkFormulas(workbook: XLSX.WorkBook, results: CheckResult[]): void {
  // Return early if SheetNames is missing or not an array
  if (!workbook.SheetNames || !Array.isArray(workbook.SheetNames)) {
    return;
  }

  // Basic formula check
  for (const sheetName of workbook.SheetNames) {
    // Check if workbook.Sheets and the specific sheet exist
    if (!workbook.Sheets || !workbook.Sheets[sheetName]) {
      continue;
    }
    
    const worksheet = workbook.Sheets[sheetName];
    
    // Safely iterate through worksheet cells
    Object.keys(worksheet).forEach(cell => {
      if (cell[0] === '!') return; // Skip special keys
      
      const cellData = worksheet[cell];
      if (cellData && typeof cellData === 'object' && cellData.f) {
        // Check formula length
        if (typeof cellData.f === 'string' && cellData.f.length > 8192) {
          results.push({
            status: 'error',
            message: 'Formula too long',
            details: `Formula in ${sheetName}!${cell} exceeds Excel's maximum length of 8,192 characters.`,
            location: `${sheetName}!${cell}`
          });
        }
      }
    });
  }
}

/**
 * Checks chart compatibility
 */
function checkCharts(workbook: ExcelJS.Workbook, results: CheckResult[]): void {
  let chartCount = 0;
  
  workbook.eachSheet(worksheet => {
    if ((worksheet as any).drawings && Array.isArray((worksheet as any).drawings)) {
      chartCount += (worksheet as any).drawings.length;
    }
  });
  
  if (chartCount > 0) {
    results.push({
      status: 'success',
      message: `Workbook contains ${chartCount} chart(s)`
    });
  }
}

/**
 * Checks macros compatibility
 */
function checkMacros(workbook: XLSX.WorkBook, results: CheckResult[]): void {
  // Check if file contains macros (basic check)
  if (workbook.vbaraw) {
    results.push({
      status: 'warning',
      message: 'Workbook contains VBA macros',
      details: 'VBA macros may have compatibility issues between different Excel versions.'
    });
  }
}

/**
 * Checks named ranges for compatibility
 */
function checkNamedRanges(workbook: XLSX.WorkBook, results: CheckResult[]): void {
  // Basic named range check - not comprehensive without specific Excel workbook properties
  if (workbook.Workbook && 
      workbook.Workbook.Names && 
      Array.isArray(workbook.Workbook.Names) && 
      workbook.Workbook.Names.length > 0) {
    
    const namedRanges = workbook.Workbook.Names;
    
    for (const namedRange of namedRanges) {
      if (namedRange && namedRange.Name && namedRange.Name.length > 255) {
        results.push({
          status: 'error',
          message: 'Named range name too long',
          details: `The named range "${namedRange.Name}" exceeds Excel's 255 character limit.`
        });
      }
    }
    
    results.push({
      status: 'success',
      message: `Found ${namedRanges.length} named range(s)`
    });
  }
}

/**
 * Checks conditional formatting for compatibility issues
 */
function checkConditionalFormatting(workbook: ExcelJS.Workbook, results: CheckResult[]): void {
  let cfCount = 0;
  
  workbook.eachSheet(worksheet => {
    // Use type assertion to access conditionalFormattingRules
    const rules = (worksheet as any).conditionalFormattingRules;
    if (rules && Array.isArray(rules)) {
      cfCount += rules.length;
    }
  });
  
  if (cfCount > 0) {
    if (cfCount > 64000) {
      results.push({
        status: 'error',
        message: 'Too many conditional formatting rules',
        details: `The workbook contains ${cfCount} conditional formatting rules. Excel's limit is 64,000.`
      });
    } else {
      results.push({
        status: 'success',
        message: `Found ${cfCount} conditional formatting rule(s)`
      });
    }
  }
} 