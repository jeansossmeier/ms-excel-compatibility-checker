import React, { useState, useEffect } from 'react';

interface FileRecoveryPreviewProps {
  file: File;
}

interface RecoveredContent {
  text: string[];
  tables: SpreadsheetTable[];
  type: 'xml' | 'text' | 'binary';
  source: string;
}

interface SpreadsheetTable {
  name: string;
  data: string[][];
}

export default function FileRecoveryPreview({ file }: FileRecoveryPreviewProps) {
  const [isLoading, setIsLoading] = useState(false);
  const [recoveredContent, setRecoveredContent] = useState<RecoveredContent | null>(null);
  const [showPreview, setShowPreview] = useState(false);
  
  useEffect(() => {
    if (file && showPreview && !recoveredContent) {
      attemptContentRecovery();
    }
  }, [file, showPreview]);
  
  const attemptContentRecovery = async () => {
    setIsLoading(true);
    
    try {
      // Read file as ArrayBuffer
      const buffer = await file.arrayBuffer();
      const bytes = new Uint8Array(buffer);
      
      // Try to locate XML content
      let content: RecoveredContent | null = null;
      
      // Check if it's XLSX (ZIP) format
      const isZip = bytes.length >= 4 && 
                    bytes[0] === 0x50 && bytes[1] === 0x4B && 
                    bytes[2] === 0x03 && bytes[3] === 0x04;
                    
      if (isZip) {
        // It's a ZIP, try to extract readable text content
        content = await extractTextFromZip(bytes);
      } else {
        // Try to recover any textual content
        content = await extractTextFromBinary(bytes);
      }
      
      setRecoveredContent(content);
    } catch (error) {
      console.error("Error recovering content:", error);
    } finally {
      setIsLoading(false);
    }
  };
  
  const extractTextFromZip = async (bytes: Uint8Array): Promise<RecoveredContent> => {
    // Try to find XML fragments in the ZIP
    const decoder = new TextDecoder();
    const fullText = decoder.decode(bytes);
    
    // Extract document properties
    const docProps: string[] = [];
    
    // Creator
    const creatorMatch = /<dc:creator>([^<]+)<\/dc:creator>/;
    const creatorResult = creatorMatch.exec(fullText);
    if (creatorResult && creatorResult[1]) {
      docProps.push(`Creator: ${creatorResult[1]}`);
    }
    
    // Last modified by
    const modifiedByMatch = /<cp:lastModifiedBy>([^<]+)<\/cp:lastModifiedBy>/;
    const modifiedByResult = modifiedByMatch.exec(fullText);
    if (modifiedByResult && modifiedByResult[1]) {
      docProps.push(`Last modified by: ${modifiedByResult[1]}`);
    }
    
    // Title
    const titleMatch = /<dc:title>([^<]+)<\/dc:title>/;
    const titleResult = titleMatch.exec(fullText);
    if (titleResult && titleResult[1]) {
      docProps.push(`Title: ${titleResult[1]}`);
    }

    // Extract worksheet names
    const worksheetNames: string[] = [];
    const worksheetRegex = /<sheet name="([^"]+)"/g;
    let worksheetMatch;
    const sheetIds: {[key: string]: string} = {};
    
    while ((worksheetMatch = worksheetRegex.exec(fullText)) !== null) {
      if (worksheetMatch[1]) {
        const sheetName = worksheetMatch[1];
        worksheetNames.push(`Sheet: ${sheetName}`);
        
        // Try to get sheet ID
        const idMatch = new RegExp(`<sheet[^>]*name="${sheetName}"[^>]*r:id="([^"]+)"`, 'i');
        const idResult = idMatch.exec(fullText);
        if (idResult && idResult[1]) {
          sheetIds[sheetName] = idResult[1];
        }
      }
    }
    
    // Extract Tables
    const tables: SpreadsheetTable[] = [];
    
    // First attempt: Try to find intact Sheet XML and extract row/cell data
    const sheetDataRegex = /<sheetData>([\s\S]*?)<\/sheetData>/g;
    let sheetDataMatch;
    let tableCount = 0;
    
    while ((sheetDataMatch = sheetDataRegex.exec(fullText)) !== null && tableCount < 3) {
      if (sheetDataMatch[1]) {
        const sheetData = sheetDataMatch[1];
        const tableData: string[][] = [];
        
        // Extract rows
        const rowRegex = /<row[^>]*>([\s\S]*?)<\/row>/g;
        let rowMatch;
        
        while ((rowMatch = rowRegex.exec(sheetData)) !== null) {
          if (rowMatch[1]) {
            const rowContent = rowMatch[1];
            const rowData: string[] = [];
            
            // Extract cells
            const cellRegex = /<c[^>]*><v>([^<]*)<\/v><\/c>/g;
            let cellMatch;
            let hasContent = false;
            
            while ((cellMatch = cellRegex.exec(rowContent)) !== null) {
              rowData.push(cellMatch[1] || "");
              if (cellMatch[1] && cellMatch[1].trim().length > 0) {
                hasContent = true;
              }
            }
            
            // Only add rows that have content
            if (hasContent && rowData.length > 0) {
              tableData.push(rowData);
            }
          }
        }
        
        if (tableData.length > 0) {
          tableCount++;
          tables.push({
            name: `Sheet ${tableCount}`,
            data: tableData
          });
        }
      }
    }
    
    // Second attempt: Try to extract table data from shared strings and cell references
    if (tables.length === 0) {
      // Build a map of shared strings
      const sharedStrings: string[] = [];
      const stringTableRegex = /<sst[^>]*>([\s\S]*?)<\/sst>/;
      const stringTableMatch = stringTableRegex.exec(fullText);
      
      if (stringTableMatch && stringTableMatch[1]) {
        const stringTable = stringTableMatch[1];
        const stringRegex = /<si>[\s\S]*?<t[^>]*>([^<]*)<\/t>[\s\S]*?<\/si>/g;
        let stringMatch;
        
        while ((stringMatch = stringRegex.exec(stringTable)) !== null) {
          if (stringMatch[1]) {
            sharedStrings.push(stringMatch[1]);
          }
        }
        
        // Now try to construct a table from cell references and shared strings
        // This is a simplified approach - real Excel files use more complex addressing
        const cellRefRegex = /<c r="([A-Z]+)(\d+)"[^>]*t="s"><v>(\d+)<\/v><\/c>/g;
        let cellRefMatch;
        
        const cellData: {[key: string]: {row: number, col: string, value: string}} = {};
        
        while ((cellRefMatch = cellRefRegex.exec(fullText)) !== null) {
          const colRef = cellRefMatch[1];
          const rowRef = parseInt(cellRefMatch[2]);
          const stringIndex = parseInt(cellRefMatch[3]);
          
          if (!isNaN(stringIndex) && stringIndex < sharedStrings.length) {
            const cellValue = sharedStrings[stringIndex];
            cellData[`${colRef}${rowRef}`] = { 
              row: rowRef, 
              col: colRef, 
              value: cellValue 
            };
          }
        }
        
        // Also look for inline string values
        const inlineStringRegex = /<c r="([A-Z]+)(\d+)"[^>]*><is><t>([^<]*)<\/t><\/is><\/c>/g;
        let inlineStringMatch;
        
        while ((inlineStringMatch = inlineStringRegex.exec(fullText)) !== null) {
          const colRef = inlineStringMatch[1];
          const rowRef = parseInt(inlineStringMatch[2]);
          const cellValue = inlineStringMatch[3];
          
          cellData[`${colRef}${rowRef}`] = { 
            row: rowRef, 
            col: colRef, 
            value: cellValue 
          };
        }
        
        // Look for numeric values
        const numericRegex = /<c r="([A-Z]+)(\d+)"[^>]*><v>([^<]*)<\/v><\/c>/g;
        let numericMatch;
        
        while ((numericMatch = numericRegex.exec(fullText)) !== null) {
          const colRef = numericMatch[1];
          const rowRef = parseInt(numericMatch[2]);
          const cellValue = numericMatch[3];
          
          cellData[`${colRef}${rowRef}`] = { 
            row: rowRef, 
            col: colRef, 
            value: cellValue 
          };
        }
        
        // Organize cells into a table
        if (Object.keys(cellData).length > 0) {
          // Get unique row and column references
          const rows = Array.from(new Set(Object.values(cellData).map(cell => cell.row)))
            .sort((a, b) => a - b);
          const cols = Array.from(new Set(Object.values(cellData).map(cell => cell.col)))
            .sort();
          
          if (rows.length > 0 && cols.length > 0) {
            const tableData: string[][] = [];
            
            // Create header row with column letters
            tableData.push([''].concat(cols));
            
            // Create data rows
            for (const row of rows) {
              const rowData = [row.toString()]; // Row number as first cell
              
              for (const col of cols) {
                const cellKey = `${col}${row}`;
                rowData.push(cellData[cellKey]?.value || '');
              }
              
              tableData.push(rowData);
            }
            
            tables.push({
              name: 'Recovered Spreadsheet',
              data: tableData
            });
          }
        }
      }
    }
    
    // If we have tables, use them
    if (tables.length > 0) {
      return {
        text: [...docProps, ...worksheetNames],
        tables: tables,
        type: 'xml',
        source: 'Spreadsheet content recovered from Excel file'
      };
    }

    // Look for shared strings (text content)
    const stringValues: string[] = [];
    const stringRegex = /<t[^>]*>([^<]+)<\/t>/g;
    let stringMatch;
    while ((stringMatch = stringRegex.exec(fullText)) !== null) {
      if (stringMatch[1] && stringMatch[1].trim().length > 0) {
        stringValues.push(stringMatch[1].trim());
      }
    }
    
    // Extract cell values from sheet data
    const cellValues: string[] = [];
    const cellRegex = /<c[^>]*><v>([^<]+)<\/v><\/c>/g;
    let cellMatch;
    while ((cellMatch = cellRegex.exec(fullText)) !== null) {
      if (cellMatch[1] && cellMatch[1].length > 0) {
        // Only add if it looks like meaningful text (not just numbers)
        if (cellMatch[1].length > 3 || /[a-zA-Z]/.test(cellMatch[1])) {
          cellValues.push(cellMatch[1]);
        }
      }
    }
    
    // Combine all meaningful content
    const allContent: string[] = [
      ...docProps,
      ...worksheetNames,
      ...stringValues.filter(str => 
        // Filter out strings that look like XML or have no alphabetic characters
        str.length > 1 && 
        !str.includes('<') && 
        !str.includes('>') &&
        /[a-zA-Z]/.test(str) &&
        !/^[0-9.,-]+$/.test(str) // Not just numbers and punctuation
      ).slice(0, 50), // Limit to 50 strings
      ...cellValues.filter(val => 
        val.length > 1 && 
        /[a-zA-Z]/.test(val) &&
        !/^[0-9.,-]+$/.test(val)
      ).slice(0, 20) // Limit to 20 cell values
    ];
    
    // Try to create a simple table from the text content
    if (stringValues.length > 0 || cellValues.length > 0) {
      const tableData: string[][] = [];
      
      // Create chunks of related content
      const chunks: string[][] = [];
      let currentChunk: string[] = [];
      let lastWasEmpty = false;
      
      [...stringValues, ...cellValues].forEach(value => {
        if (value.trim().length === 0) {
          if (!lastWasEmpty) {
            lastWasEmpty = true;
            if (currentChunk.length > 0) {
              chunks.push([...currentChunk]);
              currentChunk = [];
            }
          }
        } else {
          lastWasEmpty = false;
          currentChunk.push(value);
          if (currentChunk.length >= 10) {
            chunks.push([...currentChunk]);
            currentChunk = [];
          }
        }
      });
      
      if (currentChunk.length > 0) {
        chunks.push(currentChunk);
      }
      
      // Use the largest chunk for a table
      if (chunks.length > 0) {
        const largestChunk = chunks.reduce((prev, current) => 
          current.length > prev.length ? current : prev, chunks[0]);
          
        if (largestChunk.length >= 3) {
          // Simple table with one column
          tableData.push(['Value']);
          largestChunk.forEach(value => {
            tableData.push([value]);
          });
          
          tables.push({
            name: 'Recovered Data',
            data: tableData
          });
        }
      }
    }
    
    return {
      text: allContent.length > 0 ? allContent : ['No readable text content found'],
      tables: tables,
      type: 'xml',
      source: 'Content extracted from Excel file'
    };
  };
  
  const extractTextFromBinary = async (bytes: Uint8Array): Promise<RecoveredContent> => {
    // Try to extract text strings from binary data
    const decoder = new TextDecoder('utf-8', { fatal: false });
    
    // Find continuous blocks of text (at least 4 ASCII chars in a row)
    const textBlocks: string[] = [];
    let currentBlock: number[] = [];
    
    for (let i = 0; i < bytes.length; i++) {
      const byte = bytes[i];
      // Check if it's a printable ASCII character
      if ((byte >= 32 && byte <= 126) || byte === 9 || byte === 10 || byte === 13) {
        currentBlock.push(byte);
      } else {
        // End of a potential text block
        if (currentBlock.length >= 4) {
          const blockText = decoder.decode(new Uint8Array(currentBlock)).trim();
          // Only keep if it contains letters and is not just numbers or symbols
          if (/[a-zA-Z]{2,}/.test(blockText) && !/^[0-9.,-]+$/.test(blockText)) {
            textBlocks.push(blockText);
          }
        }
        currentBlock = [];
      }
    }
    
    // Handle any remaining block
    if (currentBlock.length >= 4) {
      const blockText = decoder.decode(new Uint8Array(currentBlock)).trim();
      if (/[a-zA-Z]{2,}/.test(blockText) && !/^[0-9.,-]+$/.test(blockText)) {
        textBlocks.push(blockText);
      }
    }
    
    // Filter and clean text blocks
    const cleanedBlocks = textBlocks
      .filter(block => 
        // Must contain actual words
        block.length >= 3 && 
        /[a-zA-Z]{3,}/.test(block) &&
        // No binary garbage
        !/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/.test(block) &&
        // Reasonable text-to-nontext ratio
        block.replace(/[^a-zA-Z0-9.,;:'"!? -]/g, '').length > block.length * 0.7
      )
      // Remove duplicates
      .filter((block, index, self) => self.indexOf(block) === index)
      .slice(0, 40); // Limit to 40 blocks
    
    // Try to create a table from the text content
    const tables: SpreadsheetTable[] = [];
    
    // Look for table-like structures in the text
    const tableLines = cleanedBlocks.filter(block => 
      // Potential table rows often have similar patterns
      (block.includes('\t') || block.includes(',')) &&
      // Has multiple "cells"
      (block.split('\t').length > 1 || block.split(',').length > 1)
    );
    
    if (tableLines.length >= 3) {
      // Detect the delimiter
      const delimiter = tableLines[0].includes('\t') ? '\t' : ',';
      
      // Create a table
      const tableData: string[][] = [];
      
      tableLines.forEach(line => {
        tableData.push(line.split(delimiter).map(cell => cell.trim()));
      });
      
      tables.push({
        name: 'Recovered Table',
        data: tableData
      });
    }
      
    if (cleanedBlocks.length > 0) {
      return {
        text: cleanedBlocks,
        tables,
        type: 'binary',
        source: 'Text recovered from binary Excel file'
      };
    }
    
    return {
      text: ['No readable text could be recovered from this file.'],
      tables: [],
      type: 'binary',
      source: 'Binary file analysis'
    };
  };
  
  const togglePreview = () => {
    setShowPreview(!showPreview);
  };
  
  return (
    <div className="file-recovery-preview">
      <button 
        className="preview-toggle"
        onClick={togglePreview}
        disabled={isLoading}
      >
        {isLoading ? 'Analyzing file...' : 
          (showPreview ? 'Hide Content Preview' : 'Attempt Content Recovery')}
      </button>
      
      {showPreview && (
        <div className="preview-content">
          {isLoading ? (
            <div className="loading">
              <div className="spinner"></div>
              <p>Attempting to recover content...</p>
            </div>
          ) : recoveredContent ? (
            <>
              <h3>Recovered Content Preview</h3>
              <p className="source-info">{recoveredContent.source}</p>
              
              {/* Display tables first if available */}
              {recoveredContent.tables && recoveredContent.tables.length > 0 && (
                <div className="spreadsheet-tables">
                  {recoveredContent.tables.map((table, tableIndex) => (
                    <div key={tableIndex} className="table-container">
                      <h4>{table.name}</h4>
                      <div className="spreadsheet-table">
                        <table>
                          <tbody>
                            {table.data.map((row, rowIndex) => (
                              <tr key={rowIndex}>
                                {row.map((cell, cellIndex) => (
                                  <td key={cellIndex}>{cell}</td>
                                ))}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  ))}
                </div>
              )}
              
              {/* Then show text content if available */}
              {recoveredContent.text.length > 0 && recoveredContent.text[0] !== 'No readable text could be recovered from this file.' && (
                <div className="text-preview">
                  <h4>Additional Recovered Text</h4>
                  {recoveredContent.text.map((line, index) => (
                    <div key={index} className="preview-line">
                      {line}
                    </div>
                  ))}
                </div>
              )}
              
              {recoveredContent.tables.length === 0 && 
               (recoveredContent.text.length === 0 || 
                recoveredContent.text[0] === 'No readable text could be recovered from this file.') && (
                <p className="no-content">No spreadsheet content could be recovered from this file.</p>
              )}
              
              <p className="recovery-note">
                This is just a preview of recoverable content. Not all data could be recovered,
                and some recovered content may be incomplete or out of context.
              </p>
            </>
          ) : (
            <p className="no-content">Unable to recover content from this file.</p>
          )}
        </div>
      )}
    </div>
  );
} 