# Excel Compatibility Checker

A web application that analyzes Microsoft Excel spreadsheets to ensure 100% compatibility with Microsoft Excel formats and limitations.

## Features

- **Drag & Drop Interface**: Easily upload your Excel files
- **Comprehensive Checks**: Verifies compatibility with numerous Excel constraints
- **Detailed Reports**: Get a clear breakdown of issues found, with errors, warnings and passed checks

## Excel Compatibility Checks

The application performs the following compatibility checks:

- **File Format**: Validates Excel file format and extension
- **File Size**: Checks if the file size exceeds Excel's recommended limits
- **Worksheet Structure**: Verifies worksheet count and naming conventions
- **Cell Limits**: Ensures data doesn't exceed Excel's row/column limits
- **Formula Validation**: Checks formula length and complexity
- **Macros**: Identifies potential issues with VBA macros
- **Named Ranges**: Validates named ranges against Excel's limitations
- **Conditional Formatting**: Checks if conditional formatting rules exceed limits
- **Charts**: Basic detection of chart elements

## Getting Started

### Prerequisites

- Node.js (v14 or higher)
- npm or yarn

### Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/ms-excel-compatibility-checker.git
cd ms-excel-compatibility-checker
```

2. Install dependencies:
```bash
npm install
```

3. Start the development server:
```bash
npm run dev
```

4. Open your browser and navigate to `http://localhost:5173`

## Building for Production

```bash
npm run build
```

The build files will be generated in the `dist` directory.

## Technical Details

This project is built using:

- **React**: For the user interface
- **TypeScript**: For type safety
- **Vite**: For fast development and optimized builds
- **SheetJS**: For parsing Excel files
- **ExcelJS**: For detailed Excel feature analysis

## Contributing

Contributions are welcome! Here are some ways you can contribute:

- Report bugs
- Suggest new features or compatibility checks
- Submit pull requests
- Improve documentation

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgements

- [SheetJS](https://sheetjs.com/) - For Excel file parsing functionality
- [ExcelJS](https://github.com/exceljs/exceljs) - For detailed Excel structure analysis
