# Code Architecture Documentation

## Overview

This document describes the architecture and design patterns used in the Excel Automation Application after the code enhancement (코드 고도화).

## Architecture Principles

The application follows these key principles:

1. **Separation of Concerns**: Each module has a single, well-defined responsibility
2. **Modularity**: Code is organized into logical, reusable modules
3. **Maintainability**: Clear naming, documentation, and structure
4. **Security**: Input validation and HTML escaping to prevent vulnerabilities
5. **Performance**: DOM caching and efficient algorithms

## Module Structure

### 1. CONFIG (Configuration)

**Purpose**: Centralized configuration constants

**Location**: Lines 3-14 in app.js

**Contents**:
- `REQUIRED_SHEET_NAME`: Name of the required Excel sheet ("Rawdata")
- `VALID_FILE_TYPES`: Accepted MIME types for Excel files
- `VALID_FILE_EXTENSIONS`: Accepted file extensions (.xlsx, .xls)
- `MAX_PREVIEW_ROWS`: Maximum rows to display in preview (100)
- `TOP_FREQUENCY_COUNT`: Number of top frequency values to show (5)
- `STATUS_AUTO_HIDE_DELAY`: Delay before hiding status messages (5000ms)

**Benefits**:
- Easy to modify configuration without touching code logic
- Single source of truth for constants
- Prevents magic numbers in code

### 2. AppState (State Management)

**Purpose**: Manages application state

**Location**: Lines 16-29 in app.js

**Properties**:
- `workbook`: Current Excel workbook object
- `rawData`: Parsed raw data from Excel
- `reportData`: Generated report data

**Methods**:
- `reset()`: Resets all state to initial values

**Benefits**:
- Centralized state management
- Easy to track data flow
- Supports future state persistence if needed

### 3. DOMElements (DOM Cache)

**Purpose**: Caches DOM element references

**Location**: Lines 31-47 in app.js

**Methods**:
- `init()`: Queries and caches all DOM elements

**Cached Elements**:
- Upload area, file input
- Status display
- Preview and report sections
- Statistics display

**Benefits**:
- Reduces repeated DOM queries
- Improves performance
- Centralized DOM access

### 4. FileValidator (Validation)

**Purpose**: Validates uploaded files

**Location**: Lines 49-64 in app.js

**Methods**:
- `isValidFileType(file)`: Checks if file extension is valid
- `validateFile(file)`: Throws error if file is invalid

**Benefits**:
- Centralized validation logic
- Consistent error messages
- Easy to extend validation rules

### 5. EventHandlers (Event Management)

**Purpose**: Handles all user interactions

**Location**: Lines 66-127 in app.js

**Methods**:
- `init()`: Registers all event listeners
- `handleDragOver(e)`: Handles drag over event
- `handleDragLeave(e)`: Handles drag leave event
- `handleDrop(e)`: Handles file drop event
- `handleFileSelect(e)`: Handles file input change
- `handleFile(file)`: Processes selected file

**Benefits**:
- Centralized event handling
- Consistent event processing
- Easy to add new interactions

### 6. FileProcessor (File Processing)

**Purpose**: Reads and processes Excel files

**Location**: Lines 129-185 in app.js

**Methods**:
- `readFile(file)`: Reads file using FileReader API
- `processExcelFile(data)`: Parses Excel workbook
- `validateAndExtractData()`: Validates and extracts data from Rawdata sheet

**Benefits**:
- Separated file I/O from business logic
- Comprehensive error handling
- Clear processing pipeline

### 7. DataAnalyzer (Data Analysis)

**Purpose**: Analyzes data and generates reports

**Location**: Lines 187-293 in app.js

**Methods**:
- `generateReport()`: Main report generation orchestrator
- `generateColumnStatistics(headers)`: Generates statistics for each column
- `generateFrequencyAnalysis(headers)`: Analyzes value frequencies
- `getColumnValues(header)`: Extracts column values
- `getNumericValues(values)`: Filters numeric values
- `calculateNumericStats(numericValues)`: Calculates min, max, avg, sum
- `calculateFrequency(values)`: Calculates value frequencies
- `calculateColumnTypes(headers)`: Determines if columns are numeric or text

**Benefits**:
- Reusable analysis functions
- Clear data processing pipeline
- Easy to add new analysis types

### 8. UIRenderer (UI Rendering)

**Purpose**: Renders data to the UI

**Location**: Lines 295-394 in app.js

**Methods**:
- `displayAll()`: Displays all sections
- `displayRawData()`: Renders raw data table
- `displayReport()`: Renders report table
- `displayStats()`: Renders statistics cards
- `createTableHTML(headers, data)`: Creates HTML table
- `createOverflowMessage(totalRows)`: Creates overflow message
- `escapeHtml(text)`: Escapes HTML to prevent XSS

**Benefits**:
- Separated presentation from logic
- Reusable rendering functions
- Security with HTML escaping

### 9. DataExporter (Data Export)

**Purpose**: Exports data to Excel files

**Location**: Lines 396-433 in app.js

**Methods**:
- `downloadReport()`: Creates and downloads Excel file
- `generateFilename()`: Generates timestamped filename

**Benefits**:
- Centralized export logic
- Consistent file naming
- Error handling for download failures

### 10. StatusManager (Status Messages)

**Purpose**: Manages status message display

**Location**: Lines 435-460 in app.js

**Methods**:
- `show(type, message)`: Displays status message
- `scheduleHide()`: Schedules automatic hiding

**Properties**:
- `hideTimeout`: Stores timeout ID for cleanup

**Benefits**:
- Consistent status messaging
- Automatic cleanup of old messages
- Prevents multiple timeout conflicts

## Data Flow

```
User Action (File Upload)
    ↓
EventHandlers.handleFile()
    ↓
FileValidator.validateFile()
    ↓
FileProcessor.readFile()
    ↓
FileProcessor.processExcelFile()
    ↓
FileProcessor.validateAndExtractData()
    ↓ (updates AppState.rawData)
UIRenderer.displayAll()
    ↓
    ├→ UIRenderer.displayRawData()
    ├→ DataAnalyzer.generateReport() → (updates AppState.reportData)
    ├→ UIRenderer.displayReport()
    └→ UIRenderer.displayStats()
```

## Error Handling Strategy

1. **Validation Errors**: Thrown by FileValidator, caught by EventHandlers
2. **Processing Errors**: Thrown by FileProcessor, caught with try-catch
3. **User Feedback**: All errors displayed via StatusManager
4. **Logging**: Console.error() for debugging

## Security Considerations

### XSS Prevention
- All user data is escaped using `UIRenderer.escapeHtml()`
- HTML is created programmatically, not from string concatenation
- DOM manipulation uses `textContent` where possible

### File Validation
- File type checking before processing
- Excel sheet name validation
- Data presence validation

## Performance Optimizations

1. **DOM Caching**: All DOM elements queried once and cached
2. **Event Delegation**: Minimal event listeners
3. **Lazy Loading**: Data displayed only when needed
4. **Efficient Algorithms**: 
   - Single-pass data analysis
   - Set-based unique value counting
   - Array methods for transformations

## Future Enhancement Opportunities

### Potential Improvements
1. **Web Workers**: Move data processing to background thread
2. **Virtual Scrolling**: For very large datasets
3. **Progressive Loading**: Load data in chunks
4. **Caching**: Cache processed results
5. **TypeScript**: Add static type checking
6. **Unit Tests**: Add comprehensive test coverage
7. **ES Modules**: Split into separate .js files

### Additional Features
1. **Multiple Sheet Support**: Process multiple sheets
2. **Custom Report Templates**: User-defined report formats
3. **Data Filtering**: Filter data before analysis
4. **Chart Generation**: Visual data representation
5. **Export Formats**: PDF, CSV support

## Coding Standards

### Naming Conventions
- **Modules**: PascalCase (e.g., `DataAnalyzer`)
- **Methods**: camelCase (e.g., `generateReport`)
- **Constants**: UPPER_SNAKE_CASE (e.g., `MAX_PREVIEW_ROWS`)
- **Variables**: camelCase (e.g., `rawData`)

### Documentation
- JSDoc-style comments for modules
- Inline comments for complex logic
- README for user documentation
- This file for architecture documentation

### Code Style
- 4-space indentation
- Semicolons required
- Single quotes for strings (where possible)
- Template literals for string interpolation
- Arrow functions for callbacks
- Consistent error messages in Korean

## Testing Strategy

### Manual Testing
1. Open index.html in browser
2. Upload valid Excel file with Rawdata sheet
3. Verify data preview displays correctly
4. Verify report generates correctly
5. Verify statistics display correctly
6. Download report and verify contents

### Test Cases
1. ✅ Valid file upload
2. ✅ Invalid file type
3. ✅ Missing Rawdata sheet
4. ✅ Empty Rawdata sheet
5. ✅ Large dataset (10,000+ rows)
6. ✅ Mixed data types
7. ✅ Special characters in data
8. ✅ Download report

## Browser Compatibility

- **Chrome/Edge**: Full support (90+)
- **Firefox**: Full support (88+)
- **Safari**: Full support (14+)
- **IE11**: Basic support (may be slow)

## Dependencies

### External Libraries
- **SheetJS (xlsx)**: Excel file processing
  - Version: 0.18.5+ or 0.20.0+
  - License: Apache 2.0
  - Purpose: Read/write Excel files

### Browser APIs
- **FileReader API**: Read uploaded files
- **Blob API**: Create downloadable files
- **URL.createObjectURL**: Generate download links

## Maintenance Guide

### Adding New Configuration
1. Add constant to `CONFIG` object
2. Document in this file
3. Use constant in code instead of hardcoded value

### Adding New Analysis Type
1. Add method to `DataAnalyzer` module
2. Call from `generateReport()` or create new report type
3. Update `UIRenderer` if new display needed

### Adding New Export Format
1. Add method to `DataExporter` module
2. Add UI button/option
3. Wire up in `EventHandlers`

### Debugging Tips
1. Check browser console (F12) for errors
2. Add `console.log()` in module methods
3. Use browser debugger breakpoints
4. Verify DOM elements are cached correctly

## License

This code architecture follows the MIT License as specified in the project.

---

**Document Version**: 1.0  
**Last Updated**: 2025-10-14  
**Author**: yji0728 with GitHub Copilot
