# Table Functionality Documentation

This document lists all the new table functionalities that need to be implemented for the PowerPoint generator.

## 1. Merge Cells
### Function: `mergeTableCells(slide, table, startRow, startCol, endRow, endCol)`
- **Purpose**: Merges cells in a table to create a larger cell spanning multiple rows/columns
- **Parameters**:
  - `slide`: The slide object containing the table
  - `table`: The table object to modify
  - `startRow`: Starting row index (0-based) (passed via terminal arguments)
  - `startCol`: Starting column index (0-based) (passed via terminal arguments)
  - `endRow`: Ending row index (0-based) (passed via terminal arguments)
  - `endCol`: Ending column index (0-based) (passed via terminal arguments)
- **Returns**: Modified table object or boolean indicating success

## 2. Split Cells
### Function: `splitTableCell(slide, table, row, col)`
- **Purpose**: Splits a merged cell back into individual cells
- **Parameters**:
  - `slide`: The slide object containing the table
  - `table`: The table object to modify
  - `row`: Row index of the cell to split (0-based)
  - `col`: Column index of the cell to split (0-based)
- **Returns**: Modified table object or boolean indicating success

## 3. Table Themes
### Function: `applyTableTheme(slide, table, themeName)`
- **Purpose**: Applies predefined themes to tables for consistent styling
- **Parameters**:
  - `slide`: The slide object containing the table
  - `table`: The table object to style
  - `themeName`: Name of the theme to apply (e.g., 'modern', 'classic', 'professional')
- **Returns**: Modified table object or boolean indicating success

## 4. Conditional Formatting
### Function: `applyConditionalFormatting(slide, table, rule)`
- **Purpose**: Applies conditional formatting rules to table cells based on content or values
- **Parameters**:
  - `slide`: The slide object containing the table
  - `table`: The table object to format
  - `rule`: Object defining the formatting rule with properties:
    - `column`: Column index to apply rule to
    - `condition`: Comparison condition (e.g., 'greaterThan', 'lessThan', 'equals')
    - `value`: Value to compare against
    - `format`: Formatting options to apply when condition is met
- **Returns**: Modified table object or boolean indicating success

## 5. Sort Rows
### Function: `sortTableRows(slide, table, columnIndex, ascending = true)`
- **Purpose**: Sorts table rows based on values in a specified column
- **Parameters**:
  - `slide`: The slide object containing the table
  - `table`: The table object to sort
  - `columnIndex`: Index of the column to sort by (0-based)
  - `ascending`: Boolean indicating sort order (default: true)
- **Returns**: Modified table object or boolean indicating success

## 6. Filter Rows
### Function: `filterTableRows(slide, table, filterFunction)`
- **Purpose**: Filters table rows based on a custom filtering function
- **Parameters**:
  - `slide`: The slide object containing the table
  - `table`: The table object to filter
  - `filterFunction`: Function that takes a row and returns true/false for inclusion
- **Returns**: Modified table object or boolean indicating success

## 7. Auto-size Columns
### Function: `autoSizeTableColumns(slide, table)`
- **Purpose**: Automatically sizes table columns to fit their content
- **Parameters**:
  - `slide`: The slide object containing the table
  - `table`: The table object to resize
- **Returns**: Modified table object or boolean indicating success

## 8. Excel Import
### Function: `importExcelTable(slide, excelFilePath, sheetName, range)`
- **Purpose**: Imports data from Excel file and creates a table in PowerPoint
- **Parameters**:
  - `slide`: The slide object to add the table to
  - `excelFilePath`: Path to the Excel (.xlsx) file
  - `sheetName`: Name of the worksheet to import from
  - `range`: Cell range to import (e.g., 'A1:D10')
- **Returns**: Table object created from Excel data or boolean indicating success

## Implementation Notes

All new table functions should:
1. Be designed to work with existing table objects created by `addTable()`
2. Not modify or interfere with existing table functionality
3. Handle edge cases gracefully without throwing exceptions
4. Follow consistent parameter structures and naming conventions
5. Integrate well with the PowerPoint generation system
6. Be backward compatible with existing code

Note: These functions are designed to extend the existing table functionality and should be added to the table.js file or a new dedicated table utilities file, not to modify the existing `addTable` function itself.