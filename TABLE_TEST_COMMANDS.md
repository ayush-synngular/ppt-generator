# Table Functionality Test Commands

This document lists the test commands that would be used to manually test the new table functionalities.

## 1. Merge Cells Test Command
```bash
node src/index.js create table --rows '[["Name","Department","Salary"],["John","Engineering","100000"],["Jane","Product","120000"]]' --title "Employee Data" --output table-merge-test.pptx
```
This would be the command to test merging cells after implementing the `mergeTableCells` function. Table data should be passed as a JSON string via the `--rows` flag.

## 2. Split Cells Test Command
```bash
node examples/test-table.js
```
This would be extended to include a test case for splitting cells after implementing the `splitTableCell` function.

## 3. Table Themes Test Command
```bash
node examples/test-table.js
```
This would be extended to include a test case for applying table themes after implementing the `applyTableTheme` function.

## 4. Conditional Formatting Test Command
```bash
node examples/test-table.js
```
This would be extended to include a test case for conditional formatting after implementing the `applyConditionalFormatting` function.

## 5. Sort Rows Test Command
```bash
node examples/test-table.js
```
This would be extended to include a test case for sorting rows after implementing the `sortTableRows` function.

## 6. Filter Rows Test Command
```bash
node examples/test-table.js
```
This would be extended to include a test case for filtering rows after implementing the `filterTableRows` function.

## 7. Auto-size Columns Test Command
```bash
node examples/test-table.js
```
This would be extended to include a test case for auto-sizing columns after implementing the `autoSizeTableColumns` function.

## 8. Excel Import Test Command
```bash
node examples/test-table.js
```
This would be extended to include a test case for importing Excel data after implementing the `importExcelTable` function.

## General Testing Approach

To test these functionalities, you would typically:

1. Create a test presentation with a table using the existing `addTable` function
2. Call one of the new table functions on that table
3. Save the presentation and verify the results visually
4. Test edge cases like invalid indices, empty tables, etc.

## Sample Test Structure

For testing purposes, you could create a new test file like `test-table-advanced.js`:

```javascript
const { Presentation } = require('../src/index');
const { addTable } = require('../src/elements/table');

async function testAdvancedTableFeatures() {
  try {
    const ppt = new Presentation('Advanced Table Features Test');
    const slide = ppt.addSlide();
    
    // Create a basic table first
    const rows = [
      ['Name', 'Department', 'Salary'],
      ['John', 'Engineering', '100000'],
      ['Jane', 'Product', '120000'],
      ['Mike', 'Sales', '90000']
    ];
    
    const table = addTable(slide, rows, {
      x: 0.5,
      y: 1.5,
      w: 9,
      h: 3,
      borderColor: '000000',
      borderWidth: 1,
      fontSize: 12,
      color: '000000',
      boldHeader: true
    });
    
    // Test merge cells functionality
    // mergeTableCells(slide, table, 0, 0, 1, 1);
    
    // Test other features here...
    
    await ppt.save('test-advanced-tables.pptx');
    console.log('Advanced table features test saved successfully');
  } catch (error) {
    console.error('Error in advanced table test:', error);
  }
}

testAdvancedTableFeatures();
```

Each new function would be tested individually with appropriate parameters and validation.