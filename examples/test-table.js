'use strict';

const { Presentation } = require('../src/index');
const { addTable } = require('../src/elements/table');

async function testTable() {
  try {
    const ppt = new Presentation('Table Test Presentation');
    const slide = ppt.addSlide();

    slide.addText('Table Test', {
      x: 0.5,
      y: 0.5,
      w: 8,
      h: 1,
      fontFace: 'Arial',
      fontSize: 28,
      bold: true,
      align: 'center'
    });

    const rows = [
      ['Name', 'Department', 'Salary'],
      ['John', 'Engineering', '100000'],
      ['Jane', 'Product', '120000'],
      ['Mike', 'Sales', '90000']
    ];

    addTable(slide, rows, {
      x: 0.5,
      y: 1.5,
      w: 9,
      h: 3,
      borderColor: '000000',
      borderWidth: 1,
      fontSize: 12,
      color: '000000',
      fillColor: 'F2F2F2',
      boldHeader: true,
      autoFit: true
    });

    await ppt.save('test-table.pptx');
    console.log('Presentation saved successfully as test-table.pptx');
  } catch (error) {
    console.error('Error creating table presentation:', error);
  }
}

// Run the test
testTable();