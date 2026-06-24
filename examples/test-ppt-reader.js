'use strict';

const fs = require('fs');
const path = require('path');
const PptReader = require('../src/import/PptReader');

async function testPptReader() {
  const reader = new PptReader();
  let filePath = process.argv[2] || path.resolve(__dirname, 'test-auto-layout.pptx');

  if (!fs.existsSync(filePath)) {
    const alternate = path.resolve(process.cwd(), 'test-auto-layout.pptx');
    if (fs.existsSync(alternate)) {
      filePath = alternate;
    }
  }

  if (!fs.existsSync(filePath)) {
    console.error(`PPTX file not found: ${filePath}`);
    process.exit(1);
  }

  const presentation = await reader.read(filePath);
  const result = presentation.toJSON();

  fs.writeFileSync(path.resolve(__dirname, 'ppt-structure.json'), JSON.stringify(result, null, 2), 'utf8');

  console.log(`Slide count: ${presentation.getSlideCount()}`);
  presentation.slides.forEach((slide) => {
    console.log(`Slide ${slide.slideNumber}:`);
    if (slide.text.length === 0) {
      console.log('- (no visible text)');
    } else {
      slide.text.forEach((line) => {
        console.log(`- ${line}`);
      });
    }
  });
  console.log('Wrote ppt-structure.json');
}

testPptReader();
