'use strict';

const fs = require('fs');
const path = require('path');
const PptReader = require('../src/import/PptReader');
const PptModifier = require('../src/import/PptModifier');

async function testPptModifier() {
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
  const modifier = new PptModifier(presentation);

  modifier.replaceText('Content Area', 'Modified Content Area');

  modifier.getSlides().forEach((slide) => {
    console.log(`Slide ${slide.slideNumber}:`);
    if (!Array.isArray(slide.text) || slide.text.length === 0) {
      console.log('- (no visible text)');
    } else {
      slide.text.forEach((line) => {
        console.log(`- ${line}`);
      });
    }
    console.log('');
  });

  const outputPath = path.resolve(__dirname, 'modified-presentation.json');
  modifier.saveJSON(outputPath);
  console.log(`Wrote ${path.basename(outputPath)}`);
}

testPptModifier();
