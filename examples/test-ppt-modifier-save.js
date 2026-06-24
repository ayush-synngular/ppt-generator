'use strict';

const fs = require('fs');
const path = require('path');
const PptReader = require('../src/import/PptReader');
const PptModifier = require('../src/import/PptModifier');

async function testPptModifierSave() {
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

  console.log(`Original slides: ${presentation.getSlideCount()}`);
  console.log(`Modified slides: ${modifier.getSlides().length}`);

  const outputFileName = 'modified-presentation.pptx';
  const outputPath = path.resolve(process.cwd(), outputFileName);
  await modifier.save(outputPath);

  console.log(`Saved PPTX: ${outputFileName}`);
}

testPptModifierSave();
