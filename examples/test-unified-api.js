'use strict';

const path = require('path');
const { Presentation } = require('../src');

async function testUnifiedApi() {
  let inputPath = path.resolve(process.cwd(), 'test-auto-layout.pptx');
  if (!require('fs').existsSync(inputPath)) {
    inputPath = path.resolve(__dirname, 'test-auto-layout.pptx');
  }

  const outputPath = path.resolve(process.cwd(), 'unified-output.pptx');

  const ppt = await Presentation.open(inputPath);
  ppt.replaceText('Content Area', 'Updated Content Area');
  
  const slide1 = ppt.slide(1);
  console.log(`Slide 1 text count: ${slide1 && slide1.text ? slide1.text.length : 0}`);
  
  await ppt.save(outputPath);

  console.log('unified-output.pptx created');
  console.log(`Slide count preserved: ${ppt._importModel ? ppt._importModel.slides.length : 0}`);
}

testUnifiedApi();
