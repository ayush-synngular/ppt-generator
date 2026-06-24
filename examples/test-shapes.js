'use strict';

const { Presentation } = require('../src/index');
const { addRectangle, addCircle, addLine, addArrow } = require('../src/elements/shape');

async function testShapes() {
  try {
    const ppt = new Presentation('Shape Test Presentation');
    const slide = ppt.addSlide();

    slide.addText('Shape Test', {
      x: 0.5,
      y: 0.5,
      w: 8,
      h: 1,
      fontFace: 'Arial',
      fontSize: 28,
      bold: true,
      align: 'center'
    });

    addRectangle(slide, {
      x: 0.5,
      y: 1.5,
      w: 3,
      h: 1.5,
      fillColor: '92D050',
      lineColor: '000000',
      lineWidth: 2,
      transparency: 20
    });

    addCircle(slide, {
      x: 4.5,
      y: 1.5,
      w: 1.5,
      h: 1.5,
      fillColor: 'FF9900',
      lineColor: '000000',
      lineWidth: 2,
      transparency: 10
    });

    addLine(slide, {
      x1: 0.5,
      y1: 3.4,
      x2: 4.5,
      y2: 3.4,
      lineColor: '0000FF',
      lineWidth: 3
    });

    addArrow(slide, {
      x1: 5,
      y1: 3.4,
      x2: 8.5,
      y2: 3.4,
      lineColor: 'FF0000',
      lineWidth: 3
    });

    await ppt.save('test-shapes.pptx');
    console.log('Presentation saved successfully as test-shapes.pptx');
  } catch (error) {
    console.error('Error creating shape presentation:', error);
  }
}

// Run the test
testShapes();
