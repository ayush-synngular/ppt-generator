'use strict';

// Import the Presentation class and image helper
const { Presentation } = require('../src/index');
const { addImage } = require('../src/elements/image');

async function testImage() {
  try {
    const ppt = new Presentation('Image Test Presentation');
    const slide = ppt.addSlide();

    // Add a title
    slide.addText('Sample Image Slide', {
      x: 0.5,
      y: 0.5,
      w: 8,
      h: 1,
      fontFace: 'Arial',
      fontSize: 28,
      bold: true,
      align: 'center'
    });

    // Add an image from the assets folder
    addImage(slide, 'assets/sample.png', {
      x: 1,
      y: 1.75,
      w: 6,
      h: 4,
      transparency: 0,
      rotation: 0,
      hyperlink: undefined
    });

    await ppt.save('test-image.pptx');
    console.log('Presentation saved successfully as test-image.pptx');
  } catch (error) {
    console.error('Error creating image presentation:', error);
  }
}

// Run the test
testImage();