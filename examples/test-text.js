'use strict';

// Import the PPT generator components
const { Presentation, TextElement } = require('../src/index');

// Alternative approach using the direct function
const { addText } = require('../src/index');

async function testText() {
  try {
    // Create a new presentation
    const ppt = new Presentation('Test Presentation');
    
    // Add a slide
    const slide = ppt.addSlide();
    
    // Add a title using addText
    addText(slide, 'Welcome to Our Presentation', {
      x: 0.5,
      y: 0.5,
      w: 9,
      h: 1,
      fontFace: 'Calibri Light',
      fontSize: 32,
      color: '#2E74B5',
      bold: true,
      align: 'center'
    });
    
    // Add a paragraph using addText
    addText(slide, 'This is a sample paragraph demonstrating the text functionality.\n\n' +
                  'You can add multiple lines of text and format them as needed.', {
      x: 1,
      y: 2,
      w: 8,
      h: 3,
      fontFace: 'Calibri',
      fontSize: 18,
      color: '#333333',
      align: 'left'
    });
    
    // Save the presentation
    await ppt.save('test-text.pptx');
    console.log('Presentation saved successfully as test-text.pptx');
    
  } catch (error) {
    console.error('Error creating presentation:', error);
  }
}

// Run the test
testText();