'use strict';

// Import the Presentation class
const { Presentation } = require('../src/index');

async function testPresentation() {
  try {
    console.log('Testing Presentation class functions...');
    
    // Test 1: Create a new presentation
    console.log('1. Creating presentation with title "Test Presentation"');
    const ppt = new Presentation('Test Presentation');
    
    // Test 2: Set presentation properties
    console.log('2. Setting presentation subject');
    ppt.setSubject('Testing Subject');
    
    // Test 3: Add slides
    console.log('3. Adding a single slide');
    const slide1 = ppt.addSlide();
    
    // Test 4: Check slides count
    console.log(`4. Total slides: ${ppt.slides.length}`);
    
    // Slide deletion is unsupported by the public PptxGenJS API, so we only
    // create the slides that should exist in the final presentation.
    console.log('5. Skipping slide removal because PptxGenJS does not support it');
    console.log(`6. Total slides after skip: ${ppt.slides.length}`);
    
    // Test 6: Update presentation title
    console.log('7. Updating presentation title');
    ppt.setTitle('Updated Test Presentation');
    
    // Test 7: Save presentation (this will create a file)
    console.log('8. Saving presentation as test-presentation.pptx');
    await ppt.save('test-presentation.pptx');
    console.log('SUCCESS: Presentation saved successfully as test-presentation.pptx');
    
  } catch (error) {
    console.error('Error testing presentation:', error);
  }
}

// Run the test
testPresentation();