'use strict';

/**
 * Example CLI script to demonstrate CLI functionality
 * This can be used with: node src/index.js generate ./examples/cli-example.js
 */

async function cliExample() {
  try {
    console.log('Running CLI example...');
    
    // Import the required modules
    const { Presentation, addText, renderTitle } = require('../src/index');
    
    // Create a new presentation
    const presentation = new Presentation('CLI Example Presentation');
    
    // Add a title slide
    const slide1 = presentation.addSlide();
    slide1.addTitle('Welcome to CLI Example');
    slide1.addSubtitle('Demonstrating CLI functionality');
    
    // Add a content slide
    const slide2 = presentation.addSlide();
    slide2.addTitle('Features Demonstrated');
    
    // Add some text content
    const text1 = slide2.addText('• CLI Command Support', { x: 1, y: 1.5, w: 8, h: 0.5 });
    const text2 = slide2.addText('• Programmatic Usage Still Works', { x: 1, y: 2.5, w: 8, h: 0.5 });
    const text3 = slide2.addText('• Easy Generation from Scripts', { x: 1, y: 3.5, w: 8, h: 0.5 });
    
    // Save the presentation
    await presentation.save('cli-example-output.pptx');
    console.log('SUCCESS: CLI example presentation saved as cli-example-output.pptx');
    
  } catch (error) {
    console.error('Error in CLI example:', error);
  }
}

// Export the function so it can be called by the CLI
module.exports = cliExample;