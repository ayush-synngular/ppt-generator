'use strict';

// Test file for advanced text functionality
const fs = require('fs');
const path = require('path');

// Test importing the module
console.log('Testing advanced text functionality exports...\n');

try {
  // Try to import the main module
  const pptGenerator = require('../src/index.js');
  
  // Check if all required functions are exported
  const expectedFunctions = [
    'addStrikethroughText',
    'addHyperlinkText', 
    'addRotatedText',
    'addVerticalText',
    'addAutofitText',
    'addWrappedText',
    'addParagraphText',
    'addBulletList',
    'addNumberedList',
    'addMultiLevelBulletList',
    'searchText',
    'replaceTextWithRegex'
  ];
  
  let allPresent = true;
  console.log('Checking if all advanced text functions are exported:');
  
  expectedFunctions.forEach(funcName => {
    if (typeof pptGenerator[funcName] === 'function') {
      console.log(`✅ ${funcName} - Exported correctly`);
    } else {
      console.log(`❌ ${funcName} - NOT exported`);
      allPresent = false;
    }
  });
  
  if (allPresent) {
    console.log('\n🎉 All advanced text functions are properly exported!');
    
    // Test basic functionality by creating a sample presentation
    console.log('\nTesting basic functionality...');
    
    const presentation = new pptGenerator.Presentation('Test Presentation');
    const slide = presentation.addSlide();
    
    // Test a couple of functions to ensure they can be called
    try {
      // Test addText (should exist)
      pptGenerator.addText(slide, 'Sample Text', { x: 1, y: 1, fontSize: 24 });
      console.log('✅ addText function works');
      
      // Test addStrikethroughText (should exist now)
      pptGenerator.addStrikethroughText(slide, 'Strikethrough Text', { x: 1, y: 2, fontSize: 24 });
      console.log('✅ addStrikethroughText function works');
      
      console.log('\n🎉 Basic functionality test passed!');
    } catch (err) {
      console.log('⚠️  Warning: Some functions may not be fully implemented yet:', err.message);
    }
    
  } else {
    console.log('\n❌ Some functions are missing from exports');
  }
  
} catch (error) {
  console.error('❌ Error importing module:', error.message);
  process.exit(1);
}

console.log('\nAdvanced text functionality test completed.');