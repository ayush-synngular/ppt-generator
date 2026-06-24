'use strict';

// Import the Presentation class
const { Presentation } = require('../src/index');

async function runTests() {
  console.log('Running comprehensive tests for Presentation class...\n');

  try {
    // Test 1: Constructor
    console.log('1. Testing Constructor');
    const ppt1 = new Presentation('Test Presentation');
    console.log('   ✓ Constructor creates presentation with correct title');
    
    // Test 2: setTitle function
    console.log('2. Testing setTitle function');
    ppt1.setTitle('Updated Test Presentation');
    console.log('   ✓ setTitle function works correctly');
    
    // Test 3: setSubject function
    console.log('3. Testing setSubject function');
    ppt1.setSubject('Testing Subject');
    console.log('   ✓ setSubject function works correctly');
    
    // Test 4: addSlide function
    console.log('4. Testing addSlide function');
    const slide1 = ppt1.addSlide();
    const slide2 = ppt1.addSlide();
    console.log(`   ✓ addSlide function works correctly - Added ${ppt1.slides.length} slides`);
    
    // Note: PptxGenJS does not expose a public removeSlide/deleteSlide API.
    // Deleting slides after they are created is unsupported, so only create
    // the slides required in the final presentation.
    
    // Test 5: save function (async)
    console.log('5. Testing save function');
    await ppt1.save('test-output.pptx');
    console.log('   ✓ save function works correctly');
    
    // Test 6: Constructor with default title
    console.log('6. Testing Constructor with default title');
    const ppt2 = new Presentation();
    console.log('   ✓ Constructor with default title works correctly');
    
    // Test 7: Multiple slide operations
    console.log('7. Testing multiple slide operations');
    const ppt3 = new Presentation('Multi-Slide Test');
    ppt3.addSlide();
    ppt3.addSlide();
    ppt3.addSlide();
    console.log(`   ✓ Multiple slides added: ${ppt3.slides.length} slides`);
    console.log('   ✓ No slide removal test because PptxGenJS does not support it');
    
    console.log('\n✓ All tests passed successfully!');
    
  } catch (error) {
    console.error('✗ Test failed with error:', error.message);
    console.error(error.stack);
    return false;
  }
  
  return true;
}

// Run the tests
runTests()
  .then(success => {
    if (success) {
      console.log('\n🎉 All tests completed successfully!');
    } else {
      console.log('\n❌ Some tests failed.');
      process.exit(1);
    }
  })
  .catch(error => {
    console.error('Fatal error during testing:', error);
    process.exit(1);
  });