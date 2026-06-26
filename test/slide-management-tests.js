'use strict';

// Import the Presentation class
const { Presentation } = require('../src/index');

async function runSlideManagementTests() {
  console.log('Running slide management tests...\n');

  try {
    // Test 1: Create a new presentation
    console.log('1. Creating presentation with title "Slide Management Test"');
    const ppt = new Presentation('Slide Management Test');
    
    // Test 2: Add initial slides
    console.log('2. Adding initial slides');
    const slide1 = ppt.addSlide();
    const slide2 = ppt.addSlide();
    const slide3 = ppt.addSlide();
    console.log(`   ✓ Added ${ppt.slides.length} slides`);
    
    // Test 3: Insert slide
    console.log('3. Testing insertSlide function');
    const insertedSlide = ppt.insertSlide(1);
    console.log(`   ✓ Inserted slide at index 1. Total slides: ${ppt.slides.length}`);
    
    // Test 4: Duplicate slide
    console.log('4. Testing duplicateSlide function');
    const duplicatedSlide = ppt.duplicateSlide(0);
    console.log(`   ✓ Duplicated slide at index 0. Total slides: ${ppt.slides.length}`);
    
    // Test 5: Move slide
    console.log('5. Testing moveSlide function');
    const moved = ppt.moveSlide(2, 0);
    console.log(`   ✓ Moved slide from index 2 to 0: ${moved}`);
    
    // Test 6: Hide slide
    console.log('6. Testing hideSlide function');
    const hidden = ppt.hideSlide(1);
    console.log(`   ✓ Hidden slide at index 1: ${hidden}`);
    
    // Test 7: Delete slide
    console.log('7. Testing deleteSlide function');
    const deleted = ppt.deleteSlide(3);
    console.log(`   ✓ Deleted slide at index 3: ${deleted}`);
    
    // Test 8: Verify slide count
    console.log('8. Verifying final slide count');
    console.log(`   ✓ Final slide count: ${ppt.slides.length}`);
    
    // Test 9: Test error handling for invalid inputs
    console.log('9. Testing error handling');
    try {
      ppt.insertSlide(-1);
      console.log('   ✗ Should have thrown error for negative index');
    } catch (e) {
      console.log('   ✓ Correctly threw error for negative index');
    }
    
    try {
      ppt.duplicateSlide(100);
      console.log('   ✗ Should have thrown error for invalid index');
    } catch (e) {
      console.log('   ✓ Correctly threw error for invalid index');
    }
    
    try {
      ppt.moveSlide(100, 0);
      console.log('   ✗ Should have returned false for invalid indices');
    } catch (e) {
      console.log('   ✓ Correctly handled invalid indices in moveSlide');
    }
    
    try {
      ppt.hideSlide(-1);
      console.log('   ✗ Should have returned false for invalid index');
    } catch (e) {
      console.log('   ✓ Correctly handled invalid index in hideSlide');
    }
    
    // Test 10: Save presentation (this will create a file)
    console.log('10. Saving presentation as slide-management-test.pptx');
    await ppt.save('slide-management-test.pptx');
    console.log('   ✓ Presentation saved successfully as slide-management-test.pptx');
    
    console.log('\n✅ All slide management tests passed successfully!');
    
    return true;
    
  } catch (error) {
    console.error('❌ Test failed with error:', error.message);
    console.error(error.stack);
    return false;
  }
}

// Run the tests
runSlideManagementTests()
  .then(success => {
    if (success) {
      console.log('\n🎉 All slide management tests completed successfully!');
    } else {
      console.log('\n❌ Some tests failed.');
      process.exit(1);
    }
  })
  .catch(error => {
    console.error('Fatal error during testing:', error);
    process.exit(1);
  });