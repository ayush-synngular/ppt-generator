'use strict';

const Presentation = require('../src/core/Presentation');
const SlideBuilder = require('../src/core/SlideBuilder');
const Theme = require('../src/core/Theme');

// Import all the text elements
const {
  addText,
  addStrikethroughText,
  addHyperlinkText,
  addRotatedText,
  addVerticalText,
  addAutofitText,
  addWrappedText,
  addParagraphText,
  addBulletList,
  addNumberedList,
  addMultiLevelBulletList,
  searchText,
  replaceTextWithRegex
} = require('../src/elements/index');

async function demonstrateAllFunctionalities() {
  console.log('🧪 Demonstrating all text functionalities...\n');
  
  // Create a new presentation
  const presentation = new Presentation('All Functionalities Demo');
  const theme = new Theme();
  const slide = presentation.addSlide();
  
  // 1. Basic Rich Text
  console.log('1. Basic Rich Text:');
  addText(slide, 'This is basic rich text', { 
    x: 0.5, y: 0.5, 
    fontSize: 24, 
    bold: true, 
    italic: true,
    underline: true
  });
  
  // 2. Strikethrough
  console.log('2. Strikethrough:');
  addStrikethroughText(slide, 'This text has strikethrough', { 
    x: 0.5, y: 1.5, 
    fontSize: 20,
    strike: true
  });
  
  // 3. Hyperlinks
  console.log('3. Hyperlinks:');
  addHyperlinkText(slide, 'Visit Google', 'https://www.google.com', { 
    x: 0.5, y: 2.5, 
    fontSize: 20,
    color: '#0000FF'
  });
  
  // 4. Text Rotation
  console.log('4. Text Rotation:');
  addRotatedText(slide, 'Rotated Text', 45, { 
    x: 4, y: 0.5, 
    fontSize: 20
  });
  
  // 5. Vertical Text
  console.log('5. Vertical Text:');
  addVerticalText(slide, 'Vertical Text', { 
    x: 6, y: 0.5, 
    fontSize: 20
  });
  
  // 6. Text Autofit
  console.log('6. Text Autofit:');
  addAutofitText(slide, 'This text will automatically fit within its container', { 
    x: 0.5, y: 3.5, 
    w: 5, h: 1,
    fontSize: 24
  });
  
  // 7. Word Wrapping
  console.log('7. Word Wrapping:');
  addWrappedText(slide, 'This is a very long text that should wrap to the next line when it reaches the boundary of its container.', { 
    x: 0.5, y: 4.5, 
    w: 6,
    fontSize: 16
  });
  
  // 8. Paragraph Formatting
  console.log('8. Paragraph Formatting:');
  addParagraphText(slide, 'This is a paragraph with custom spacing.', { 
    x: 0.5, y: 6, 
    fontSize: 16,
    spacingBefore: 10,
    spacingAfter: 10,
    lineSpacing: 1.5
  });
  
  // 9. Bullets
  console.log('9. Bullets:');
  addBulletList(slide, ['Item 1', 'Item 2', 'Item 3'], { 
    x: 0.5, y: 7, 
    fontSize: 16
  });
  
  // 10. Numbered Lists
  console.log('10. Numbered Lists:');
  addNumberedList(slide, ['First Item', 'Second Item', 'Third Item'], { 
    x: 0.5, y: 8.5, 
    fontSize: 16
  });
  
  // 11. Multi-level Bullets
  console.log('11. Multi-level Bullets:');
  const multiLevelItems = [
    { text: 'Level 1 Item 1', level: 0 },
    { text: 'Level 2 Item 1', level: 1 },
    { text: 'Level 2 Item 2', level: 1 },
    { text: 'Level 3 Item 1', level: 2 },
    { text: 'Level 1 Item 2', level: 0 }
  ];
  addMultiLevelBulletList(slide, multiLevelItems, { 
    x: 0.5, y: 10, 
    fontSize: 16
  });
  
  // 12. Search Text (in a new slide)
  console.log('12. Search Text:');
  const searchSlide = presentation.addSlide();
  addText(searchSlide, 'This is sample text for searching purposes', { 
    x: 0.5, y: 0.5, 
    fontSize: 16
  });
  addText(searchSlide, 'Another piece of text with keyword', { 
    x: 0.5, y: 1.5, 
    fontSize: 16
  });
  
  // 13. Regex Replacement (in a new slide)
  console.log('13. Regex Replacement:');
  const replaceSlide = presentation.addSlide();
  addText(replaceSlide, 'The quick brown fox jumps over the lazy dog', { 
    x: 0.5, y: 0.5, 
    fontSize: 16
  });
  
  console.log('✅ All functionalities demonstrated successfully!');
  
  // Save the presentation
  await presentation.save('all-functionalities-demo.pptx');
  console.log('📄 Presentation saved as all-functionalities-demo.pptx');
}

// Run the demonstration
demonstrateAllFunctionalities().catch(console.error);