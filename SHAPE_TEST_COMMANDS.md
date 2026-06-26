# Shape Manipulation Test Commands

This document lists the test commands that would be used to manually test the new shape manipulation functionalities.

## 1. Triangle Test Command
```bash
node src/index.js create shape --type triangle --x 1 --y 1 --w 2 --h 2 --fillColor FF0000 --lineColor 000000 --lineWidth 2 --output triangle-shape.pptx
```
This would be the command to test triangle functionality after implementing the `addTriangle` function. Shape parameters should be passed via terminal arguments.

## 2. Diamond Test Command
```bash
node src/index.js create shape --type diamond --x 4 --y 1 --w 2 --h 2 --fillColor 00FF00 --lineColor 000000 --transparency 30 --output diamond-shape.pptx
```
This would be the command to test diamond functionality after implementing the `addDiamond` function. Shape parameters should be passed via terminal arguments.

## 3. Pentagon Test Command
```bash
node src/index.js create shape --type pentagon --x 1 --y 4 --w 2 --h 2 --fillColor 0000FF --lineColor 000000 --rotation 45 --output pentagon-shape.pptx
```
This would be the command to test pentagon functionality after implementing the `addPentagon` function. Shape parameters should be passed via terminal arguments.

## 4. Hexagon Test Command
```bash
node src/index.js create shape --type hexagon --x 4 --y 4 --w 2 --h 2 --fillColor FFFF00 --lineColor 000000 --output hexagon-shape.pptx
```
This would be the command to test hexagon functionality after implementing the `addHexagon` function. Shape parameters should be passed via terminal arguments.

## 5. Callout Test Command
```bash
node src/index.js create shape --type callout --text "This is a callout" --x 1 --y 7 --w 3 --h 1.5 --fillColor FFFFFF --lineColor 000000 --output callout-shape.pptx
```
This would be the command to test callout functionality after implementing the `addCallout` function. Callout parameters should be passed via terminal arguments.

## 6. Chevron Test Command
```bash
node src/index.js create shape --type chevron --x 5 --y 7 --w 2 --h 2 --fillColor FF00FF --lineColor 000000 --rotation 90 --output chevron-shape.pptx
```
This would be the command to test chevron functionality after implementing the `addChevron` function. Shape parameters should be passed via terminal arguments.

## 7. Flowchart Shapes Test Command
```bash
node src/index.js create shape --type decision --x 1 --y 10 --w 2 --h 2 --fillColor E0E0E0 --lineColor 000000 --output flowchart-shape.pptx
```
This would be the command to test flowchart shapes functionality after implementing the `addFlowchartShape` function. Flowchart parameters should be passed via terminal arguments.

## 8. Freeform Paths Test Command
```bash
node src/index.js create shape --type freeform --points '[{"x":1,"y":12},{"x":3,"y":13},{"x":5,"y":12},{"x":3,"y":11}]' --fillColor 00FFFF --lineColor 000000 --output freeform-shape.pptx
```
This would be the command to test freeform paths functionality after implementing the `addFreeformPath` function. Points should be passed as a JSON array via the `--points` flag.

## 9. Group/Ungroup Test Command
```bash
# Create individual shapes first
node src/index.js create shape --type rect --x 1 --y 15 --w 1 --h 1 --output grouped-shapes.pptx
node src/index.js create shape --type circle --x 3 --y 15 --w 1 --h 1 --output grouped-shapes.pptx

# Grouping would be handled in a separate command or through scripting
# This demonstrates the concept of grouping shapes
```
This would be the command to test group/ungroup functionality after implementing the `groupShapes` and `ungroupShapes` functions. Grouping would typically be handled through a combination of commands or in a script.

## General Testing Approach

To test these shape manipulation functionalities, you would typically:

1. Create a test presentation with existing shape examples
2. Add one of the new shape manipulation functions to the presentation
3. Save the presentation and verify the results visually
4. Test edge cases like invalid parameters, unsupported shapes, etc.
5. Test with various styling options and transformations

## Sample Test Structure

For testing purposes, you could create a new test file like `test-shape-manipulation.js`:

```javascript
const { Presentation } = require('../src/index');
const { addRectangle, addCircle } = require('../src/elements/shape');

async function testShapeManipulation() {
  try {
    const ppt = new Presentation('Shape Manipulation Test');
    const slide = ppt.addSlide();
    
    // Create basic shapes first (existing functionality)
    const rect = addRectangle(slide, {
      x: 1,
      y: 1,
      w: 2,
      h: 2,
      fillColor: 'FF0000'
    });
    
    const circle = addCircle(slide, {
      x: 4,
      y: 1,
      w: 2,
      h: 2,
      fillColor: '00FF00'
    });
    
    // Test new shape manipulation features (these would be implemented)
    // addTriangle(slide, { x: 1, y: 4, w: 2, h: 2, fillColor: '0000FF' });
    // addDiamond(slide, { x: 4, y: 4, w: 2, h: 2, fillColor: 'FFFF00' });
    // addPentagon(slide, { x: 1, y: 7, w: 2, h: 2, fillColor: 'FF00FF' });
    // addHexagon(slide, { x: 4, y: 7, w: 2, h: 2, fillColor: '00FFFF' });
    // addCallout(slide, "Sample callout", { x: 1, y: 10, w: 3, h: 1.5, fillColor: 'FFFFFF' });
    // addChevron(slide, { x: 5, y: 10, w: 2, h: 2, fillColor: 'CCCCCC' });
    // addFlowchartShape(slide, 'process', { x: 1, y: 12, w: 2, h: 2, fillColor: 'E0E0E0' });
    // addFreeformPath(slide, [[1,14],[3,15],[5,14],[3,13]], { fillColor: 'A0A0A0' });
    // groupShapes(slide, [rect.id, circle.id]); // Assuming shapes have IDs
    
    await ppt.save('test-shape-manipulation.pptx');
    console.log('Shape manipulation test saved successfully');
  } catch (error) {
    console.error('Error in shape manipulation test:', error);
  }
}

testShapeManipulation();
```

Each new function would be tested individually with appropriate parameters and validation, similar to how existing shape functionality is tested.

## Implementation Notes

All commands should:
- Pass data via terminal arguments (JSON strings for complex parameters like points)
- Follow the existing CLI pattern established in the codebase
- Support the same error handling and validation as other commands
- Be backward compatible with existing shape functionality