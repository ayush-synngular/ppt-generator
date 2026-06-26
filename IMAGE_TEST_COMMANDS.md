# Image Manipulation Test Commands

This document lists the test commands that would be used to manually test the new image manipulation functionalities.

## 1. Crop Test Command
```bash
node src/index.js create image --path "assets/sample.png" --x 1 --y 1 --w 5 --h 4 --crop '{"x":0.2,"y":0.2,"w":0.6,"h":0.6}' --output cropped-image.pptx
```
This would be the command to test cropping functionality after implementing the `cropImage` function. Crop parameters should be passed as a JSON string via the `--crop` flag.

## 2. Resize Test Command
```bash
node src/index.js create image --path "assets/sample.png" --x 1 --y 1 --w 5 --h 4 --resize '{"width":3,"height":2}' --output resized-image.pptx
```
This would be the command to test resizing functionality after implementing the `resizeImage` function. Resize parameters should be passed as a JSON string via the `--resize` flag.

## 3. Rotate Test Command
```bash
node src/index.js create image --path "assets/sample.png" --x 1 --y 1 --w 5 --h 4 --rotation 45 --output rotated-image.pptx
```
This would be the command to test rotation functionality after implementing the `rotateImage` function. Rotation angle should be passed via the `--rotation` flag.

## 4. Transparency Test Command
```bash
node src/index.js create image --path "assets/sample.png" --x 1 --y 1 --w 5 --h 4 --transparency 30 --output transparent-image.pptx
```
This would be the command to test transparency functionality after implementing the `setImageTransparency` function. Transparency level should be passed via the `--transparency` flag.

## 5. Image Filters Test Command
```bash
node src/index.js create image --path "assets/sample.png" --x 1 --y 1 --w 5 --h 4 --filter '{"type":"grayscale","intensity":50}' --output filtered-image.pptx
```
This would be the command to test image filter functionality after implementing the `applyImageFilter` function. Filter parameters should be passed as a JSON string via the `--filter` flag.

## 6. Background Removal Test Command
```bash
node src/index.js create image --path "assets/sample.png" --x 1 --y 1 --w 5 --h 4 --backgroundRemoval '{"threshold":128,"feather":20}' --output bg-removed-image.pptx
```
This would be the command to test background removal functionality after implementing the `removeImageBackground` function. Background removal parameters should be passed as a JSON string via the `--backgroundRemoval` flag.

## 7. SVG Support Test Command
```bash
node src/index.js create image --svg "assets/graphic.svg" --x 1 --y 1 --w 5 --h 4 --scale 1.5 --output svg-image.pptx
```
This would be the command to test SVG support functionality after implementing the `addSvgImage` function. SVG parameters should be passed via appropriate flags.

## 8. Web Image URL Support Test Command
```bash
node src/index.js create image --url "https://example.com/image.jpg" --x 1 --y 1 --w 5 --h 4 --timeout 5000 --output web-image.pptx
```
This would be the command to test web image URL functionality after implementing the `addWebImage` function. Web image parameters should be passed via appropriate flags.

## General Testing Approach

To test these image manipulation functionalities, you would typically:

1. Create a test presentation with existing image examples
2. Add one of the new image manipulation functions to the presentation
3. Save the presentation and verify the results visually
4. Test edge cases like invalid parameters, unsupported formats, etc.
5. Test with various image formats and sizes

## Sample Test Structure

For testing purposes, you could create a new test file like `test-image-manipulation.js`:

```javascript
const { Presentation } = require('../src/index');
const { addImage } = require('../src/elements/image');

async function testImageManipulation() {
  try {
    const ppt = new Presentation('Image Manipulation Test');
    const slide = ppt.addSlide();
    
    // Create a basic image first (existing functionality)
    const image = addImage(slide, 'assets/sample.png', {
      x: 1,
      y: 1,
      w: 5,
      h: 4
    });
    
    // Test new image manipulation features (these would be implemented)
    // cropImage(slide, image, { x: 0.2, y: 0.2, w: 0.6, h: 0.6 });
    // resizeImage(slide, image, 3, 2);
    // rotateImage(slide, image, 45);
    // setImageTransparency(slide, image, 30);
    // applyImageFilter(slide, image, 'grayscale', { intensity: 50 });
    // removeImageBackground(slide, image, { threshold: 128, feather: 20 });
    // addSvgImage(slide, 'assets/graphic.svg', { x: 1, y: 6, w: 5, h: 2, scale: 1.5 });
    // addWebImage(slide, 'https://example.com/image.jpg', { x: 1, y: 9, w: 5, h: 2, timeout: 5000 });
    
    await ppt.save('test-image-manipulation.pptx');
    console.log('Image manipulation test saved successfully');
  } catch (error) {
    console.error('Error in image manipulation test:', error);
  }
}

testImageManipulation();
```

Each new function would be tested individually with appropriate parameters and validation, similar to how existing image functionality is tested.

## Implementation Notes

All commands should:
- Pass data via terminal arguments (JSON strings for complex parameters)
- Follow the existing CLI pattern established in the codebase
- Support the same error handling and validation as other commands
- Be backward compatible with existing image functionality