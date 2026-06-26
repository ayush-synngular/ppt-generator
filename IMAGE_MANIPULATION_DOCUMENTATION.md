# Image Manipulation Functionality Documentation

This document lists all the new image manipulation functionalities that need to be implemented for the PowerPoint generator.

## 1. Crop
### Function: `cropImage(slide, image, cropOptions)`
- **Purpose**: Crops an image to a specified region
- **Parameters**:
  - `slide`: The slide object containing the image
  - `image`: The image object to crop
  - `cropOptions`: Object containing cropping parameters:
    - `x`: X coordinate of crop area (0-1, relative to image)
    - `y`: Y coordinate of crop area (0-1, relative to image)
    - `w`: Width of crop area (0-1, relative to image)
    - `h`: Height of crop area (0-1, relative to image)
- **Returns**: Modified image object or boolean indicating success
- **Implementation Notes**: Should work with existing image objects and support relative positioning. Data should be passed via terminal arguments.

## 2. Resize
### Function: `resizeImage(slide, image, width, height)`
- **Purpose**: Resizes an image to specified dimensions
- **Parameters**:
  - `slide`: The slide object containing the image
  - `image`: The image object to resize
  - `width`: New width (in inches or relative units)
  - `height`: New height (in inches or relative units)
- **Returns**: Modified image object or boolean indicating success
- **Implementation Notes**: Should maintain aspect ratio when only one dimension is specified. Data should be passed via terminal arguments.

## 3. Rotate
### Function: `rotateImage(slide, image, rotationAngle)`
- **Purpose**: Rotates an image by a specified angle
- **Parameters**:
  - `slide`: The slide object containing the image
  - `image`: The image object to rotate
  - `rotationAngle`: Rotation angle in degrees (positive = clockwise)
- **Returns**: Modified image object or boolean indicating success
- **Implementation Notes**: Should support arbitrary rotation angles. Data should be passed via terminal arguments.

## 4. Transparency
### Function: `setImageTransparency(slide, image, transparencyLevel)`
- **Purpose**: Sets the transparency level of an image
- **Parameters**:
  - `slide`: The slide object containing the image
  - `image`: The image object to modify
  - `transparencyLevel`: Transparency level (0-100, where 0 is opaque, 100 is fully transparent)
- **Returns**: Modified image object or boolean indicating success
- **Implementation Notes**: Should support percentage-based transparency. Data should be passed via terminal arguments.

## 5. Image Filters
### Function: `applyImageFilter(slide, image, filterType, filterOptions)`
- **Purpose**: Applies visual filters to images
- **Parameters**:
  - `slide`: The slide object containing the image
  - `image`: The image object to filter
  - `filterType`: Type of filter to apply (e.g., 'grayscale', 'sepia', 'blur', 'brightness')
  - `filterOptions`: Additional options for the filter (specific to each filter type)
- **Returns**: Modified image object or boolean indicating success
- **Implementation Notes**: Should support common image filters. Data should be passed via terminal arguments.

## 6. Background Removal
### Function: `removeImageBackground(slide, image, removalOptions)`
- **Purpose**: Removes background from images using AI or masking techniques
- **Parameters**:
  - `slide`: The slide object containing the image
  - `image`: The image object to modify
  - `removalOptions`: Options for background removal:
    - `threshold`: Threshold for background detection (0-255)
    - `feather`: Softness of the mask edges (0-100)
    - `color`: Color to replace background with (if applicable)
- **Returns**: Modified image object or boolean indicating success
- **Implementation Notes**: Should support automatic background detection. Data should be passed via terminal arguments.

## 7. SVG Support
### Function: `addSvgImage(slide, svgPath, options)`
- **Purpose**: Adds SVG images to slides with proper scaling and rendering
- **Parameters**:
  - `slide`: The slide object to add the image to
  - `svgPath`: Path to the SVG file
  - `options`: SVG-specific options:
    - `x`, `y`: Position coordinates
    - `w`, `h`: Width and height
    - `scale`: Scale factor for SVG rendering
    - `backgroundColor`: Background color for SVG (if needed)
- **Returns**: SVG image object or boolean indicating success
- **Implementation Notes**: Should properly render SVG vector graphics. Data should be passed via terminal arguments.

## 8. Web Image URL Support
### Function: `addWebImage(slide, imageUrl, options)`
- **Purpose**: Downloads and adds images from web URLs to slides
- **Parameters**:
  - `slide`: The slide object to add the image to
  - `imageUrl`: URL of the image to download
  - `options`: Image options:
    - `x`, `y`: Position coordinates
    - `w`, `h`: Width and height
    - `timeout`: Download timeout in milliseconds
    - `cache`: Whether to cache downloaded images
- **Returns**: Web image object or boolean indicating success
- **Implementation Notes**: Should handle HTTP/HTTPS downloads with error handling. Data should be passed via terminal arguments.

## Implementation Notes

All new image manipulation functions should:
1. Be designed to work with existing image objects created by `addImage()`
2. Not modify or interfere with existing image functionality
3. Handle edge cases gracefully without throwing exceptions
4. Follow consistent parameter structures and naming conventions
5. Integrate well with the PowerPoint generation system
6. Be backward compatible with existing code
7. Support data passing via terminal arguments (JSON strings for complex parameters)
8. Support common image manipulation workflows

Note: These functions are designed to extend the existing image functionality and should be added to the image.js file or a new dedicated image utilities file, not to modify the existing `addImage` function itself.

## Terminal Usage Examples

All image manipulation functions should be usable through the CLI with terminal arguments:

```bash
# Basic image creation with terminal arguments
node src/index.js create image --path "assets/photo.jpg" --x 1 --y 1 --w 5 --h 4 --output presentation.pptx

# Image with rotation and transparency
node src/index.js create image --path "assets/photo.jpg" --rotation 45 --transparency 30 --output presentation.pptx

# Web image with download options
node src/index.js create image --url "https://example.com/image.jpg" --timeout 5000 --output presentation.pptx

# SVG image
node src/index.js create image --svg "assets/graphic.svg" --scale 1.5 --output presentation.pptx
```

The implementation should follow the established pattern in the codebase where data is passed through JSON strings via terminal arguments, ensuring that all functionality works with terminal input as required.