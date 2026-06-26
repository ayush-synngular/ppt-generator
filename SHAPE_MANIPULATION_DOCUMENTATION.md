# Shape Manipulation Functionality Documentation

This document lists all the new shape manipulation functionalities that need to be implemented for the PowerPoint generator.

## 1. Triangle
### Function: `addTriangle(slide, options)`
- **Purpose**: Adds a triangle shape to a slide
- **Parameters**:
  - `slide`: The slide object to add the shape to
  - `options`: Object containing shape options:
    - `x`, `y`: Position coordinates
    - `w`, `h`: Width and height
    - `fillColor`: Fill color of the triangle
    - `lineColor`: Border color of the triangle
    - `lineWidth`: Border width
    - `transparency`: Transparency level (0-100)
    - `rotation`: Rotation angle in degrees
- **Returns**: Triangle shape object or boolean indicating success
- **Implementation Notes**: Should create a triangular shape with three points. Data should be passed via terminal arguments.

## 2. Diamond
### Function: `addDiamond(slide, options)`
- **Purpose**: Adds a diamond (rhombus) shape to a slide
- **Parameters**:
  - `slide`: The slide object to add the shape to
  - `options`: Object containing shape options:
    - `x`, `y`: Position coordinates
    - `w`, `h`: Width and height
    - `fillColor`: Fill color of the diamond
    - `lineColor`: Border color of the diamond
    - `lineWidth`: Border width
    - `transparency`: Transparency level (0-100)
    - `rotation`: Rotation angle in degrees
- **Returns**: Diamond shape object or boolean indicating success
- **Implementation Notes**: Should create a diamond shape with four equal sides. Data should be passed via terminal arguments.

## 3. Pentagon
### Function: `addPentagon(slide, options)`
- **Purpose**: Adds a pentagon shape to a slide
- **Parameters**:
  - `slide`: The slide object to add the shape to
  - `options`: Object containing shape options:
    - `x`, `y`: Position coordinates
    - `w`, `h`: Width and height
    - `fillColor`: Fill color of the pentagon
    - `lineColor`: Border color of the pentagon
    - `lineWidth`: Border width
    - `transparency`: Transparency level (0-100)
    - `rotation`: Rotation angle in degrees
- **Returns**: Pentagon shape object or boolean indicating success
- **Implementation Notes**: Should create a five-sided polygon. Data should be passed via terminal arguments.

## 4. Hexagon
### Function: `addHexagon(slide, options)`
- **Purpose**: Adds a hexagon shape to a slide
- **Parameters**:
  - `slide`: The slide object to add the shape to
  - `options`: Object containing shape options:
    - `x`, `y`: Position coordinates
    - `w`, `h`: Width and height
    - `fillColor`: Fill color of the hexagon
    - `lineColor`: Border color of the hexagon
    - `lineWidth`: Border width
    - `transparency`: Transparency level (0-100)
    - `rotation`: Rotation angle in degrees
- **Returns**: Hexagon shape object or boolean indicating success
- **Implementation Notes**: Should create a six-sided polygon. Data should be passed via terminal arguments.

## 5. Callout
### Function: `addCallout(slide, text, options)`
- **Purpose**: Adds a callout shape with text to a slide
- **Parameters**:
  - `slide`: The slide object to add the shape to
  - `text`: Text content for the callout
  - `options`: Object containing shape options:
    - `x`, `y`: Position coordinates
    - `w`, `h`: Width and height
    - `fillColor`: Fill color of the callout
    - `lineColor`: Border color of the callout
    - `lineWidth`: Border width
    - `transparency`: Transparency level (0-100)
    - `rotation`: Rotation angle in degrees
    - `textOptions`: Additional text formatting options
- **Returns**: Callout shape object or boolean indicating success
- **Implementation Notes**: Should create a speech bubble-like shape with attached text. Data should be passed via terminal arguments.

## 6. Chevron
### Function: `addChevron(slide, options)`
- **Purpose**: Adds a chevron (V-shaped) pattern to a slide
- **Parameters**:
  - `slide`: The slide object to add the shape to
  - `options`: Object containing shape options:
    - `x`, `y`: Position coordinates
    - `w`, `h`: Width and height
    - `fillColor`: Fill color of the chevron
    - `lineColor`: Border color of the chevron
    - `lineWidth`: Border width
    - `transparency`: Transparency level (0-100)
    - `rotation`: Rotation angle in degrees
- **Returns**: Chevron shape object or boolean indicating success
- **Implementation Notes**: Should create a V-shaped or chevron pattern. Data should be passed via terminal arguments.

## 7. Flowchart Shapes
### Function: `addFlowchartShape(slide, shapeType, options)`
- **Purpose**: Adds various flowchart shapes to a slide
- **Parameters**:
  - `slide`: The slide object to add the shape to
  - `shapeType`: Type of flowchart shape (e.g., 'process', 'decision', 'document', 'data', etc.)
  - `options`: Object containing shape options:
    - `x`, `y`: Position coordinates
    - `w`, `h`: Width and height
    - `fillColor`: Fill color of the shape
    - `lineColor`: Border color of the shape
    - `lineWidth`: Border width
    - `transparency`: Transparency level (0-100)
    - `rotation`: Rotation angle in degrees
- **Returns**: Flowchart shape object or boolean indicating success
- **Implementation Notes**: Should support common flowchart shapes. Data should be passed via terminal arguments.

## 8. Freeform Paths
### Function: `addFreeformPath(slide, points, options)`
- **Purpose**: Adds a freeform path shape to a slide
- **Parameters**:
  - `slide`: The slide object to add the shape to
  - `points`: Array of point coordinates [[x1,y1], [x2,y2], ...] defining the path
  - `options`: Object containing shape options:
    - `fillColor`: Fill color of the shape
    - `lineColor`: Border color of the shape
    - `lineWidth`: Border width
    - `transparency`: Transparency level (0-100)
    - `rotation`: Rotation angle in degrees
- **Returns**: Freeform path shape object or boolean indicating success
- **Implementation Notes**: Should allow custom polygon shapes by defining points. Data should be passed via terminal arguments.

## 9. Group/Ungroup
### Function: `groupShapes(slide, shapeIds)`
- **Purpose**: Groups multiple shapes together
- **Parameters**:
  - `slide`: The slide object containing the shapes
  - `shapeIds`: Array of shape IDs to group together
- **Returns**: Grouped shape object or boolean indicating success
- **Implementation Notes**: Should combine multiple shapes into a single group. Data should be passed via terminal arguments.

### Function: `ungroupShapes(slide, groupId)`
- **Purpose**: Ungroups a previously grouped shape
- **Parameters**:
  - `slide`: The slide object containing the group
  - `groupId`: ID of the group to ungroup
- **Returns**: Boolean indicating success
- **Implementation Notes**: Should separate a grouped shape back into individual shapes. Data should be passed via terminal arguments.

## Implementation Notes

All new shape manipulation functions should:
1. Be designed to work with existing shape objects created by other shape functions
2. Not modify or interfere with existing shape functionality
3. Handle edge cases gracefully without throwing exceptions
4. Follow consistent parameter structures and naming conventions
5. Integrate well with the PowerPoint generation system
6. Be backward compatible with existing code
7. Support data passing via terminal arguments (JSON strings for complex parameters)
8. Support common shape manipulation workflows

Note: These functions are designed to extend the existing shape functionality and should be added to the shape.js file or a new dedicated shape utilities file, not to modify the existing shape functions themselves.

## Terminal Usage Examples

All shape manipulation functions should be usable through the CLI with terminal arguments:

```bash
# Basic rectangle creation
node src/index.js create shape --type rect --x 1 --y 1 --w 3 --h 2 --output presentation.pptx

# Triangle with custom styling
node src/index.js create shape --type triangle --x 2 --y 2 --w 2 --h 2 --fillColor FF0000 --lineColor 000000 --output presentation.pptx

# Grouping shapes
node src/index.js create shape --type rect --x 1 --y 1 --w 2 --h 2 --output presentation.pptx
node src/index.js create shape --type circle --x 4 --y 1 --w 2 --h 2 --output presentation.pptx
# Grouping would be handled separately in CLI commands

# Flowchart shape
node src/index.js create shape --type decision --x 1 --y 1 --w 3 --h 2 --output presentation.pptx
```

The implementation should follow the established pattern in the codebase where data is passed through JSON strings via terminal arguments, ensuring that all functionality works with terminal input as required.