# PowerPoint Text Functionality Documentation

This document lists all the implemented text functionalities with their corresponding functions and proper implementation details.

## 1. Rich Text
### Function: `addText(slide, text, options)`
- **Purpose**: Adds basic rich text with formatting options
- **Parameters**:
  - `slide`: The slide object to add text to
  - `text`: The text content to display
  - `options`: Object containing formatting options:
    - `x`, `y`: Position coordinates
    - `w`, `h`: Width and height (optional)
    - `fontFace`: Font family (default: Arial)
    - `fontSize`: Font size (default: 12)
    - `color`: Text color (default: #000000)
    - `bold`: Bold formatting (default: false)
    - `italic`: Italic formatting (default: false)
    - `underline`: Underline formatting (default: false)
    - `align`: Text alignment (default: left)
    - `valign`: Vertical alignment (default: top)
    - `margin`: Margin around text (default: 0)

## 2. Strike
### Function: `addStrikethroughText(slide, text, options)`
- **Purpose**: Adds text with strikethrough formatting
- **Parameters**:
  - `slide`: The slide object to add text to
  - `text`: The text content to display
  - `options`: Object containing formatting options (same as addText plus):
    - `strike`: Boolean for strikethrough (default: true)

## 3. Hyperlinks
### Function: `addHyperlinkText(slide, text, url, options)`
- **Purpose**: Adds text with hyperlink functionality
- **Parameters**:
  - `slide`: The slide object to add text to
  - `text`: The text content to display
  - `url`: The URL to link to
  - `options`: Object containing formatting options (same as addText plus):
    - `hyperlink`: URL for hyperlink (automatically set from url parameter)

## 4. Text Rotation
### Function: `addRotatedText(slide, text, rotation, options)`
- **Purpose**: Adds text with specified rotation angle
- **Parameters**:
  - `slide`: The slide object to add text to
  - `text`: The text content to display
  - `rotation`: Rotation angle in degrees
  - `options`: Object containing formatting options (same as addText plus):
    - `rotation`: Rotation angle (set from rotation parameter)

## 5. Vertical Text
### Function: `addVerticalText(slide, text, options)`
- **Purpose**: Adds text in vertical orientation (rotated 90 degrees)
- **Parameters**:
  - `slide`: The slide object to add text to
  - `text`: The text content to display
  - `options`: Object containing formatting options (same as addText plus):
    - `rotation`: Automatically set to 90 degrees for vertical text

## 6. Text Autofit
### Function: `addAutofitText(slide, text, options)`
- **Purpose**: Adds text that automatically adjusts font size to fit container
- **Parameters**:
  - `slide`: The slide object to add text to
  - `text`: The text content to display
  - `options`: Object containing formatting options (same as addText plus):
    - `autofit`: Boolean for autofit (default: true)

## 7. Word Wrapping
### Function: `addWrappedText(slide, text, options)`
- **Purpose**: Adds text with automatic word wrapping
- **Parameters**:
  - `slide`: The slide object to add text to
  - `text`: The text content to display
  - `options`: Object containing formatting options (same as addText plus):
    - `wrap`: Boolean for word wrapping (default: true)

## 8. Paragraph Formatting
### Function: `addParagraphText(slide, text, options)`
- **Purpose**: Adds text with paragraph-specific formatting
- **Parameters**:
  - `slide`: The slide object to add text to
  - `text`: The text content to display
  - `options`: Object containing formatting options (same as addText plus):
    - `spacingBefore`: Space before paragraph (default: 0)
    - `spacingAfter`: Space after paragraph (default: 0)
    - `lineSpacing`: Line spacing multiplier (default: 1.0)

## 9. Bullets
### Function: `addBulletList(slide, items, options)`
- **Purpose**: Adds a bulleted list
- **Parameters**:
  - `slide`: The slide object to add text to
  - `items`: Array of strings for list items
  - `options`: Object containing formatting options (same as addText plus):
    - `bulletChar`: Character to use for bullets (default: •)
    - `bulletIndent`: Indentation for nested items (default: 0.25)

## 10. Numbered Lists
### Function: `addNumberedList(slide, items, options)`
- **Purpose**: Adds a numbered list
- **Parameters**:
  - `slide`: The slide object to add text to
  - `items`: Array of strings for list items
  - `options`: Object containing formatting options (same as addText plus):
    - `startNumber`: Starting number for list (default: 1)
    - `bulletIndent`: Indentation for nested items (default: 0.25)

## 11. Multi-level Bullets
### Function: `addMultiLevelBulletList(slide, items, options)`
- **Purpose**: Adds a multi-level bullet list with different bullet characters for each level
- **Parameters**:
  - `slide`: The slide object to add text to
  - `items`: Array of objects with `text` and `level` properties
  - `options`: Object containing formatting options (same as addText plus):
    - `bulletChars`: Array of bullet characters for different levels (default: ['•', '◦', '▪'])
    - `bulletIndent`: Indentation for each level (default: 0.25)

## 12. Search Text
### Function: `searchText(presentation, searchText)`
- **Purpose**: Searches for text across all slides in a presentation
- **Parameters**:
  - `presentation`: The presentation object to search in
  - `searchText`: String to search for
- **Returns**: Array of matching text objects with slide number and position information

## 13. Regex Replacement
### Function: `replaceTextWithRegex(presentation, searchPattern, replacement)`
- **Purpose**: Replaces text using regular expressions across all slides in a presentation
- **Parameters**:
  - `presentation`: The presentation object to modify
  - `searchPattern`: Regular expression pattern to search for
  - `replacement`: Replacement text string
- **Returns**: Number of replacements made

## Implementation Notes

All functions are designed to:
1. Accept slide objects and appropriate parameters
2. Return properly formatted text elements that integrate with the PowerPoint generation system
3. Handle edge cases gracefully without throwing exceptions
4. Follow consistent parameter structures and naming conventions
5. Be compatible with existing codebase and integration points
6. Accept data via terminal arguments (JSON strings for complex parameters)

The implementation leverages the underlying PowerPoint generation library (pptxgenjs) through the slide.addText() method, with additional properties passed through the options object to control formatting behaviors.

## Shape Functionality Integration

All shape manipulation functions follow the same data passing pattern as text functions:
- Parameters are passed via terminal arguments
- Complex parameters are encoded as JSON strings
- All functions integrate seamlessly with existing shape functionality
- Backward compatibility is maintained with existing shape creation methods
