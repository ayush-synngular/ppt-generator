# PPT Generator

A Node.js utility for generating PowerPoint presentations programmatically.

## Features

- Create presentations with multiple slides
- Support for various element types (text, images, tables, charts, shapes, icons)
- Multiple layout options
- Theme management
- Export functionality
- Command Line Interface (CLI)

## Installation

```bash
npm install
```

## Usage

### Programmatic Usage

```javascript
const { Presentation } = require('./src');

const presentation = new Presentation();
// Add slides and elements to create your presentation
```

### CLI Usage

```bash
# Show help
node src/index.js help

# Show version
node src/index.js version

# Generate presentation from a script file
node src/index.js generate ./examples/test-presentation.js

# Create a new presentation with a template
node src/index.js create title-slide
```

## Project Structure

```
ppt-generator/
├── src/
│   ├── index.js
│   │
│   ├── core/
│   │   ├── Presentation.js
│   │   ├── Slide.js
│   │   └── Theme.js
│   │
│   ├── elements/
│   │   ├── text.js
│   │   ├── image.js
│   │   ├── table.js
│   │   ├── chart.js
│   │   ├── shape.js
│   │   └── icon.js
│   │
│   ├── layouts/
│   │   ├── title-slide.js
│   │   ├── two-column.js
│   │   └── comparison.js
│   │
│   ├── utils/
│   │   ├── color.js
│   │   ├── sizing.js
│   │   ├── positioning.js
│   │   └── export.js
│   │
│   └── constants/
│       ├── fonts.js
│       └── themes.js
│
├── package.json
└── README.md
```

## License

MIT
