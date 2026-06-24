'use strict';

/**
 * Add a table to a slide using PptxGenJS.
 * @param {Object} slide - The slide object returned by pptxgenjs.
 * @param {Array<Array>} rows - Table rows and cells.
 * @param {Object} options - Table layout and styling options.
 * @returns {Object} The slide object.
 */
function addTable(slide, rows, options = {}) {
  const {
    x = 0.5,
    y = 1.25,
    w = 9,
    h = 3,
    borderColor = '000000',
    borderWidth = 1,
    fontSize = 12,
    color = '000000',
    fillColor,
    boldHeader = false,
    autoFit
  } = options;

  const tableOptions = {
    x,
    y,
    w,
    h,
    fontSize,
    color,
    border: {
      type: 'solid',
      color: borderColor,
      pt: borderWidth
    }
  };

  if (fillColor !== undefined) {
    tableOptions.fill = { color: fillColor };
  }

  if (autoFit !== undefined) {
    tableOptions.autoFit = autoFit;
  }

  const tableRows = rows.map((row, rowIndex) => {
    if (rowIndex !== 0 || !boldHeader) {
      return row;
    }

    return row.map(cell => {
      if (typeof cell === 'object' && cell !== null) {
        return {
          ...cell,
          options: {
            ...(cell.options || {}),
            bold: true
          }
        };
      }

      return {
        text: cell,
        options: {
          bold: true
        }
      };
    });
  });

  return slide.addTable(tableRows, tableOptions);
}

module.exports = {
  addTable
};