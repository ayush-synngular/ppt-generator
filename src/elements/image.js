'use strict';

/**
 * Add an image to a slide using PptxGenJS.
 * @param {Object} slide - The slide object returned by pptxgenjs.
 * @param {string} imagePath - The image file path.
 * @param {Object} options - Image formatting options.
 * @returns {Object} The added image object.
 */
function addImage(slide, imagePath, options = {}) {
  const {
    x = 1,
    y = 1,
    w = 4,
    h = 3,
    transparency,
    rotation,
    hyperlink
  } = options;

  const imageOptions = {
    path: imagePath,
    x,
    y,
    w,
    h
  };

  if (transparency !== undefined) {
    imageOptions.transparency = transparency;
  }

  if (rotation !== undefined) {
    imageOptions.rotation = rotation;
  }

  if (hyperlink !== undefined) {
    imageOptions.hyperlink = hyperlink;
  }

  return slide.addImage(imageOptions);
}

module.exports = {
  addImage
};