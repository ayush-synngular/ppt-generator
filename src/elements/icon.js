'use strict';

/**
 * Icon element class
 */
class IconElement {
  constructor(name = 'default-icon', size = 24) {
    this.type = 'icon';
    this.iconName = name;
    this.size = size;
    this.position = { x: 0, y: 0 };
  }

  /**
   * Set the position of the icon element
   * @param {number} x - X coordinate
   * @param {number} y - Y coordinate
   */
  setPosition(x, y) {
    this.position.x = x;
    this.position.y = y;
  }

  /**
   * Set the size of the icon element
   * @param {number} size - Icon size
   */
  setSize(size) {
    this.size = size;
  }
}

function addIcon(slide, iconName = 'star', options = {}) {
  const iconMap = {
    star: '★',
    check: '✔',
    cross: '✖',
    heart: '♥',
    info: 'ℹ',
    warning: '⚠',
    diamond: '◆',
    circle: '●'
  };

  const iconChar = options.iconChar || iconMap[iconName] || iconMap[iconName?.toLowerCase()] || iconName || '★';
  const {
    x = 4.25,
    y = 2,
    w = 1.5,
    h = 1.5,
    fontFace = 'Arial Black',
    fontSize = 120,
    color = '#0078D4',
    bold = true,
    align = 'center',
    valign = 'middle'
  } = options;

  return slide.addText(iconChar, {
    x,
    y,
    w,
    h,
    fontFace,
    fontSize,
    color,
    bold,
    align,
    valign
  });
}

module.exports = {
  IconElement,
  addIcon
};