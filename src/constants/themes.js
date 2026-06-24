'use strict';

/**
 * Theme constants
 */
const Themes = {
  DEFAULT: {
    name: 'Default',
    colors: {
      primary: '#007acc',
      secondary: '#6c757d',
      accent: '#17a2b8',
      background: '#ffffff',
      text: '#212529'
    },
    fonts: {
      title: 'Calibri Light',
      heading: 'Calibri',
      body: 'Tahoma'
    }
  },
  MODERN: {
    name: 'Modern',
    colors: {
      primary: '#4a90e2',
      secondary: '#50c878',
      accent: '#ff6b6b',
      background: '#f8f9fa',
      text: '#333333'
    },
    fonts: {
      title: 'Segoe UI',
      heading: 'Segoe UI',
      body: 'Segoe UI'
    }
  },
  CORPORATE: {
    name: 'Corporate',
    colors: {
      primary: '#003366',
      secondary: '#0066cc',
      accent: '#0099cc',
      background: '#ffffff',
      text: '#333333'
    },
    fonts: {
      title: 'Arial',
      heading: 'Arial',
      body: 'Arial'
    }
  }
};

module.exports = Themes;