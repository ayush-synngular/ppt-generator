'use strict';

const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');
const xml2js = require('xml2js');
const PresentationModel = require('./models/PresentationModel');
const SlideModel = require('./models/SlideModel');
const ShapeModel = require('./models/ShapeModel');

class PptReader {
  constructor() {
    this.parser = new xml2js.Parser({
      explicitArray: false,
      ignoreAttrs: true,
      trim: true,
      normalizeTags: false,
      normalize: true
    });
  }

  async read(filePath) {
    if (!filePath || typeof filePath !== 'string') {
      throw new Error('A valid file path is required for PptReader.read().');
    }

    if (!fs.existsSync(filePath)) {
      throw new Error(`PPTX file not found: ${filePath}`);
    }

    const zip = new AdmZip(filePath);
    const presentation = new PresentationModel({ filePath });

    presentation.metadata = await this.extractMetadata(zip);
    presentation.slides = await this.extractSlides(zip);

    return presentation;
  }

  async extractSlides(zip) {
    const slideEntries = zip
      .getEntries()
      .filter((entry) => /^ppt\/slides\/slide\d+\.xml$/.test(entry.entryName))
      .sort((a, b) => {
        const aIndex = Number(a.entryName.match(/slide(\d+)\.xml$/)[1]);
        const bIndex = Number(b.entryName.match(/slide(\d+)\.xml$/)[1]);
        return aIndex - bIndex;
      });

    const slideModels = [];
    for (let index = 0; index < slideEntries.length; index += 1) {
      const entry = slideEntries[index];
      const xml = entry.getData().toString('utf8');
      const parsed = await this.parser.parseStringPromise(xml);
      const text = this.extractText(parsed);
      const shapes = this.extractShapes(parsed);
      slideModels.push(new SlideModel({ slideNumber: index + 1, text, shapes }));
    }

    return slideModels;
  }

  extractText(parsedSlide) {
    const textSegments = [];

    const getParagraphText = (paragraph) => {
      const pieces = [];

      if (paragraph['a:r']) {
        const runs = Array.isArray(paragraph['a:r']) ? paragraph['a:r'] : [paragraph['a:r']];
        runs.forEach((run) => {
          if (run['a:t']) {
            pieces.push(run['a:t']);
          }
        });
      }

      if (paragraph['a:fld']) {
        const fields = Array.isArray(paragraph['a:fld']) ? paragraph['a:fld'] : [paragraph['a:fld']];
        fields.forEach((field) => {
          if (field['a:t']) {
            pieces.push(field['a:t']);
          }
        });
      }

      if (paragraph['a:t']) {
        pieces.push(paragraph['a:t']);
      }

      return pieces
        .filter((part) => typeof part === 'string' && part.trim() !== '')
        .map((part) => part.trim())
        .join(' ');
    };

    const scan = (node) => {
      if (!node || typeof node !== 'object') {
        return;
      }

      if (node['a:p']) {
        const paragraphs = Array.isArray(node['a:p']) ? node['a:p'] : [node['a:p']];
        paragraphs.forEach((paragraph) => {
          const text = getParagraphText(paragraph);
          if (text) {
            textSegments.push(text);
          }
        });
      }

      Object.values(node).forEach((child) => {
        if (child && typeof child === 'object') {
          scan(child);
        }
      });
    };

    scan(parsedSlide);

    return textSegments.filter((value, index, self) => typeof value === 'string' && value.trim() !== '' && self.indexOf(value) === index);
  }

  extractShapes(parsedSlide) {
    const shapes = [];

    const scan = (node) => {
      if (!node || typeof node !== 'object') {
        return;
      }

      if (node['p:sp'] && node['p:sp']['p:txBody']) {
        const text = this.extractText(node['p:sp']['p:txBody']).join(' ');
        shapes.push(new ShapeModel({ type: 'text', text, bounds: this.extractBounds(node['p:sp']) }));
      }

      if (node['p:pic']) {
        shapes.push(new ShapeModel({ type: 'image', text: '', bounds: this.extractBounds(node['p:pic']) }));
      }

      Object.values(node).forEach((child) => {
        if (typeof child === 'object') {
          scan(child);
        }
      });
    };

    scan(parsedSlide);

    return shapes;
  }

  extractBounds(shapeNode) {
    if (!shapeNode || typeof shapeNode !== 'object') {
      return {};
    }

    const xfrm = shapeNode['p:xfrm'] || shapeNode.xfrm || {};
    const off = xfrm['a:off'] || xfrm.off || {};
    const ext = xfrm['a:ext'] || xfrm.ext || {};
    const x = parseInt(off.x, 10) || 0;
    const y = parseInt(off.y, 10) || 0;
    const cx = parseInt(ext.cx, 10) || 0;
    const cy = parseInt(ext.cy, 10) || 0;

    return { x, y, w: cx, h: cy };
  }

  async extractMetadata(zip) {
    const metadata = {};
    const coreEntry = zip.getEntry('docProps/core.xml');
    const appEntry = zip.getEntry('docProps/app.xml');

    if (coreEntry) {
      try {
        const coreXml = coreEntry.getData().toString('utf8');
        const coreParsed = await this.parser.parseStringPromise(coreXml);
        metadata.core = this.flattenMetadata(coreParsed.coreProperties || coreParsed);
      } catch (error) {
        metadata.core = { error: 'Unable to parse core metadata.' };
      }
    }

    if (appEntry) {
      try {
        const appXml = appEntry.getData().toString('utf8');
        const appParsed = await this.parser.parseStringPromise(appXml);
        metadata.app = this.flattenMetadata(appParsed.Properties || appParsed);
      } catch (error) {
        metadata.app = { error: 'Unable to parse app metadata.' };
      }
    }

    return metadata;
  }

  flattenMetadata(node) {
    if (!node || typeof node !== 'object') {
      return node;
    }

    const result = {};

    Object.entries(node).forEach(([key, value]) => {
      if (key === 'xmlns' || key === 'xmlns:cp' || key === 'xmlns:dc' || key === 'xmlns:dcmitype' || key === 'xmlns:xsi') {
        return;
      }

      if (typeof value === 'object' && value.hasOwnProperty('_')) {
        result[key] = value['_'];
      } else if (typeof value === 'string') {
        result[key] = value;
      } else {
        result[key] = value;
      }
    });

    return result;
  }
}

module.exports = PptReader;
