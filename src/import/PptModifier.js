'use strict';

const fs = require('fs');

const EMU_PER_INCH = 914400;
const DEFAULT_SLIDE_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

function inchesToEmu(value) {
  return Math.round(Number(value) * EMU_PER_INCH);
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function normalizeAlign(align) {
  switch (align) {
    case 'center':
      return 'ctr';
    case 'right':
      return 'r';
    case 'left':
    default:
      return 'l';
  }
}

function normalizeColor(color) {
  if (!color) return '000000';
  return String(color).replace(/^#/, '').toUpperCase();
}

class PptModifier {
  constructor(presentation) {
    if (!presentation || typeof presentation !== 'object') {
      throw new Error('PptModifier requires a valid presentation object.');
    }

    this.presentation = presentation;
    this.replacements = [];
    this.newSlides = [];
    this.newTextBoxes = [];
  }

  getSlides() {
    return Array.isArray(this.presentation.slides) ? this.presentation.slides : [];
  }

  getSlide(slideNumber) {
    if (!Number.isInteger(slideNumber) || slideNumber < 1) {
      return undefined;
    }

    return this.getSlides().find((slide) => slide.slideNumber === slideNumber);
  }

  replaceText(oldText, newText) {
    if (typeof oldText !== 'string' || typeof newText !== 'string') {
      return;
    }

    this.getSlides().forEach((slide) => {
      const replacements = this._replaceTextInSlideObject(slide, oldText, newText);
      if (replacements > 0) {
        this.replacements.push({ slideNumber: slide.slideNumber, oldText, newText });
      }
    });
  }

  replaceTextInSlide(slideNumber, oldText, newText) {
    if (typeof oldText !== 'string' || typeof newText !== 'string') {
      return;
    }

    const slide = this.getSlide(slideNumber);
    if (!slide) {
      return;
    }

    const replacements = this._replaceTextInSlideObject(slide, oldText, newText);
    if (replacements > 0) {
      this.replacements.push({ slideNumber, oldText, newText });
    }
  }

  /**
   * Queue a brand-new slide containing a single text box. Persisted to disk
   * the next time save() is called. Returns the slide number this new slide
   * is expected to occupy (based on slides already known about, including
   * any previously queued additions).
   *
   * opts: { x, y, w, h, fontFace, fontSize, color, bold, align }
   * (x/y/w/h in inches; fontSize in points; color as hex, with or without '#')
   */
  addSlide(text, opts = {}) {
    const slideNumber = this.getSlides().length + this.newSlides.length + 1;
    this.newSlides.push({ slideNumber, text: text === undefined ? '' : String(text), opts });
    return slideNumber;
  }

  /**
   * Queue a new text box to be injected into an *existing* slide.
   * Only applies to slide numbers that already exist in the opened file
   * (not to slides queued via addSlide() in the same session).
   */
  addTextBoxToSlide(slideNumber, text, opts = {}) {
    if (!Number.isInteger(slideNumber) || slideNumber < 1) {
      throw new Error('addTextBoxToSlide requires a valid slideNumber.');
    }
    this.newTextBoxes.push({ slideNumber, text: text === undefined ? '' : String(text), opts });
  }

  findText(text) {
    if (typeof text !== 'string') {
      return [];
    }

    const matches = [];
    this.getSlides().forEach((slide) => {
      if (Array.isArray(slide.text)) {
        slide.text.forEach((segment) => {
          if (segment === text) {
            matches.push({ slideNumber: slide.slideNumber, text: segment });
          }
        });
      }

      if (Array.isArray(slide.shapes)) {
        slide.shapes.forEach((shape) => {
          if (shape && typeof shape.text === 'string' && shape.text === text) {
            matches.push({ slideNumber: slide.slideNumber, text: shape.text });
          }
        });
      }
    });

    return matches;
  }

  toJSON() {
    if (typeof this.presentation.toJSON === 'function') {
      return this.presentation.toJSON();
    }

    return {
      ...this.presentation,
      slides: this.getSlides().map((slide) => (typeof slide.toJSON === 'function' ? slide.toJSON() : slide))
    };
  }

  saveJSON(filePath) {
    if (typeof filePath !== 'string' || filePath.trim() === '') {
      throw new Error('A valid file path is required to save JSON.');
    }

    const json = this.toJSON();
    fs.writeFileSync(filePath, JSON.stringify(json, null, 2), 'utf8');
  }



  // -------------------------------------------------------------------
  // Internal: text replacement (existing behavior, unchanged)
  // -------------------------------------------------------------------
  _applyTextReplacementsToZip(zip) {
    const slideEntries = zip
      .getEntries()
      .filter((entry) => /^ppt\/slides\/slide\d+\.xml$/.test(entry.entryName));

    const escapeRegExp = (value) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

    for (const entry of slideEntries) {
      let xml = entry.getData().toString('utf8');
      this.replacements.forEach(({ oldText, newText }) => {
        const pattern = new RegExp(`<a:t>${escapeRegExp(oldText)}</a:t>`, 'g');
        xml = xml.replace(pattern, `<a:t>${escapeXml(newText)}</a:t>`);
      });
      zip.updateFile(entry.entryName, Buffer.from(xml, 'utf8'));
    }
  }

  // -------------------------------------------------------------------
  // Internal: inject a text box into an existing slide's XML
  // -------------------------------------------------------------------
  _applyTextBoxAdditionsToZip(zip) {
    const bySlide = new Map();
    this.newTextBoxes.forEach((item) => {
      if (!bySlide.has(item.slideNumber)) bySlide.set(item.slideNumber, []);
      bySlide.get(item.slideNumber).push(item);
    });

    bySlide.forEach((items, slideNumber) => {
      const entryName = `ppt/slides/slide${slideNumber}.xml`;
      const entry = zip.getEntry(entryName);
      if (!entry) {
        throw new Error(`Cannot add text box: slide ${slideNumber} does not exist in this presentation.`);
      }

      let xml = entry.getData().toString('utf8');
      const closeTag = '</p:spTree>';
      const idx = xml.lastIndexOf(closeTag);
      if (idx === -1) {
        throw new Error(`Slide ${slideNumber} XML is missing <p:spTree> — cannot inject text box.`);
      }

      let shapeId = this._nextShapeId(xml);
      let injected = '';
      items.forEach((item) => {
        injected += this._buildTextBoxShapeXml(shapeId, item.text, item.opts);
        shapeId += 1;
      });

      xml = xml.slice(0, idx) + injected + xml.slice(idx);
      zip.updateFile(entryName, Buffer.from(xml, 'utf8'));
    });
  }

  // -------------------------------------------------------------------
  // Internal: create brand-new slide(s) and wire them into the package
  // -------------------------------------------------------------------
  _applyNewSlidesToZip(zip) {
    const slideEntries = zip
      .getEntries()
      .filter((entry) => /^ppt\/slides\/slide\d+\.xml$/.test(entry.entryName));

    let maxSlideNum = 0;
    slideEntries.forEach((entry) => {
      const match = entry.entryName.match(/^ppt\/slides\/slide(\d+)\.xml$/);
      if (match) {
        maxSlideNum = Math.max(maxSlideNum, parseInt(match[1], 10));
      }
    });

    if (maxSlideNum === 0) {
      throw new Error('No existing slides found to base a new slide on — cannot determine layout reference.');
    }

    const layoutTarget = this._getSlideLayoutTarget(zip, maxSlideNum);

    this.newSlides.forEach(({ text, opts }) => {
      maxSlideNum += 1;
      const newSlideNum = maxSlideNum;

      const slideXml = this._buildSlideXml(text, opts);
      const relsXml = this._buildSlideRelsXml(layoutTarget);

      zip.addFile(`ppt/slides/slide${newSlideNum}.xml`, Buffer.from(slideXml, 'utf8'));
      zip.addFile(`ppt/slides/_rels/slide${newSlideNum}.xml.rels`, Buffer.from(relsXml, 'utf8'));

      this._registerSlideInContentTypes(zip, newSlideNum);
      const presRelId = this._registerSlideRelationship(zip, newSlideNum);
      this._registerSlideInPresentationXml(zip, presRelId);
    });
  }

  _getSlideLayoutTarget(zip, slideNumber) {
    const relsEntry = zip.getEntry(`ppt/slides/_rels/slide${slideNumber}.xml.rels`);
    if (!relsEntry) {
      return '../slideLayouts/slideLayout1.xml';
    }
    const relsXml = relsEntry.getData().toString('utf8');
    const match = relsXml.match(/Target="([^"]*slideLayout[^"]*\.xml)"/);
    return match ? match[1] : '../slideLayouts/slideLayout1.xml';
  }

  _nextShapeId(slideXml) {
    const ids = [...slideXml.matchAll(/<p:cNvPr\s+id="(\d+)"/g)].map((m) => parseInt(m[1], 10));
    return ids.length ? Math.max(...ids) + 1 : 2;
  }

  _buildTextBoxShapeXml(shapeId, text, opts = {}) {
    const x = inchesToEmu(opts.x !== undefined ? opts.x : 0.5);
    const y = inchesToEmu(opts.y !== undefined ? opts.y : 0.5);
    const w = inchesToEmu(opts.w !== undefined ? opts.w : 9);
    const h = inchesToEmu(opts.h !== undefined ? opts.h : 1);
    const fontSize = Math.round((opts.fontSize !== undefined ? opts.fontSize : 24) * 100);
    const color = normalizeColor(opts.color || '000000');
    const align = normalizeAlign(opts.align);
    const bold = opts.bold ? '1' : '0';
    const fontFace = escapeXml(opts.fontFace || 'Arial');
    const safeText = escapeXml(text);

    return (
      `<p:sp><p:nvSpPr><p:cNvPr id="${shapeId}" name="TextBox ${shapeId}"/>` +
      `<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
      `<p:txBody><a:bodyPr wrap="square"><a:spAutoFit/></a:bodyPr><a:lstStyle/>` +
      `<a:p><a:pPr algn="${align}"/><a:r>` +
      `<a:rPr lang="en-US" sz="${fontSize}" b="${bold}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` +
      `<a:latin typeface="${fontFace}"/></a:rPr><a:t>${safeText}</a:t></a:r></a:p>` +
      `</p:txBody></p:sp>`
    );
  }

  _buildSlideXml(text, opts = {}) {
    const body = this._buildTextBoxShapeXml(2, text, opts);
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<p:sld ${DEFAULT_SLIDE_NS}><p:cSld><p:spTree>` +
      `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
      `${body}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`
    );
  }

  _buildSlideRelsXml(layoutTarget) {
    return (
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="${layoutTarget}"/>` +
      `</Relationships>`
    );
  }

  _registerSlideInContentTypes(zip, slideNumber) {
    const entry = zip.getEntry('[Content_Types].xml');
    if (!entry) {
      throw new Error('[Content_Types].xml not found in package.');
    }
    let xml = entry.getData().toString('utf8');
    const override =
      `<Override PartName="/ppt/slides/slide${slideNumber}.xml" ` +
      `ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;

    if (xml.includes(`/ppt/slides/slide${slideNumber}.xml"`)) {
      return; // already registered, avoid duplicate
    }

    xml = xml.replace('</Types>', `${override}</Types>`);
    zip.updateFile('[Content_Types].xml', Buffer.from(xml, 'utf8'));
  }

  _registerSlideRelationship(zip, slideNumber) {
    const entryName = 'ppt/_rels/presentation.xml.rels';
    const entry = zip.getEntry(entryName);
    if (!entry) {
      throw new Error(`${entryName} not found in package.`);
    }
    let xml = entry.getData().toString('utf8');

    const existingIds = [...xml.matchAll(/Id="rId(\d+)"/g)].map((m) => parseInt(m[1], 10));
    const nextId = existingIds.length ? Math.max(...existingIds) + 1 : 1;
    const relId = `rId${nextId}`;

    const relationship =
      `<Relationship Id="${relId}" ` +
      `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" ` +
      `Target="slides/slide${slideNumber}.xml"/>`;

    xml = xml.replace('</Relationships>', `${relationship}</Relationships>`);
    zip.updateFile(entryName, Buffer.from(xml, 'utf8'));

    return relId;
  }

  _registerSlideInPresentationXml(zip, relId) {
    const entryName = 'ppt/presentation.xml';
    const entry = zip.getEntry(entryName);
    if (!entry) {
      throw new Error(`${entryName} not found in package.`);
    }
    let xml = entry.getData().toString('utf8');

    const existingSldIds = [...xml.matchAll(/<p:sldId\s+id="(\d+)"/g)].map((m) => parseInt(m[1], 10));
    const nextSldId = existingSldIds.length ? Math.max(...existingSldIds) + 1 : 256;

    const sldIdTag = `<p:sldId id="${nextSldId}" r:id="${relId}"/>`;

    if (xml.includes('</p:sldIdLst>')) {
      xml = xml.replace('</p:sldIdLst>', `${sldIdTag}</p:sldIdLst>`);
    } else {
      // No slide list exists yet (unusual for a non-empty deck) — create one
      // right after <p:sldMasterIdLst>...</p:sldMasterIdLst> if present,
      // otherwise right after the opening <p:presentation ...> tag.
      const sldIdLst = `<p:sldIdLst>${sldIdTag}</p:sldIdLst>`;
      if (xml.includes('</p:sldMasterIdLst>')) {
        xml = xml.replace('</p:sldMasterIdLst>', `</p:sldMasterIdLst>${sldIdLst}`);
      } else {
        xml = xml.replace(/(<p:presentation[^>]*>)/, `$1${sldIdLst}`);
      }
    }

    zip.updateFile(entryName, Buffer.from(xml, 'utf8'));
  }

  // -------------------------------------------------------------------
  // Internal: in-memory model helpers (used by replaceText / findText)
  // -------------------------------------------------------------------
  _replaceTextInSlideObject(slide, oldText, newText) {
    if (!slide || typeof slide !== 'object') {
      return 0;
    }

    let replacements = 0;

    if (Array.isArray(slide.text)) {
      slide.text = slide.text.map((segment) => {
        if (segment === oldText) {
          replacements += 1;
          return newText;
        }
        return segment;
      });
    }

    if (Array.isArray(slide.shapes)) {
      slide.shapes = slide.shapes.map((shape) => {
        if (shape && typeof shape.text === 'string' && shape.text === oldText) {
          replacements += 1;
          return { ...shape, text: newText };
        }
        return shape;
      });
    }

    return replacements;
  }


  // -------------------------------------------------------------------
  // Public: slide count
  // -------------------------------------------------------------------

  /**
   * Returns the number of slides currently in the opened presentation.
   * Does not count slides queued via addSlide() that haven't been saved yet.
   */
  getSlideCount() {
    return this.getSlides().length;
  }

  // -------------------------------------------------------------------
  // Public: delete a slide
  // Queue a slide number for deletion. Applied on the next save().
  // slideNumber is the 1-based position as reported by getSlides().
  // -------------------------------------------------------------------

  deleteSlide(slideNumber) {
    if (!Number.isInteger(slideNumber) || slideNumber < 1) {
      throw new Error('deleteSlide requires a valid slideNumber (integer >= 1).');
    }
    if (!this.slideDeletions) this.slideDeletions = [];
    if (!this.slideDeletions.includes(slideNumber)) {
      this.slideDeletions.push(slideNumber);
    }
  }

  // -------------------------------------------------------------------
  // Public: duplicate an existing slide
  // Copies slideNumber and appends it as a new last slide on save().
  // -------------------------------------------------------------------

  duplicateSlide(slideNumber) {
    if (!Number.isInteger(slideNumber) || slideNumber < 1) {
      throw new Error('duplicateSlide requires a valid slideNumber (integer >= 1).');
    }
    if (!this.slideDuplications) this.slideDuplications = [];
    this.slideDuplications.push(slideNumber);
  }

  // -------------------------------------------------------------------
  // Public: move a slide to a new position
  // Both fromSlideNumber and toPosition are 1-based.
  // Applied on the next save().
  // -------------------------------------------------------------------

  moveSlide(fromSlideNumber, toPosition) {
    if (!Number.isInteger(fromSlideNumber) || fromSlideNumber < 1) {
      throw new Error('moveSlide: fromSlideNumber must be a valid integer >= 1.');
    }
    if (!Number.isInteger(toPosition) || toPosition < 1) {
      throw new Error('moveSlide: toPosition must be a valid integer >= 1.');
    }
    if (!this.slideMoves) this.slideMoves = [];
    this.slideMoves.push({ from: fromSlideNumber, to: toPosition });
  }

  // -------------------------------------------------------------------
  // Public: partial / case-insensitive text replacement across all slides
  // flags: standard RegExp flags string, e.g. 'gi' (default: 'g')
  // -------------------------------------------------------------------

  replaceTextGlobal(oldText, newText, flags = 'g') {
    if (typeof oldText !== 'string' || typeof newText !== 'string') {
      return;
    }
    if (!this.globalReplacements) this.globalReplacements = [];
    this.globalReplacements.push({ oldText, newText, flags });
  }

  // -------------------------------------------------------------------
  // Public: set background color of an existing slide
  // color: hex string with or without '#', e.g. '#FF0000' or 'FF0000'
  // Applied on the next save().
  // -------------------------------------------------------------------

  setSlideBackground(slideNumber, color) {
    if (!Number.isInteger(slideNumber) || slideNumber < 1) {
      throw new Error('setSlideBackground requires a valid slideNumber.');
    }
    if (typeof color !== 'string') {
      throw new Error('setSlideBackground requires a color string.');
    }
    if (!this.backgroundChanges) this.backgroundChanges = [];
    this.backgroundChanges.push({ slideNumber, color: normalizeColor(color) });
  }

  // -------------------------------------------------------------------
  // Public: add an image to an existing slide
  // imagePath: absolute or relative path to image file on disk (png/jpg/gif/bmp)
  // opts: { x, y, w, h } in inches
  // Applied on the next save().
  // -------------------------------------------------------------------

  addImageToSlide(slideNumber, imagePath, opts = {}) {
    if (!Number.isInteger(slideNumber) || slideNumber < 1) {
      throw new Error('addImageToSlide requires a valid slideNumber.');
    }
    if (typeof imagePath !== 'string' || !fs.existsSync(imagePath)) {
      throw new Error(`addImageToSlide: image file not found at "${imagePath}".`);
    }
    if (!this.imageAdditions) this.imageAdditions = [];
    this.imageAdditions.push({ slideNumber, imagePath, opts });
  }

  // -------------------------------------------------------------------
  // Override save() to wire in all new queued operations
  // -------------------------------------------------------------------

  async save(filePath) {
    if (typeof filePath !== 'string' || filePath.trim() === '') {
      throw new Error('A valid file path is required to save PPTX.');
    }

    const AdmZip = require('adm-zip');

    const originalFile = this.presentation.filePath;
    if (!originalFile || typeof originalFile !== 'string' || !fs.existsSync(originalFile)) {
      throw new Error('Original PPTX file is required to preserve layout and media.');
    }

    const hasPendingWork =
      this.replacements.length > 0 ||
      this.newSlides.length > 0 ||
      this.newTextBoxes.length > 0 ||
      (this.slideDeletions && this.slideDeletions.length > 0) ||
      (this.slideDuplications && this.slideDuplications.length > 0) ||
      (this.slideMoves && this.slideMoves.length > 0) ||
      (this.globalReplacements && this.globalReplacements.length > 0) ||
      (this.backgroundChanges && this.backgroundChanges.length > 0) ||
      (this.imageAdditions && this.imageAdditions.length > 0);

    const zip = new AdmZip(originalFile);

    if (!hasPendingWork) {
      zip.writeZip(filePath);
      return;
    }

    // Order matters: deletions first so slide numbers stay consistent for
    // operations that reference them, then moves, then additions.
    if (this.slideDeletions && this.slideDeletions.length > 0) {
      this._applyDeletionsToZip(zip);
    }

    if (this.slideDuplications && this.slideDuplications.length > 0) {
      this._applyDuplicationsToZip(zip);
    }

    if (this.slideMoves && this.slideMoves.length > 0) {
      this._applyMovesToZip(zip);
    }

    if (this.replacements.length > 0) {
      this._applyTextReplacementsToZip(zip);
    }

    if (this.globalReplacements && this.globalReplacements.length > 0) {
      this._applyGlobalReplacementsToZip(zip);
    }

    if (this.backgroundChanges && this.backgroundChanges.length > 0) {
      this._applyBackgroundChangesToZip(zip);
    }

    if (this.newTextBoxes.length > 0) {
      this._applyTextBoxAdditionsToZip(zip);
    }

    if (this.newSlides.length > 0) {
      this._applyNewSlidesToZip(zip);
    }

    if (this.imageAdditions && this.imageAdditions.length > 0) {
      this._applyImageAdditionsToZip(zip);
    }

    zip.writeZip(filePath);
  }

  // -------------------------------------------------------------------
  // Internal: resolve the zip rId that points to a given slide file
  // Returns null if not found.
  // -------------------------------------------------------------------

  _getRelIdForSlide(zip, slideNumber) {
    const entryName = 'ppt/_rels/presentation.xml.rels';
    const entry = zip.getEntry(entryName);
    if (!entry) return null;
    const xml = entry.getData().toString('utf8');
    const pattern = new RegExp(
      `<Relationship[^>]*Id="([^"]+)"[^>]*Target="slides/slide${slideNumber}\\.xml"[^>]*/>`
    );
    const match = xml.match(pattern);
    return match ? match[1] : null;
  }

  // -------------------------------------------------------------------
  // Internal: delete queued slides from the zip
  // -------------------------------------------------------------------

  _applyDeletionsToZip(zip) {
    // Sort descending so removing a higher-numbered slide doesn't shift lower ones
    const toDelete = [...new Set(this.slideDeletions)].sort((a, b) => b - a);

    for (const slideNumber of toDelete) {
      const slideEntry = `ppt/slides/slide${slideNumber}.xml`;
      const relsEntry  = `ppt/slides/_rels/slide${slideNumber}.xml.rels`;

      // 1. Get the rId before we remove the rels file
      const relId = this._getRelIdForSlide(zip, slideNumber);

      // 2. Remove the slide XML and its rels file
      if (zip.getEntry(slideEntry)) zip.deleteFile(slideEntry);
      if (zip.getEntry(relsEntry))  zip.deleteFile(relsEntry);

      // 3. Remove from [Content_Types].xml
      const ctEntry = zip.getEntry('[Content_Types].xml');
      if (ctEntry) {
        let xml = ctEntry.getData().toString('utf8');
        xml = xml.replace(
          new RegExp(
            `<Override[^>]*PartName="/ppt/slides/slide${slideNumber}\\.xml"[^>]*/?>`,
            'g'
          ),
          ''
        );
        zip.updateFile('[Content_Types].xml', Buffer.from(xml, 'utf8'));
      }

      // 4. Remove from ppt/_rels/presentation.xml.rels
      if (relId) {
        const presRelsEntry = zip.getEntry('ppt/_rels/presentation.xml.rels');
        if (presRelsEntry) {
          let xml = presRelsEntry.getData().toString('utf8');
          xml = xml.replace(
            new RegExp(`<Relationship[^>]*Id="${relId}"[^>]*/?>`, 'g'),
            ''
          );
          zip.updateFile('ppt/_rels/presentation.xml.rels', Buffer.from(xml, 'utf8'));
        }
      }

      // 5. Remove from ppt/presentation.xml <p:sldIdLst>
      if (relId) {
        const presEntry = zip.getEntry('ppt/presentation.xml');
        if (presEntry) {
          let xml = presEntry.getData().toString('utf8');
          xml = xml.replace(
            new RegExp(`<p:sldId[^>]*r:id="${relId}"[^>]*/?>`, 'g'),
            ''
          );
          zip.updateFile('ppt/presentation.xml', Buffer.from(xml, 'utf8'));
        }
      }
    }
  }

  // -------------------------------------------------------------------
  // Internal: duplicate queued slides (append copy as last slide)
  // -------------------------------------------------------------------

  _applyDuplicationsToZip(zip) {
    for (const slideNumber of this.slideDuplications) {
      const srcEntry = zip.getEntry(`ppt/slides/slide${slideNumber}.xml`);
      if (!srcEntry) {
        throw new Error(`duplicateSlide: slide ${slideNumber} not found in package.`);
      }

      // Find the next available slide number in the zip
      const existing = zip
        .getEntries()
        .filter((e) => /^ppt\/slides\/slide\d+\.xml$/.test(e.entryName))
        .map((e) => parseInt(e.entryName.match(/slide(\d+)/)[1], 10));
      const newNum = Math.max(...existing) + 1;

      // Copy slide XML
      const slideXml = srcEntry.getData().toString('utf8');
      zip.addFile(`ppt/slides/slide${newNum}.xml`, Buffer.from(slideXml, 'utf8'));

      // Copy rels (update any internal references if needed — layout ref is reused as-is)
      const layoutTarget = this._getSlideLayoutTarget(zip, slideNumber);
      const relsXml = this._buildSlideRelsXml(layoutTarget);
      zip.addFile(`ppt/slides/_rels/slide${newNum}.xml.rels`, Buffer.from(relsXml, 'utf8'));

      // Register in Content_Types, rels, and presentation.xml
      this._registerSlideInContentTypes(zip, newNum);
      const presRelId = this._registerSlideRelationship(zip, newNum);
      this._registerSlideInPresentationXml(zip, presRelId);
    }
  }

  // -------------------------------------------------------------------
  // Internal: reorder slides in presentation.xml <p:sldIdLst>
  // -------------------------------------------------------------------

  _applyMovesToZip(zip) {
    const presEntry = zip.getEntry('ppt/presentation.xml');
    if (!presEntry) return;

    let xml = presEntry.getData().toString('utf8');

    // Extract the full <p:sldIdLst>...</p:sldIdLst> block
    const lstMatch = xml.match(/<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/);
    if (!lstMatch) return;

    // Pull individual <p:sldId .../> entries preserving order
    const itemPattern = /<p:sldId[^/]*(\/?>|>[\s\S]*?<\/p:sldId>)/g;
    const items = [...lstMatch[1].matchAll(/<p:sldId\b[^>]*\/>/g)].map((m) => m[0]);

    if (items.length === 0) return;

    // Apply each move in sequence (1-based → 0-based index)
    for (const { from, to } of this.slideMoves) {
      const fromIdx = from - 1;
      const toIdx   = Math.min(to - 1, items.length - 1);
      if (fromIdx < 0 || fromIdx >= items.length) continue;
      const [moved] = items.splice(fromIdx, 1);
      items.splice(toIdx, 0, moved);
    }

    const newLst = `<p:sldIdLst>${items.join('')}</p:sldIdLst>`;
    xml = xml.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, newLst);
    zip.updateFile('ppt/presentation.xml', Buffer.from(xml, 'utf8'));
  }

  // -------------------------------------------------------------------
  // Internal: partial / regex text replacement across all slide XMLs
  // -------------------------------------------------------------------

  _applyGlobalReplacementsToZip(zip) {
    const slideEntries = zip
      .getEntries()
      .filter((e) => /^ppt\/slides\/slide\d+\.xml$/.test(e.entryName));

    for (const entry of slideEntries) {
      let xml = entry.getData().toString('utf8');
      for (const { oldText, newText, flags } of this.globalReplacements) {
        // Escape for use inside a RegExp, then apply with caller-specified flags
        const escaped = oldText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const pattern = new RegExp(escaped, flags);
        // Replace only inside <a:t>...</a:t> tags to avoid corrupting XML markup
        xml = xml.replace(/<a:t>([^<]*)<\/a:t>/g, (match, inner) => {
          const replaced = inner.replace(pattern, escapeXml(newText));
          return `<a:t>${replaced}</a:t>`;
        });
      }
      zip.updateFile(entry.entryName, Buffer.from(xml, 'utf8'));
    }
  }

  // -------------------------------------------------------------------
  // Internal: set slide background color in slide XML
  // -------------------------------------------------------------------

  _applyBackgroundChangesToZip(zip) {
    for (const { slideNumber, color } of this.backgroundChanges) {
      const entryName = `ppt/slides/slide${slideNumber}.xml`;
      const entry = zip.getEntry(entryName);
      if (!entry) {
        throw new Error(`setSlideBackground: slide ${slideNumber} not found in package.`);
      }

      let xml = entry.getData().toString('utf8');

      // Build the solid fill background element
      const bgFill =
        `<p:bg><p:bgPr>` +
        `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` +
        `<a:effectLst/>` +
        `</p:bgPr></p:bg>`;

      // Remove any existing <p:bg>...</p:bg> block first
      xml = xml.replace(/<p:bg>[\s\S]*?<\/p:bg>/g, '');

      // Inject right after <p:cSld ...> or <p:cSld>
      xml = xml.replace(/(<p:cSld[^>]*>)/, `$1${bgFill}`);

      zip.updateFile(entryName, Buffer.from(xml, 'utf8'));
    }
  }

  // -------------------------------------------------------------------
  // Internal: embed image file into the zip and inject <p:pic> shape
  // -------------------------------------------------------------------

  _applyImageAdditionsToZip(zip) {
    const path = require('path');

    // Detect media type from extension
    const mimeMap = {
      '.png':  'image/png',
      '.jpg':  'image/jpeg',
      '.jpeg': 'image/jpeg',
      '.gif':  'image/gif',
      '.bmp':  'image/bmp',
      '.webp': 'image/webp',
    };

    for (const { slideNumber, imagePath, opts } of this.imageAdditions) {
      const ext      = path.extname(imagePath).toLowerCase();
      const mimeType = mimeMap[ext] || 'image/png';
      const imgData  = fs.readFileSync(imagePath);

      // Find a unique media filename inside the zip
      const existingMedia = zip
        .getEntries()
        .filter((e) => e.entryName.startsWith('ppt/media/'))
        .map((e) => e.entryName);
      const mediaNum = existingMedia.length + 1;
      const mediaName = `image${mediaNum}${ext}`;
      const mediaPath = `ppt/media/${mediaName}`;

      zip.addFile(mediaPath, imgData);

      // Register the image media type in [Content_Types].xml
      const ctEntry = zip.getEntry('[Content_Types].xml');
      if (ctEntry) {
        let ctXml = ctEntry.getData().toString('utf8');
        const defaultExt = ext.replace('.', '');
        if (!ctXml.includes(`Extension="${defaultExt}"`)) {
          ctXml = ctXml.replace(
            '</Types>',
            `<Default Extension="${defaultExt}" ContentType="${mimeType}"/></Types>`
          );
          zip.updateFile('[Content_Types].xml', Buffer.from(ctXml, 'utf8'));
        }
      }

      // Add relationship in slide's .rels file
      const relsEntryName = `ppt/slides/_rels/slide${slideNumber}.xml.rels`;
      let relsEntry = zip.getEntry(relsEntryName);
      let relsXml = relsEntry
        ? relsEntry.getData().toString('utf8')
        : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
          `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

      const existingRIds = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((m) => parseInt(m[1], 10));
      const nextRId = existingRIds.length ? Math.max(...existingRIds) + 1 : 2;
      const imgRelId = `rId${nextRId}`;

      relsXml = relsXml.replace(
        '</Relationships>',
        `<Relationship Id="${imgRelId}" ` +
        `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ` +
        `Target="../media/${mediaName}"/></Relationships>`
      );

      if (relsEntry) {
        zip.updateFile(relsEntryName, Buffer.from(relsXml, 'utf8'));
      } else {
        zip.addFile(relsEntryName, Buffer.from(relsXml, 'utf8'));
      }

      // Build <p:pic> shape and inject into slide XML
      const slideEntryName = `ppt/slides/slide${slideNumber}.xml`;
      const slideEntry = zip.getEntry(slideEntryName);
      if (!slideEntry) {
        throw new Error(`addImageToSlide: slide ${slideNumber} not found in package.`);
      }

      let slideXml = slideEntry.getData().toString('utf8');
      const shapeId = this._nextShapeId(slideXml);

      const x  = inchesToEmu(opts.x !== undefined ? opts.x : 0.5);
      const y  = inchesToEmu(opts.y !== undefined ? opts.y : 0.5);
      const cx = inchesToEmu(opts.w !== undefined ? opts.w : 3);
      const cy = inchesToEmu(opts.h !== undefined ? opts.h : 2);

      const picXml =
        `<p:pic>` +
        `<p:nvPicPr>` +
        `<p:cNvPr id="${shapeId}" name="Image ${shapeId}"/>` +
        `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>` +
        `<p:nvPr/>` +
        `</p:nvPicPr>` +
        `<p:blipFill>` +
        `<a:blip r:embed="${imgRelId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>` +
        `<a:stretch><a:fillRect/></a:stretch>` +
        `</p:blipFill>` +
        `<p:spPr>` +
        `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>` +
        `</p:spPr>` +
        `</p:pic>`;

      const closeTag = '</p:spTree>';
      const idx = slideXml.lastIndexOf(closeTag);
      if (idx === -1) {
        throw new Error(`addImageToSlide: slide ${slideNumber} XML is missing </p:spTree>.`);
      }
      slideXml = slideXml.slice(0, idx) + picXml + slideXml.slice(idx);
      zip.updateFile(slideEntryName, Buffer.from(slideXml, 'utf8'));
    }
  }
}

module.exports = PptModifier;