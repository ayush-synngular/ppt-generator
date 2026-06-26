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

    // Bug 1 fix: always queue the replacement regardless of what the in-memory
    // slide model reports. The actual XML edit (_applyTextReplacementsToZip)
    // scans every slide's raw XML directly and does not use slideNumber at all,
    // so gating on _replaceTextInSlideObject silently dropped replacements
    // whenever the in-memory model was unpopulated — causing save() to see an
    // empty this.replacements, skip all edits, and re-zip the original untouched.
    this.getSlides().forEach((slide) => {
      this._replaceTextInSlideObject(slide, oldText, newText); // keep in-memory model in sync
      this.replacements.push({ slideNumber: slide.slideNumber, oldText, newText });
    });

    // If there are no slides in the parsed model at all, still queue once so
    // _applyTextReplacementsToZip can scan the raw XML entries.
    if (this.getSlides().length === 0) {
      this.replacements.push({ slideNumber: null, oldText, newText });
    }
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
      this.replacements.length > 0 || this.newSlides.length > 0 || this.newTextBoxes.length > 0;

    const zip = new AdmZip(originalFile);

    if (!hasPendingWork) {
      // Nothing queued — just re-save an identical copy.
      zip.writeZip(filePath);
      return;
    }

    if (this.replacements.length > 0) {
      this._applyTextReplacementsToZip(zip);
    }

    if (this.newTextBoxes.length > 0) {
      this._applyTextBoxAdditionsToZip(zip);
    }

    if (this.newSlides.length > 0) {
      this._applyNewSlidesToZip(zip);
    }

    zip.writeZip(filePath);
  }

  // -------------------------------------------------------------------
  // Internal: text replacement (Bug 1 + Bug 2 fixed)
  // -------------------------------------------------------------------
  _applyTextReplacementsToZip(zip) {
    const slideEntries = zip
      .getEntries()
      .filter((entry) => /^ppt\/slides\/slide\d+\.xml$/.test(entry.entryName));

    const escapeRegExp = (value) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

    /**
     * Bug 2 fix: replaceAcrossRuns
     *
     * The original code only matched a search phrase when the entire string
     * sat inside a single <a:t> tag:
     *   <a:t>Artificial Intelligence</a:t>
     *
     * PowerPoint routinely splits a logical phrase across consecutive runs, e.g.:
     *   <a:t>Artificial</a:t><a:t> Intelligence</a:t>
     * In that case the original regex could never match, so no replacement was
     * ever written to the XML.
     *
     * This function tries the fast single-run path first. If that finds nothing
     * it falls back to a sliding-window scan over consecutive <a:t> runs,
     * concatenates their text, and checks for a match. When found it puts the
     * full replacement in the first run of the window and empties the rest.
     *
     * Formatting caveat: only <a:t> text content is changed. The enclosing
     * <a:r>/<a:rPr> markup is left intact, so the merged text visually inherits
     * the first run's formatting. That is an accepted tradeoff of operating at
     * the text-content level rather than with a full XML parser.
     */
    const replaceAcrossRuns = (xml, oldText, newText) => {
      const escapedOld = escapeRegExp(oldText);
      const safeNew = escapeXml(newText);

      // --- Fast path: entire phrase in one <a:t> tag ---
      const singlePattern = new RegExp(`(<a:t>)${escapedOld}(</a:t>)`, 'g');
      if (singlePattern.test(xml)) {
        return xml.replace(singlePattern, `$1${safeNew}$2`);
      }

      // --- Fallback: phrase split across consecutive <a:t> runs ---
      const runPattern = /<a:t>([^<]*)<\/a:t>/g;
      let match;
      const runs = [];

      while ((match = runPattern.exec(xml)) !== null) {
        runs.push({ index: match.index, full: match[0], text: match[1] });
      }

      for (let i = 0; i < runs.length; i++) {
        let combined = '';
        for (let j = i; j < runs.length; j++) {
          combined += runs[j].text;

          if (combined === oldText) {
            // Runs i..j together form the target phrase.
            // Put the replacement in the first run; empty the rest.
            let result = xml;
            let offset = 0;

            for (let k = i; k <= j; k++) {
              const replacement =
                k === i ? `<a:t>${safeNew}</a:t>` : `<a:t></a:t>`;
              const orig = runs[k].full;
              const pos = runs[k].index + offset;
              result =
                result.slice(0, pos) +
                replacement +
                result.slice(pos + orig.length);
              offset += replacement.length - orig.length;
            }
            return result;
          }

          // Overshot — no point extending this window further.
          if (combined.length > oldText.length) break;
        }
      }

      return xml; // no match found in either path
    };

    for (const entry of slideEntries) {
      let xml = entry.getData().toString('utf8');
      this.replacements.forEach(({ oldText, newText }) => {
        xml = replaceAcrossRuns(xml, oldText, newText);
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

  _applyTextReplacements(parsedSlide) {
    const target = parsedSlide;
    const oldNewPairs = this.replacements;

    const replaceTextInNode = (node) => {
      if (!node || typeof node !== 'object') {
        return;
      }

      if (node['a:t'] && typeof node['a:t'] === 'string') {
        oldNewPairs.forEach(({ oldText, newText }) => {
          if (node['a:t'] === oldText) {
            node['a:t'] = newText;
          }
        });
      }

      Object.values(node).forEach((child) => {
        if (Array.isArray(child)) {
          child.forEach(replaceTextInNode);
        } else if (typeof child === 'object') {
          replaceTextInNode(child);
        }
      });
    };

    replaceTextInNode(target);
    return target;
  }
}

module.exports = PptModifier;