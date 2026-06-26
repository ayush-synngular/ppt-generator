'use strict';

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

/**
 * PptReader responsible for unpacking a compiled PPTX presentation 
 * and building a clean, non-fragmented structured layout workspace.
 */
class PptReader {
  constructor() {
    this.utils = null;
  }

  /**
   * Reads and parses any PPTX file structure, merging text run pieces and stripping formatting tags.
   * @param {string} filePath - Absolute path to the presentation asset
   * @returns {Promise<Object>} The compiled import data model mapping slides dynamically
   */
  async read(filePath) {
    if (!filePath || typeof filePath !== 'string') {
      throw new Error('PptReader.read requires a valid file path string.');
    }

    const absolutePath = path.resolve(filePath);
    if (!fs.existsSync(absolutePath)) {
      throw new Error(`Target presentation file asset not found at path: ${absolutePath}`);
    }

    // Create a unique temporary directory to extract presentation zip files safely
    const tmpDirName = `pptx_unpack_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
    const tmpDir = path.join(process.cwd(), tmpDirName);
    fs.mkdirSync(tmpDir, { recursive: true });

    const slides = [];

    try {
      // Unpack the file using standard system extraction tools
      if (process.platform === 'win32') {
        execSync(`powershell -Command "Expand-Archive -Path '${absolutePath}' -DestinationPath '${tmpDir}'"`, { stdio: 'ignore' });
      } else {
        execSync(`unzip -q "${absolutePath}" -d "${tmpDir}"`, { stdio: 'ignore' });
      }

      const slidesPath = path.join(tmpDir, 'ppt', 'slides');
      if (!fs.existsSync(slidesPath)) {
        throw new Error('Invalid presentation structure: missing slide definitions directory.');
      }

      // Read every individual slide document file inside the zip directory stream
      const slideFiles = fs.readdirSync(slidesPath)
        .filter(file => file.startsWith('slide') && file.endsWith('.xml'))
        .sort((a, b) => {
          const numA = parseInt(a.replace(/[^0-9]/g, ''), 10);
          const numB = parseInt(b.replace(/[^0-9]/g, ''), 10);
          return numA - numB;
        });

      slideFiles.forEach((slideFile) => {
        const slideXmlPath = path.join(slidesPath, slideFile);
        const xmlContent = fs.readFileSync(slideXmlPath, 'utf8');

        const slideShapes = [];

        // 1. PARAGRAPH EXTRACTOR: Isolate full text blocks (<a:p>...</a:p>) to preserve sentence integrity
        const paragraphRegex = /<a:p[\s\S]*?>([\s\S]*?)<\/a:p>/g;
        let paragraphMatch;

        while ((paragraphMatch = paragraphRegex.exec(xmlContent)) !== null) {
          const paragraphInnerXml = paragraphMatch[1];
          
          // Gather all text fragments (<a:t>...</a:t>) residing inside this specific paragraph block
          const textRunRegex = /<a:t[^>]*>([\s\S]*?)<\/a:t>/g;
          let textRunMatch;
          let compiledParagraphText = '';

          while ((textRunMatch = textRunRegex.exec(paragraphInnerXml)) !== null) {
            compiledParagraphText += textRunMatch[1];
          }

          // FIX: Deeply sanitize the extracted text block by removing any residual embedded inline xml layout tags
          compiledParagraphText = compiledParagraphText
            .replace(/<[^>]+>/g, '') // Strips out any leaking structural markers completely
            .replace(/&apos;/g, "'")
            .replace(/&quot;/g, '"')
            .replace(/&amp;/g, '&')
            .replace(/\s+/g, ' ')
            .trim();

          // Only commit meaningful text elements to our workspace model
          if (compiledParagraphText.length > 0) {
            slideShapes.push({
              text: compiledParagraphText,
              // Establish baseline heuristics for layout metrics mapping
              fontSize: slideShapes.length === 0 ? 28 : 16,
              bold: slideShapes.length === 0,
              y: slideShapes.length === 0 ? 0.6 : 2.0
            });
          }
        }

        // 2. CONTEXTUAL SLIDE HEADER INTERCEPTOR
        let detectedTitle = null;
        if (slideShapes.length > 0) {
          // Promote the first paragraph to slide title if it fits header style parameters
          const primaryBlock = slideShapes[0];
          if (primaryBlock.text.length < 100) { 
            detectedTitle = primaryBlock.text;
          }
        }

        const currentSlideNumber = slides.length + 1;
        const finalCalculatedTitle = detectedTitle || `Slide ${currentSlideNumber}`;

        slides.push({
          slideNumber: currentSlideNumber,
          title: finalCalculatedTitle,
          shapes: slideShapes
        });
      });

      return {
        filePath: absolutePath,
        slides
      };

    } catch (err) {
      throw new Error(`Failed parsing target presentation structures dynamically: ${err.message}`);
    } finally {
      // Cleanup workspace directory cache paths
      if (fs.existsSync(tmpDir)) {
        try {
          fs.rmSync(tmpDir, { recursive: true, force: true });
        } catch (e) {
          // Guard silent cleanup failures
        }
      }
    }
  }
}

module.exports = PptReader;