'use strict';

/**
 * Export utility functions
 */
class ExportUtils {
  /**
   * Export presentation to JSON
   * @param {Presentation} presentation - The presentation to export
   * @returns {string} JSON representation of the presentation
   */
  static toJson(presentation) {
    return JSON.stringify({
      title: presentation.title,
      slides: presentation.slides.map(slide => ({
        title: slide.title,
        elements: slide.elements,
        layout: slide.layout
      })),
      theme: presentation.theme
    }, null, 2);
  }

  /**
   * Export presentation to XML (simplified)
   * @param {Presentation} presentation - The presentation to export
   * @returns {string} XML representation of the presentation
   */
  static toXml(presentation) {
    let xml = '<?xml version="1.0" encoding="UTF-8"?>\n';
    xml += '<presentation>\n';
    xml += `  <title>${presentation.title}</title>\n`;
    xml += '  <slides>\n';
    
    presentation.slides.forEach(slide => {
      xml += '    <slide>\n';
      xml += `      <title>${slide.title}</title>\n`;
      xml += '      <elements>\n';
      
      slide.elements.forEach(element => {
        xml += `        <element type="${element.type}">\n`;
        // Add element-specific properties here
        xml += '        </element>\n';
      });
      
      xml += '      </elements>\n';
      xml += '    </slide>\n';
    });
    
    xml += '  </slides>\n';
    xml += '</presentation>';
    
    return xml;
  }
}

module.exports = ExportUtils;