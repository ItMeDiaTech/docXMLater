/**
 * ImageRun - A run that contains an image (drawing)
 * Extends Run class for type-safe paragraph content
 *
 * This is a specialized Run that contains a drawing instead of text.
 * It generates proper w:r (run) XML with w:drawing child element.
 */

import { Run } from './Run.js';
import { Image } from './Image.js';
import { XMLElement, XMLBuilder } from '../xml/XMLBuilder.js';

/**
 * ImageRun - A run containing an embedded image
 *
 * In WordprocessingML, images are embedded in runs as drawing elements:
 * <w:r>
 *   <w:drawing>
 *     <wp:inline>
 *       ... image data ...
 *     </wp:inline>
 *   </w:drawing>
 * </w:r>
 */
export class ImageRun extends Run {
  private imageElement: Image;
  private _rawRunXml?: string;

  /**
   * Creates a new image run
   * @param image The image to embed in this run
   */
  constructor(image: Image) {
    // Call parent constructor with empty text
    // The text is irrelevant for image runs
    super('');
    this.imageElement = image;
  }

  setRawRunXml(xml: string): void {
    this._rawRunXml = xml;
  }

  getRawRunXml(): string | undefined {
    return this._rawRunXml;
  }

  /**
   * Gets the image element
   * @returns Image instance
   */
  getImageElement(): Image {
    return this.imageElement;
  }

  /**
   * Override toXML to generate image-specific XML
   * Generates a w:r element containing w:drawing instead of w:t
   * @returns XMLElement with w:r containing w:drawing
   */
  toXML(): XMLElement {
    if (this._rawRunXml) {
      // No live mutation since parse — emit the captured run XML verbatim (round-trip).
      if (!this.imageElement.isMutated()) {
        return { name: '__rawXml', rawXml: this._rawRunXml };
      }
      // The live image was mutated (e.g. setBorder/setSize). Splice the regenerated
      // <w:drawing> into the captured run XML so the change applies while preserving the
      // run's rPr and other captured details — a full regenerate would drop the parsed
      // rPr (not modeled on revision-nested image runs) and can clip the image in Word.
      const drawingXml = XMLBuilder.elementToString(this.imageElement.toXML());
      const spliced = this._rawRunXml.replace(/<w:drawing\b[\s\S]*<\/w:drawing>/, drawingXml);
      if (spliced !== this._rawRunXml) {
        return { name: '__rawXml', rawXml: spliced };
      }
      // No <w:drawing> found to splice (unexpected) — fall through to full regeneration.
    }
    const drawing = this.imageElement.toXML();
    const children: XMLElement[] = [];
    // Per ECMA-376 §17.3.2.28, w:rPr (if present) must precede the run's
    // content. For runs containing an inline drawing, the rPr's w:rFonts
    // affects line metrics in Word; dropping it shifts the baseline and
    // can let images (with shadow effectExtent) overflow their containing
    // cell. Emit rPr whenever the run carries formatting.
    const rPr = Run.generateRunPropertiesXML(this.getFormatting());
    if (rPr) children.push(rPr);
    children.push(drawing);
    return {
      name: 'w:r',
      children,
    };
  }
}
