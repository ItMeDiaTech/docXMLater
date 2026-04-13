import { AlternateContent } from '../../src/elements/AlternateContent';
import { MathParagraph, MathExpression } from '../../src/elements/MathElement';
import { CustomXmlBlock } from '../../src/elements/CustomXml';

describe('AlternateContent', () => {
  const sampleXml =
    '<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing/></mc:Choice></mc:AlternateContent>';

  it('should store raw XML', () => {
    const ac = new AlternateContent(sampleXml);
    expect(ac.getRawXml()).toBe(sampleXml);
  });

  it('should return type identifier', () => {
    const ac = new AlternateContent(sampleXml);
    expect(ac.getType()).toBe('alternateContent');
  });

  it('should produce XMLElement with rawXml for serialization', () => {
    const ac = new AlternateContent(sampleXml);
    const xml = ac.toXML();
    expect(xml.name).toBe('__rawXml');
    expect((xml as any).rawXml).toBe(sampleXml);
  });

  it('should handle empty XML', () => {
    const ac = new AlternateContent('');
    expect(ac.getRawXml()).toBe('');
  });
});

describe('MathParagraph', () => {
  const sampleXml = '<m:oMathPara><m:oMath><m:r><m:t>x²</m:t></m:r></m:oMath></m:oMathPara>';

  it('should store raw XML', () => {
    const mp = new MathParagraph(sampleXml);
    expect(mp.getRawXml()).toBe(sampleXml);
  });

  it('should return type identifier', () => {
    expect(new MathParagraph(sampleXml).getType()).toBe('mathParagraph');
  });

  it('should produce XMLElement for serialization', () => {
    const xml = new MathParagraph(sampleXml).toXML();
    expect(xml.name).toBe('__rawXml');
  });
});

describe('MathExpression', () => {
  const sampleXml = '<m:oMath><m:r><m:t>y=mx+b</m:t></m:r></m:oMath>';

  it('should store raw XML', () => {
    const me = new MathExpression(sampleXml);
    expect(me.getRawXml()).toBe(sampleXml);
  });

  it('should return type identifier', () => {
    expect(new MathExpression(sampleXml).getType()).toBe('mathExpression');
  });

  it('should produce XMLElement for serialization', () => {
    const xml = new MathExpression(sampleXml).toXML();
    expect(xml.name).toBe('__rawXml');
  });
});

describe('CustomXmlBlock', () => {
  const sampleXml = '<w:customXml w:uri="urn:test" w:element="data"><w:p/></w:customXml>';

  it('should store raw XML', () => {
    const cx = new CustomXmlBlock(sampleXml);
    expect(cx.getRawXml()).toBe(sampleXml);
  });

  it('should return type identifier', () => {
    expect(new CustomXmlBlock(sampleXml).getType()).toBe('customXmlBlock');
  });

  it('should produce XMLElement for serialization', () => {
    const xml = new CustomXmlBlock(sampleXml).toXML();
    expect(xml.name).toBe('__rawXml');
    expect((xml as any).rawXml).toBe(sampleXml);
  });
});
