/**
 * ST_OnOff Attribute Parsing Tests
 *
 * Per ECMA-376 Part 1 §17.17.4, the ST_OnOff simple type accepts these
 * literal values for any CT_OnOff-valued attribute:
 *   - "1", "true", "on"   → true
 *   - "0", "false", "off" → false
 *
 * Multiple attribute-based boolean parsers in DocumentParser were missing
 * the "on" / "off" / "false" literals, causing silent misinterpretation
 * of any source document that happened to use those forms:
 *   - Style       w:default, w:customStyle
 *   - Style pPr   w:beforeAutospacing, w:afterAutospacing
 *   - Comment     w:done / w15:done (tracked comment resolution state)
 *   - SDT         w:multiLine, w:showingPlcHdr
 *   - Form field  w:calcOnExit, w:enabled
 *   - Section     border w:shadow, w:frame; columns w:equalWidth, w:sep
 *
 * This suite locks parse behaviour for every ST_OnOff literal against the
 * most impactful of those attributes: style metadata and comment done.
 */

import { Document } from '../../src/core/Document';
import { ZipHandler } from '../../src/zip/ZipHandler';

async function makeDocxWithStyles(stylesXmlFragment: string): Promise<Buffer> {
  const zipHandler = new ZipHandler();

  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`
  );

  zipHandler.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );

  zipHandler.addFile(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
  );

  zipHandler.addFile(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
${stylesXmlFragment}
</w:styles>`
  );

  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>test</w:t></w:r></w:p></w:body>
</w:document>`
  );

  return await zipHandler.toBuffer();
}

async function loadStyle(styleXml: string) {
  const buffer = await makeDocxWithStyles(styleXml);
  const doc = await Document.loadFromBuffer(buffer);
  const sm = doc.getStylesManager();
  const style = sm.getStyle('Test');
  const result = {
    isDefault: style?.getIsDefault(),
    customStyle: (style as unknown as { properties?: { customStyle?: boolean } })?.properties
      ?.customStyle,
    beforeAutospacing: style?.getParagraphFormatting()?.spacing?.beforeAutospacing,
    afterAutospacing: style?.getParagraphFormatting()?.spacing?.afterAutospacing,
  };
  doc.dispose();
  return result;
}

describe('Style w:default and w:customStyle — honour ST_OnOff per ECMA-376', () => {
  const buildStyle = (defAttr: string, customAttr: string) =>
    `<w:style w:type="paragraph" w:styleId="Test" w:default="${defAttr}" w:customStyle="${customAttr}"><w:name w:val="Test"/></w:style>`;

  it('parses w:default="on" as true', async () => {
    const r = await loadStyle(buildStyle('on', '0'));
    expect(r.isDefault).toBe(true);
  });

  it('parses w:default="1" as true', async () => {
    const r = await loadStyle(buildStyle('1', '0'));
    expect(r.isDefault).toBe(true);
  });

  it('parses w:default="true" as true', async () => {
    const r = await loadStyle(buildStyle('true', '0'));
    expect(r.isDefault).toBe(true);
  });

  it('parses w:default="off" as false', async () => {
    const r = await loadStyle(buildStyle('off', '0'));
    expect(r.isDefault).toBe(false);
  });

  it('parses w:default="0" as false', async () => {
    const r = await loadStyle(buildStyle('0', '0'));
    expect(r.isDefault).toBe(false);
  });

  it('parses w:default="false" as false', async () => {
    const r = await loadStyle(buildStyle('false', '0'));
    expect(r.isDefault).toBe(false);
  });

  it('parses w:customStyle="on" as true', async () => {
    const r = await loadStyle(buildStyle('0', 'on'));
    expect(r.customStyle).toBe(true);
  });

  it('parses w:customStyle="off" as falsy', async () => {
    const r = await loadStyle(buildStyle('0', 'off'));
    expect(!r.customStyle).toBe(true);
  });

  it('parses w:customStyle="false" as falsy', async () => {
    const r = await loadStyle(buildStyle('0', 'false'));
    expect(!r.customStyle).toBe(true);
  });
});

describe('Style pPr spacing autospacing — honour ST_OnOff per ECMA-376', () => {
  const buildStyle = (beforeVal: string, afterVal: string) =>
    `<w:style w:type="paragraph" w:styleId="Test"><w:name w:val="Test"/><w:pPr><w:spacing w:beforeAutospacing="${beforeVal}" w:afterAutospacing="${afterVal}"/></w:pPr></w:style>`;

  it('parses w:beforeAutospacing="on" as true', async () => {
    const r = await loadStyle(buildStyle('on', '0'));
    expect(r.beforeAutospacing).toBe(true);
  });

  it('parses w:beforeAutospacing="off" as false', async () => {
    const r = await loadStyle(buildStyle('off', '1'));
    expect(r.beforeAutospacing).toBe(false);
  });

  it('parses w:afterAutospacing="on" as true', async () => {
    const r = await loadStyle(buildStyle('0', 'on'));
    expect(r.afterAutospacing).toBe(true);
  });

  it('parses w:afterAutospacing="false" as false', async () => {
    const r = await loadStyle(buildStyle('1', 'false'));
    expect(r.afterAutospacing).toBe(false);
  });
});

async function makeDocxWithComment(commentXml: string): Promise<Buffer> {
  const zipHandler = new ZipHandler();

  zipHandler.addFile(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>`
  );

  zipHandler.addFile(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );

  zipHandler.addFile(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>`
  );

  zipHandler.addFile(
    'word/comments.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
${commentXml}
</w:comments>`
  );

  zipHandler.addFile(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>test</w:t></w:r></w:p></w:body>
</w:document>`
  );

  return await zipHandler.toBuffer();
}

describe('Comment w:done — honour ST_OnOff per ECMA-376', () => {
  it('parses w:done="on" as resolved', async () => {
    const buffer = await makeDocxWithComment(
      `<w:comment w:id="1" w:author="Tester" w:date="2024-01-01T00:00:00Z" w:done="on"><w:p><w:r><w:t>test</w:t></w:r></w:p></w:comment>`
    );
    const doc = await Document.loadFromBuffer(buffer);
    const comment = doc.getCommentManager().getAllComments()[0];
    expect(comment?.isResolved()).toBe(true);
    doc.dispose();
  });

  it('parses w:done="1" as resolved', async () => {
    const buffer = await makeDocxWithComment(
      `<w:comment w:id="1" w:author="Tester" w:date="2024-01-01T00:00:00Z" w:done="1"><w:p><w:r><w:t>test</w:t></w:r></w:p></w:comment>`
    );
    const doc = await Document.loadFromBuffer(buffer);
    const comment = doc.getCommentManager().getAllComments()[0];
    expect(comment?.isResolved()).toBe(true);
    doc.dispose();
  });

  it('parses w:done="true" as resolved', async () => {
    const buffer = await makeDocxWithComment(
      `<w:comment w:id="1" w:author="Tester" w:date="2024-01-01T00:00:00Z" w:done="true"><w:p><w:r><w:t>test</w:t></w:r></w:p></w:comment>`
    );
    const doc = await Document.loadFromBuffer(buffer);
    const comment = doc.getCommentManager().getAllComments()[0];
    expect(comment?.isResolved()).toBe(true);
    doc.dispose();
  });

  it('parses w:done="off" as unresolved', async () => {
    const buffer = await makeDocxWithComment(
      `<w:comment w:id="1" w:author="Tester" w:date="2024-01-01T00:00:00Z" w:done="off"><w:p><w:r><w:t>test</w:t></w:r></w:p></w:comment>`
    );
    const doc = await Document.loadFromBuffer(buffer);
    const comment = doc.getCommentManager().getAllComments()[0];
    expect(comment?.isResolved()).toBe(false);
    doc.dispose();
  });

  it('parses w:done="0" as unresolved', async () => {
    const buffer = await makeDocxWithComment(
      `<w:comment w:id="1" w:author="Tester" w:date="2024-01-01T00:00:00Z" w:done="0"><w:p><w:r><w:t>test</w:t></w:r></w:p></w:comment>`
    );
    const doc = await Document.loadFromBuffer(buffer);
    const comment = doc.getCommentManager().getAllComments()[0];
    expect(comment?.isResolved()).toBe(false);
    doc.dispose();
  });

  it('parses w:done="false" as unresolved', async () => {
    const buffer = await makeDocxWithComment(
      `<w:comment w:id="1" w:author="Tester" w:date="2024-01-01T00:00:00Z" w:done="false"><w:p><w:r><w:t>test</w:t></w:r></w:p></w:comment>`
    );
    const doc = await Document.loadFromBuffer(buffer);
    const comment = doc.getCommentManager().getAllComments()[0];
    expect(comment?.isResolved()).toBe(false);
    doc.dispose();
  });

  it('parses absent w:done as unresolved', async () => {
    const buffer = await makeDocxWithComment(
      `<w:comment w:id="1" w:author="Tester" w:date="2024-01-01T00:00:00Z"><w:p><w:r><w:t>test</w:t></w:r></w:p></w:comment>`
    );
    const doc = await Document.loadFromBuffer(buffer);
    const comment = doc.getCommentManager().getAllComments()[0];
    expect(comment?.isResolved()).toBe(false);
    doc.dispose();
  });
});
