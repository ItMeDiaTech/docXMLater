/**
 * Comment instance mutations on loaded documents must persist to the
 * comments parts.
 *
 * Loaded documents keep the original comments.xml for passthrough and only
 * regenerate it when the comments dirty flag is set. The public Comment
 * mutators (resolve, unresolve, setAuthor, setInitials, setDate, addRun)
 * must flip that flag — otherwise the mutation silently vanishes on save.
 * Resolved state is serialized per the Office 2012 extension: CT_Comment
 * declares no done attribute, so w15:commentEx w15:done lives in
 * word/commentsExtended.xml keyed by the last paragraph's w14:paraId.
 */
import { Document } from '../../src/core/Document';
import { Run } from '../../src/elements/Run';
import { ZipHandler } from '../../src/zip/ZipHandler';

/** Builds a DOCX containing one unresolved comment by 'Author' */
async function buildDocxWithComment(): Promise<Buffer> {
  const doc = Document.create();
  doc.createParagraph('Body text');
  doc.createComment('Author', 'Needs review', 'AU');
  const buffer = await doc.toBuffer();
  doc.dispose();
  return buffer;
}

async function readPart(buffer: Buffer, path: string): Promise<string | undefined> {
  const zip = new ZipHandler();
  await zip.loadFromBuffer(buffer);
  return zip.getFileAsString(path);
}

describe('X42: comment mutations on loaded documents persist on save', () => {
  it('persists resolve() via commentsExtended.xml (w15:done="1")', async () => {
    const base = await buildDocxWithComment();
    const doc = await Document.loadFromBuffer(base);
    let out: Buffer;
    try {
      const comment = doc.getAllComments()[0]!;
      expect(comment.isResolved()).toBe(false);
      comment.resolve();
      out = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const extendedXml = await readPart(out, 'word/commentsExtended.xml');
    expect(extendedXml).toBeDefined();
    expect(extendedXml).toContain('w15:done="1"');

    // comments.xml anchors the commentEx entry via the last paragraph paraId
    const commentsXml = (await readPart(out, 'word/comments.xml'))!;
    const paraIdMatch = commentsXml.match(/<w:p w14:paraId="([0-9A-F]{8})">/);
    expect(paraIdMatch).toBeTruthy();
    expect(extendedXml).toContain(`w15:paraId="${paraIdMatch![1]}"`);

    // And the mutation survives a reload
    const reloaded = await Document.loadFromBuffer(out);
    try {
      expect(reloaded.getAllComments()[0]!.isResolved()).toBe(true);
    } finally {
      reloaded.dispose();
    }
  });

  it('persists unresolve() by dropping the w15:done flag', async () => {
    // Start from a resolved comment
    const seed = Document.create();
    seed.createParagraph('Body');
    seed.createComment('Author', 'Done already').resolve();
    const base = await seed.toBuffer();
    seed.dispose();
    expect(await readPart(base, 'word/commentsExtended.xml')).toContain('w15:done="1"');

    const doc = await Document.loadFromBuffer(base);
    let out: Buffer;
    try {
      const comment = doc.getAllComments()[0]!;
      expect(comment.isResolved()).toBe(true);
      comment.unresolve();
      out = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const extendedXml = await readPart(out, 'word/commentsExtended.xml');
    if (extendedXml !== undefined) {
      expect(extendedXml).not.toContain('w15:done="1"');
    }
    const reloaded = await Document.loadFromBuffer(out);
    try {
      expect(reloaded.getAllComments()[0]!.isResolved()).toBe(false);
    } finally {
      reloaded.dispose();
    }
  });

  it('persists setAuthor()/setInitials() changes', async () => {
    const base = await buildDocxWithComment();
    const doc = await Document.loadFromBuffer(base);
    let out: Buffer;
    try {
      doc.getAllComments()[0]!.setAuthor('New Author').setInitials('NA');
      out = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    const commentsXml = (await readPart(out, 'word/comments.xml'))!;
    expect(commentsXml).toContain('w:author="New Author"');
    expect(commentsXml).toContain('w:initials="NA"');
  });

  it('persists addRun() content', async () => {
    const base = await buildDocxWithComment();
    const doc = await Document.loadFromBuffer(base);
    let out: Buffer;
    try {
      doc.getAllComments()[0]!.addRun(new Run(' and more'));
      out = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    expect(await readPart(out, 'word/comments.xml')).toContain(' and more');
  });

  it('keeps the byte-exact passthrough when no comment is mutated', async () => {
    const base = await buildDocxWithComment();
    const original = await readPart(base, 'word/comments.xml');

    const doc = await Document.loadFromBuffer(base);
    let out: Buffer;
    try {
      // Read-only access must not flip the dirty flag
      doc.getAllComments()[0]!.isResolved();
      out = await doc.toBuffer();
    } finally {
      doc.dispose();
    }

    expect(await readPart(out, 'word/comments.xml')).toBe(original);
  });
});
