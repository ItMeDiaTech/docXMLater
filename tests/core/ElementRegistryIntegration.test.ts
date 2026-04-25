/**
 * Verifies ElementRegistry actually wires through DocumentParser →
 * RegisteredBodyElement. Without this hookup the registry is orphan API
 * surface.
 *
 * The OOXML validator (tests/setup.ts) rejects custom-namespace body
 * children so we exercise parse and serialize independently rather than
 * round-tripping through `doc.toBuffer()`.
 */
import { Document } from '../../src/core/Document';
import { ElementRegistry } from '../../src/core/ElementRegistry';
import { RegisteredBodyElement } from '../../src/elements/RegisteredBodyElement';

describe('ElementRegistry → DocumentParser/Generator integration', () => {
  beforeEach(() => {
    ElementRegistry.clear();
  });
  afterEach(() => {
    ElementRegistry.clear();
  });

  it('parses a registered tag at body level into a RegisteredBodyElement carrying the model', async () => {
    interface MyBlock {
      label: string;
    }
    let parseCount = 0;
    ElementRegistry.register<MyBlock>('myco:block', {
      parse: (xml) => {
        parseCount++;
        const m = /label="([^"]*)"/.exec(xml);
        return { label: m?.[1] ?? '' };
      },
      serialize: (el) => `<myco:block xmlns:myco="urn:myco" label="${el.label}"/>`,
    });

    // Build a valid base document, then inject the registered element directly
    // into document.xml at the ZIP layer.
    const seed = Document.create();
    seed.createParagraph('before');
    seed.createParagraph('after');
    const buf = await seed.toBuffer();
    seed.dispose();

    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(buf);
    const docXml = await zip.file('word/document.xml')!.async('string');
    const injected = docXml.replace(
      /<w:sectPr/,
      '<myco:block xmlns:myco="urn:myco" label="hello"/><w:sectPr'
    );
    zip.file('word/document.xml', injected);
    const injectedBuf = await zip.generateAsync({ type: 'nodebuffer' });

    const doc = await Document.loadFromBuffer(injectedBuf);
    expect(parseCount).toBe(1);

    const registered = doc
      .getBodyElements()
      .filter((el): el is RegisteredBodyElement<MyBlock> => el instanceof RegisteredBodyElement);
    expect(registered).toHaveLength(1);
    expect(registered[0]!.getTag()).toBe('myco:block');
    expect(registered[0]!.getModel().label).toBe('hello');

    doc.dispose();
  });

  it('RegisteredBodyElement.toXML() round-trips through the registered serializer', () => {
    interface MyBlock {
      n: number;
    }
    let serializeCount = 0;
    const handler = {
      parse: () => ({ n: 0 }) as MyBlock,
      serialize: (el: MyBlock) => {
        serializeCount++;
        return `<myco:block n="${el.n}"/>`;
      },
    };
    const el = new RegisteredBodyElement<MyBlock>(
      'myco:block',
      { n: 42 },
      handler,
      '<myco:block n="0"/>'
    );

    const xml = el.toXML() as { name: string; rawXml: string };
    expect(xml.name).toBe('__rawXml');
    expect(xml.rawXml).toBe('<myco:block n="42"/>');
    expect(serializeCount).toBe(1);
  });

  it('emits original XML when the serializer throws so a buggy custom serializer cannot corrupt the output', () => {
    interface MyBlock {
      n: number;
    }
    const handler = {
      parse: () => ({ n: 0 }) as MyBlock,
      serialize: (): string => {
        throw new Error('serialize failed');
      },
    };
    const el = new RegisteredBodyElement<MyBlock>(
      'myco:block',
      { n: 42 },
      handler,
      '<myco:block label="orig"/>'
    );

    const xml = el.toXML() as { name: string; rawXml: string };
    expect(xml.rawXml).toBe('<myco:block label="orig"/>');
  });

  it('falls back to PreservedElement when the custom parser throws during load', async () => {
    ElementRegistry.register('myco:bad', {
      parse: () => {
        throw new Error('boom');
      },
      serialize: () => '<myco:bad/>',
    });

    const seed = Document.create();
    seed.createParagraph('x');
    const buf = await seed.toBuffer();
    seed.dispose();

    const JSZip = require('jszip');
    const zip = await JSZip.loadAsync(buf);
    const docXml = await zip.file('word/document.xml')!.async('string');
    const injected = docXml.replace(
      /<w:sectPr/,
      '<myco:bad xmlns:myco="urn:myco" label="x"/><w:sectPr'
    );
    zip.file('word/document.xml', injected);
    const injectedBuf = await zip.generateAsync({ type: 'nodebuffer' });

    // Must not throw — load degrades gracefully.
    const doc = await Document.loadFromBuffer(injectedBuf);
    const registered = doc.getBodyElements().filter((el) => el instanceof RegisteredBodyElement);
    expect(registered).toHaveLength(0);
    doc.dispose();
  });
});
