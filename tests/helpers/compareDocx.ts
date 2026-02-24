import JSZip from 'jszip';
import { normalizeXml } from './normalizeXml';

export interface DocxDiff {
  /** Files only in the first archive */
  onlyInFirst: string[];
  /** Files only in the second archive */
  onlyInSecond: string[];
  /** Files that differ between archives */
  differs: Array<{ path: string; expected: string; actual: string }>;
}

/**
 * Compare two DOCX buffers by unzipping and comparing XML parts.
 * Binary files (images, etc.) are compared by size only.
 */
export async function compareDocx(expected: Buffer, actual: Buffer): Promise<DocxDiff> {
  const zipA = await JSZip.loadAsync(expected);
  const zipB = await JSZip.loadAsync(actual);

  const filesA = Object.keys(zipA.files)
    .filter((f) => !zipA.files[f]!.dir)
    .sort();
  const filesB = Object.keys(zipB.files)
    .filter((f) => !zipB.files[f]!.dir)
    .sort();

  const setA = new Set(filesA);
  const setB = new Set(filesB);

  const onlyInFirst = filesA.filter((f) => !setB.has(f));
  const onlyInSecond = filesB.filter((f) => !setA.has(f));
  const common = filesA.filter((f) => setB.has(f));

  const differs: DocxDiff['differs'] = [];

  for (const filePath of common) {
    if (filePath.match(/\.(png|jpg|jpeg|gif|emf|wmf|svg|tiff)$/i)) {
      // Binary comparison by size
      const sizeA = (await zipA.files[filePath]!.async('uint8array')).length;
      const sizeB = (await zipB.files[filePath]!.async('uint8array')).length;
      if (sizeA !== sizeB) {
        differs.push({
          path: filePath,
          expected: `[binary ${sizeA} bytes]`,
          actual: `[binary ${sizeB} bytes]`,
        });
      }
      continue;
    }

    const contentA = await zipA.files[filePath]!.async('string');
    const contentB = await zipB.files[filePath]!.async('string');

    const normA =
      filePath.endsWith('.xml') || filePath.endsWith('.rels') ? normalizeXml(contentA) : contentA;
    const normB =
      filePath.endsWith('.xml') || filePath.endsWith('.rels') ? normalizeXml(contentB) : contentB;

    if (normA !== normB) {
      differs.push({ path: filePath, expected: normA, actual: normB });
    }
  }

  return { onlyInFirst, onlyInSecond, differs };
}

/**
 * Assert that two DOCX buffers are structurally equivalent.
 */
export async function expectDocxEqual(expected: Buffer, actual: Buffer): Promise<void> {
  const diff = await compareDocx(expected, actual);

  const messages: string[] = [];

  if (diff.onlyInFirst.length > 0) {
    messages.push(`Files only in expected: ${diff.onlyInFirst.join(', ')}`);
  }
  if (diff.onlyInSecond.length > 0) {
    messages.push(`Files only in actual: ${diff.onlyInSecond.join(', ')}`);
  }
  for (const d of diff.differs) {
    messages.push(`File differs: ${d.path}`);
  }

  if (messages.length > 0) {
    throw new Error(`DOCX files differ:\n${messages.join('\n')}`);
  }
}
