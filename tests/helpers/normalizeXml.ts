/**
 * Normalize XML string for comparison by stripping non-deterministic content.
 */
export function normalizeXml(xml: string): string {
  let normalized = xml;

  // Strip rsid attributes (Word regenerates these)
  normalized = normalized.replace(/\s+w:rsid\w*="[^"]*"/g, '');
  normalized = normalized.replace(/\s+rsid\w*="[^"]*"/g, '');

  // Strip date/time values in dc:created and dc:modified (dcterms)
  normalized = normalized.replace(
    /(<dcterms:(created|modified)[^>]*>)[^<]*(<\/dcterms:(created|modified)>)/g,
    '$1NORMALIZED_DATE$3'
  );

  // Strip revision numbers (cp:revision)
  normalized = normalized.replace(
    /(<cp:revision>)[^<]*(<\/cp:revision>)/g,
    '$1NORMALIZED_REVISION$2'
  );

  // Normalize whitespace between XML tags (collapse multiple spaces/newlines)
  normalized = normalized.replace(/>\s+</g, '>\n<');

  // Trim lines
  normalized = normalized
    .split('\n')
    .map(line => line.trim())
    .filter(line => line.length > 0)
    .join('\n');

  return normalized;
}
