import { formatDateForXml } from '../../src/utils/dateFormatting';

describe('formatDateForXml', () => {
  it('should format a date without milliseconds', () => {
    const date = new Date('2024-01-15T10:30:00.000Z');
    expect(formatDateForXml(date)).toBe('2024-01-15T10:30:00Z');
  });

  it('should strip milliseconds from ISO string', () => {
    const date = new Date('2026-03-20T14:45:30.123Z');
    const result = formatDateForXml(date);
    expect(result).toBe('2026-03-20T14:45:30Z');
    expect(result).not.toContain('.123');
  });

  it('should handle midnight dates', () => {
    const date = new Date('2024-01-01T00:00:00.000Z');
    expect(formatDateForXml(date)).toBe('2024-01-01T00:00:00Z');
  });

  it('should handle end-of-day dates', () => {
    const date = new Date('2024-12-31T23:59:59.999Z');
    expect(formatDateForXml(date)).toBe('2024-12-31T23:59:59Z');
  });

  it('should produce valid OOXML w:date format', () => {
    const date = new Date();
    const result = formatDateForXml(date);
    // Must match ISO 8601 without milliseconds
    expect(result).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z$/);
  });

  it('should handle dates created from timestamps', () => {
    const date = new Date(1704067200000); // 2024-01-01T00:00:00.000Z
    expect(formatDateForXml(date)).toBe('2024-01-01T00:00:00Z');
  });
});
