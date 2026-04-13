/**
 * Tests for Run.splitAt()
 */

import { Run } from '../../src/elements/Run';

describe('Run.splitAt()', () => {
  describe('basic text splitting', () => {
    it('splits plain text at a character position', () => {
      const run = new Run('Hello World', { bold: true });
      const tail = run.splitAt(5);

      expect(run.getText()).toBe('Hello');
      expect(tail.getText()).toBe(' World');
    });

    it('preserves formatting on both halves', () => {
      const run = new Run('Hello World', {
        bold: true,
        italic: true,
        color: 'FF0000',
        font: 'Arial',
        size: 12,
      });
      const tail = run.splitAt(5);

      const origFmt = run.getFormatting();
      const tailFmt = tail.getFormatting();

      expect(origFmt.bold).toBe(true);
      expect(origFmt.italic).toBe(true);
      expect(origFmt.color).toBe('FF0000');
      expect(origFmt.font).toBe('Arial');

      expect(tailFmt.bold).toBe(true);
      expect(tailFmt.italic).toBe(true);
      expect(tailFmt.color).toBe('FF0000');
      expect(tailFmt.font).toBe('Arial');
    });

    it('deep-clones formatting (mutations are independent)', () => {
      const run = new Run('Hello World', { bold: true });
      const tail = run.splitAt(5);

      tail.setBold(false);
      expect(run.getFormatting().bold).toBe(true);
    });

    it('splits at first character', () => {
      const run = new Run('ABCDEF');
      const tail = run.splitAt(1);

      expect(run.getText()).toBe('A');
      expect(tail.getText()).toBe('BCDEF');
    });

    it('splits at last character', () => {
      const run = new Run('ABCDEF');
      const tail = run.splitAt(5);

      expect(run.getText()).toBe('ABCDE');
      expect(tail.getText()).toBe('F');
    });
  });

  describe('edge cases', () => {
    it('returns empty run when offset >= text length', () => {
      const run = new Run('Hello');
      const tail = run.splitAt(5);

      expect(run.getText()).toBe('Hello');
      expect(tail.getText()).toBe('');
    });

    it('returns empty run when offset > text length', () => {
      const run = new Run('Hi');
      const tail = run.splitAt(100);

      expect(run.getText()).toBe('Hi');
      expect(tail.getText()).toBe('');
    });

    it('moves all content when offset <= 0', () => {
      const run = new Run('Hello');
      const tail = run.splitAt(0);

      expect(run.getText()).toBe('');
      expect(tail.getText()).toBe('Hello');
    });

    it('moves all content when offset is negative', () => {
      const run = new Run('Hello');
      const tail = run.splitAt(-5);

      expect(run.getText()).toBe('');
      expect(tail.getText()).toBe('Hello');
    });

    it('handles empty run', () => {
      const run = new Run('');
      const tail = run.splitAt(0);

      expect(run.getText()).toBe('');
      expect(tail.getText()).toBe('');
    });

    it('handles single character run - split at 0', () => {
      const run = new Run('X');
      const tail = run.splitAt(0);

      expect(run.getText()).toBe('');
      expect(tail.getText()).toBe('X');
    });

    it('handles single character run - split at 1', () => {
      const run = new Run('X');
      const tail = run.splitAt(1);

      expect(run.getText()).toBe('X');
      expect(tail.getText()).toBe('');
    });
  });

  describe('special content elements', () => {
    it('splits around tab characters', () => {
      const run = new Run('');
      run.appendText('Name');
      run.addTab();
      run.appendText('Value');

      // "Name\tValue" → split after tab at offset 5
      const tail = run.splitAt(5);
      expect(run.getText()).toBe('Name\t');
      expect(tail.getText()).toBe('Value');
    });

    it('splits before a tab character', () => {
      const run = new Run('');
      run.appendText('Name');
      run.addTab();
      run.appendText('Value');

      // "Name\tValue" → split at offset 4 (before tab)
      const tail = run.splitAt(4);
      expect(run.getText()).toBe('Name');
      expect(tail.getText()).toBe('\tValue');
    });

    it('splits around break elements', () => {
      const run = new Run('');
      run.appendText('Line 1');
      run.addBreak();
      run.appendText('Line 2');

      // "Line 1\nLine 2" → split at offset 7 (after break)
      const tail = run.splitAt(7);
      expect(run.getText()).toBe('Line 1\n');
      expect(tail.getText()).toBe('Line 2');
    });

    it('splits before a break element', () => {
      const run = new Run('');
      run.appendText('Line 1');
      run.addBreak();
      run.appendText('Line 2');

      // Split at offset 6 (before break)
      const tail = run.splitAt(6);
      expect(run.getText()).toBe('Line 1');
      expect(tail.getText()).toBe('\nLine 2');
    });

    it('handles multiple special elements', () => {
      const run = new Run('');
      run.appendText('A');
      run.addTab();
      run.appendText('B');
      run.addTab();
      run.appendText('C');

      // "A\tB\tC" → split at offset 3 (after "A\tB")
      const tail = run.splitAt(3);
      expect(run.getText()).toBe('A\tB');
      expect(tail.getText()).toBe('\tC');
    });
  });

  describe('content independence', () => {
    it('modifications to original do not affect tail', () => {
      const run = new Run('Hello World');
      const tail = run.splitAt(5);

      run.setText('Changed');
      expect(tail.getText()).toBe(' World');
    });

    it('modifications to tail do not affect original', () => {
      const run = new Run('Hello World');
      const tail = run.splitAt(5);

      tail.setText('Changed');
      expect(run.getText()).toBe('Hello');
    });
  });

  describe('XML generation after split', () => {
    it('both halves generate valid XML', () => {
      const run = new Run('Hello World', { bold: true });
      const tail = run.splitAt(5);

      const origXml = run.toXML();
      const tailXml = tail.toXML();

      expect(origXml.name).toBe('w:r');
      expect(tailXml.name).toBe('w:r');

      // Both should have rPr with bold
      const origRPr = origXml.children?.find((c) => typeof c !== 'string' && c.name === 'w:rPr');
      const tailRPr = tailXml.children?.find((c) => typeof c !== 'string' && c.name === 'w:rPr');
      expect(origRPr).toBeDefined();
      expect(tailRPr).toBeDefined();
    });
  });

  describe('practical use cases', () => {
    it('can be used to apply formatting to a sub-range', () => {
      // "Hello World" → make "World" bold
      const run = new Run('Hello World');
      const tail = run.splitAt(6); // tail = "World"

      tail.setBold(true);

      expect(run.getText()).toBe('Hello ');
      expect(run.getFormatting().bold).toBeUndefined();
      expect(tail.getText()).toBe('World');
      expect(tail.getFormatting().bold).toBe(true);
    });

    it('can split into three parts for mid-text formatting', () => {
      // "Hello Beautiful World" → bold only "Beautiful"
      const run = new Run('Hello Beautiful World');

      const midAndTail = run.splitAt(6); // run="Hello ", midAndTail="Beautiful World"
      const tail = midAndTail.splitAt(9); // midAndTail="Beautiful", tail=" World"

      midAndTail.setBold(true);

      expect(run.getText()).toBe('Hello ');
      expect(midAndTail.getText()).toBe('Beautiful');
      expect(midAndTail.getFormatting().bold).toBe(true);
      expect(tail.getText()).toBe(' World');
      expect(tail.getFormatting().bold).toBeUndefined();
    });
  });
});
