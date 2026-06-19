/**
 * REV003/REV004 (orphaned moveFrom/moveTo) are produced by one pairing
 * pass, and REV101/REV102 (missing/invalid dates) by one date pass. Rule
 * gating must be per-code: skipping one rule of a pair must not silently
 * suppress the other (REV004 is error-severity and corruption-relevant),
 * and the fixer's onlyRules/skipRules must not let it mutate revisions
 * that belong to an excluded rule.
 */

import { Document } from '../../src/core/Document';
import { Paragraph } from '../../src/elements/Paragraph';
import { Run } from '../../src/elements/Run';
import { Revision } from '../../src/elements/Revision';
import { RevisionValidator } from '../../src/validation/RevisionValidator';
import { RevisionAutoFixer } from '../../src/validation/RevisionAutoFixer';

function makeOrphanedMoveFrom(id: number): Revision {
  return new Revision({
    id,
    type: 'moveFrom',
    author: 'Author',
    date: new Date('2025-01-15T10:00:00Z'),
    content: [new Run('moved away')],
    moveId: `mf-${id}`,
  });
}

function makeOrphanedMoveTo(id: number): Revision {
  return new Revision({
    id,
    type: 'moveTo',
    author: 'Author',
    date: new Date('2025-01-15T10:00:00Z'),
    content: [new Run('moved here')],
    moveId: `mt-${id}`,
  });
}

describe('RevisionValidator — per-rule gating for paired rules', () => {
  let doc: Document | undefined;

  afterEach(() => {
    doc?.dispose();
    doc = undefined;
  });

  function setupOrphanedMoves(): Document {
    doc = Document.create();
    const para = new Paragraph();
    const moveFrom = makeOrphanedMoveFrom(1);
    const moveTo = makeOrphanedMoveTo(2);
    para.addRevision(moveFrom);
    para.addRevision(moveTo);
    doc.addParagraph(para);
    doc.getRevisionManager().registerExisting(moveFrom);
    doc.getRevisionManager().registerExisting(moveTo);
    return doc;
  }

  it('skipRules:[REV003] still reports orphaned moveTo (REV004)', () => {
    const d = setupOrphanedMoves();
    const result = RevisionValidator.validate(d, { skipRules: ['REV003'] });

    expect(result.errors.some((i) => i.code === 'REV004')).toBe(true);
    expect(result.errors.some((i) => i.code === 'REV003')).toBe(false);
  });

  it('skipRules:[REV004] still reports orphaned moveFrom (REV003)', () => {
    const d = setupOrphanedMoves();
    const result = RevisionValidator.validate(d, { skipRules: ['REV004'] });

    expect(result.errors.some((i) => i.code === 'REV003')).toBe(true);
    expect(result.errors.some((i) => i.code === 'REV004')).toBe(false);
  });

  function setupDateIssues(): Document {
    doc = Document.create();
    const missingDate = new Revision({
      id: 1,
      type: 'insert',
      author: 'Author',
      content: [new Run('no date')],
    });
    missingDate.setDate(undefined as unknown as Date);
    const invalidDate = new Revision({
      id: 2,
      type: 'insert',
      author: 'Author',
      date: new Date('not-a-date'),
      content: [new Run('bad date')],
    });
    doc.getRevisionManager().registerExisting(missingDate);
    doc.getRevisionManager().registerExisting(invalidDate);
    return doc;
  }

  it('skipRules:[REV101] still reports invalid dates (REV102)', () => {
    const d = setupDateIssues();
    const result = RevisionValidator.validate(d, { skipRules: ['REV101'] });

    expect(result.warnings.some((i) => i.code === 'REV102')).toBe(true);
    expect(result.warnings.some((i) => i.code === 'REV101')).toBe(false);
  });

  it('skipRules:[REV102] still reports missing dates (REV101)', () => {
    const d = setupDateIssues();
    const result = RevisionValidator.validate(d, { skipRules: ['REV102'] });

    expect(result.warnings.some((i) => i.code === 'REV101')).toBe(true);
    expect(result.warnings.some((i) => i.code === 'REV102')).toBe(false);
  });
});

describe('RevisionAutoFixer — onlyRules/skipRules gate moveFrom and moveTo fixes independently', () => {
  let doc: Document | undefined;
  let para: Paragraph;
  let moveFrom: Revision;
  let moveTo: Revision;

  beforeEach(() => {
    doc = Document.create();
    para = new Paragraph();
    moveFrom = makeOrphanedMoveFrom(1);
    moveTo = makeOrphanedMoveTo(2);
    para.addRevision(moveFrom);
    para.addRevision(moveTo);
    doc.addParagraph(para);
    doc.getRevisionManager().registerExisting(moveFrom);
    doc.getRevisionManager().registerExisting(moveTo);
  });

  afterEach(() => {
    doc?.dispose();
    doc = undefined;
  });

  it('onlyRules:[REV003] removes orphaned moveFrom but leaves orphaned moveTo intact', () => {
    const result = RevisionAutoFixer.fix(doc!, { onlyRules: ['REV003'] });

    expect(result.actions.some((a) => a.issue.code === 'REV003')).toBe(true);
    expect(result.actions.some((a) => a.issue.code === 'REV004')).toBe(false);

    const manager = doc!.getRevisionManager();
    expect(manager.getAllRevisions()).not.toContain(moveFrom);
    expect(manager.getAllRevisions()).toContain(moveTo);
    expect(para.getRevisions()).not.toContain(moveFrom);
    expect(para.getRevisions()).toContain(moveTo);
  });

  it('onlyRules:[REV004] removes orphaned moveTo but leaves orphaned moveFrom intact', () => {
    const result = RevisionAutoFixer.fix(doc!, { onlyRules: ['REV004'] });

    expect(result.actions.some((a) => a.issue.code === 'REV004')).toBe(true);
    expect(result.actions.some((a) => a.issue.code === 'REV003')).toBe(false);

    const manager = doc!.getRevisionManager();
    expect(manager.getAllRevisions()).not.toContain(moveTo);
    expect(manager.getAllRevisions()).toContain(moveFrom);
    expect(para.getRevisions()).not.toContain(moveTo);
    expect(para.getRevisions()).toContain(moveFrom);
  });

  it('skipRules:[REV004] does not remove orphaned moveTo', () => {
    const result = RevisionAutoFixer.fix(doc!, { skipRules: ['REV004'] });

    expect(result.actions.some((a) => a.issue.code === 'REV004')).toBe(false);

    const manager = doc!.getRevisionManager();
    expect(manager.getAllRevisions()).toContain(moveTo);
    expect(para.getRevisions()).toContain(moveTo);
  });
});
