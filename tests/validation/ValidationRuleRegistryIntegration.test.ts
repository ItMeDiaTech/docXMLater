/**
 * Verifies ValidationRuleRegistry actually feeds into RevisionValidator —
 * without this hookup the registry was orphan API surface.
 */
import { Document } from '../../src/core/Document';
import { RevisionValidator } from '../../src/validation/RevisionValidator';
import {
  ValidationRuleRegistry,
  type CustomValidationRule,
} from '../../src/validation/ValidationRuleRegistry';

describe('ValidationRuleRegistry → RevisionValidator integration', () => {
  beforeEach(() => {
    ValidationRuleRegistry.clear();
  });
  afterEach(() => {
    ValidationRuleRegistry.clear();
  });

  it('runs registered rules and bucket issues by severity in the validation result', () => {
    const rule: CustomValidationRule = {
      code: 'CUSTOM-001',
      description: 'documents must have at least one paragraph',
      severity: 'error',
      validate: (doc) =>
        doc.getParagraphs().length === 0
          ? [
              {
                ruleCode: 'CUSTOM-001',
                severity: 'error',
                message: 'document is empty',
                location: 'body',
              },
            ]
          : [],
    };
    ValidationRuleRegistry.register(rule);

    const empty = Document.create();
    const emptyResult = RevisionValidator.validate(empty);
    expect(emptyResult.errors.some((e) => e.code === 'CUSTOM-001')).toBe(true);
    expect(emptyResult.valid).toBe(false);
    empty.dispose();

    const populated = Document.create();
    populated.createParagraph('hello');
    const populatedResult = RevisionValidator.validate(populated);
    expect(populatedResult.errors.some((e) => e.code === 'CUSTOM-001')).toBe(false);
    populated.dispose();
  });

  it('runs custom rules even when the document has no revisions', () => {
    let invoked = 0;
    ValidationRuleRegistry.register({
      code: 'WARN-001',
      description: 'always warns',
      severity: 'warning',
      validate: () => {
        invoked++;
        return [{ ruleCode: 'WARN-001', severity: 'warning', message: 'just a warning' }];
      },
    });

    const doc = Document.create();
    const result = RevisionValidator.validate(doc);
    expect(invoked).toBe(1);
    expect(result.warnings.some((w) => w.code === 'WARN-001')).toBe(true);
    doc.dispose();
  });

  it('continues running other rules when one rule throws', () => {
    ValidationRuleRegistry.register({
      code: 'BAD',
      description: 'thrower',
      severity: 'error',
      validate: () => {
        throw new Error('boom');
      },
    });
    ValidationRuleRegistry.register({
      code: 'GOOD',
      description: 'works',
      severity: 'info',
      validate: () => [{ ruleCode: 'GOOD', severity: 'info', message: 'all good' }],
    });

    const doc = Document.create();
    const result = RevisionValidator.validate(doc);
    const codes = [...result.errors, ...result.infos].map((i) => i.code).sort();
    expect(codes).toContain('BAD');
    expect(codes).toContain('GOOD');
    doc.dispose();
  });
});
