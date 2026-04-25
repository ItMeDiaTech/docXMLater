import {
  ValidationRuleRegistry,
  type CustomValidationRule,
} from '../../src/validation/ValidationRuleRegistry';
import { Document } from '../../src/core/Document';

describe('ValidationRuleRegistry', () => {
  beforeEach(() => {
    ValidationRuleRegistry.clear();
  });

  function makeRule(code: string, message: string): CustomValidationRule {
    return {
      code,
      description: 'test rule',
      severity: 'warning',
      validate: () => [{ ruleCode: code, severity: 'warning', message }],
    };
  }

  it('registers and retrieves rules', () => {
    const rule = makeRule('TEST-001', 'hello');
    ValidationRuleRegistry.register(rule);
    expect(ValidationRuleRegistry.has('TEST-001')).toBe(true);
    expect(ValidationRuleRegistry.get('TEST-001')).toBe(rule);
  });

  it('throws on duplicate code', () => {
    const rule = makeRule('DUP', 'x');
    ValidationRuleRegistry.register(rule);
    expect(() => ValidationRuleRegistry.register(rule)).toThrow(/already registered/);
  });

  it('runs all registered rules against a document', () => {
    ValidationRuleRegistry.register(makeRule('R1', 'r1 msg'));
    ValidationRuleRegistry.register(makeRule('R2', 'r2 msg'));
    const doc = Document.create();
    const issues = ValidationRuleRegistry.runAll(doc);
    expect(issues).toHaveLength(2);
    expect(issues.map((i) => i.ruleCode).sort()).toEqual(['R1', 'R2']);
    doc.dispose();
  });

  it('does not abort other rules when a rule throws', () => {
    ValidationRuleRegistry.register({
      code: 'BAD',
      description: 'thrower',
      severity: 'error',
      validate: () => {
        throw new Error('boom');
      },
    });
    ValidationRuleRegistry.register(makeRule('GOOD', 'works'));
    const doc = Document.create();
    const issues = ValidationRuleRegistry.runAll(doc);
    expect(issues).toHaveLength(2);
    const codes = issues.map((i) => i.ruleCode).sort();
    expect(codes).toEqual(['BAD', 'GOOD']);
    const badIssue = issues.find((i) => i.ruleCode === 'BAD')!;
    expect(badIssue.severity).toBe('error');
    expect(badIssue.message).toMatch(/threw/);
    doc.dispose();
  });

  it('unregister + clear work as expected', () => {
    ValidationRuleRegistry.register(makeRule('A', ''));
    ValidationRuleRegistry.register(makeRule('B', ''));
    expect(ValidationRuleRegistry.unregister('A')).toBe(true);
    expect(ValidationRuleRegistry.getAll()).toHaveLength(1);
    ValidationRuleRegistry.clear();
    expect(ValidationRuleRegistry.getAll()).toHaveLength(0);
  });
});
