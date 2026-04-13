import {
  REVISION_RULES,
  createIssueFromRule,
  getRuleByCode,
  getRulesBySeverity,
  getAutoFixableRules,
} from '../../src/validation/ValidationRules';

describe('REVISION_RULES', () => {
  it('should define 10 rules', () => {
    expect(Object.keys(REVISION_RULES).length).toBe(10);
  });

  it('should have unique codes', () => {
    const codes = Object.values(REVISION_RULES).map((r) => r.code);
    expect(new Set(codes).size).toBe(codes.length);
  });

  it('should have 4 error-level rules', () => {
    const errors = Object.values(REVISION_RULES).filter((r) => r.severity === 'error');
    expect(errors.length).toBe(4);
  });

  it('should have 4 warning-level rules', () => {
    const warnings = Object.values(REVISION_RULES).filter((r) => r.severity === 'warning');
    expect(warnings.length).toBe(4);
  });

  it('should have 2 info-level rules', () => {
    const infos = Object.values(REVISION_RULES).filter((r) => r.severity === 'info');
    expect(infos.length).toBe(2);
  });

  it('should have error codes REV001-REV004', () => {
    expect(REVISION_RULES.DUPLICATE_ID.code).toBe('REV001');
    expect(REVISION_RULES.MISSING_AUTHOR.code).toBe('REV002');
    expect(REVISION_RULES.ORPHANED_MOVE_FROM.code).toBe('REV003');
    expect(REVISION_RULES.ORPHANED_MOVE_TO.code).toBe('REV004');
  });

  it('should have warning codes REV101-REV104', () => {
    expect(REVISION_RULES.MISSING_DATE.code).toBe('REV101');
    expect(REVISION_RULES.INVALID_DATE_FORMAT.code).toBe('REV102');
    expect(REVISION_RULES.EMPTY_REVISION.code).toBe('REV103');
    expect(REVISION_RULES.NON_SEQUENTIAL_IDS.code).toBe('REV104');
  });

  it('should have info codes REV201-REV202', () => {
    expect(REVISION_RULES.LARGE_REVISION_COUNT.code).toBe('REV201');
    expect(REVISION_RULES.OLD_REVISION_DATE.code).toBe('REV202');
  });

  it('should mark info rules as not auto-fixable', () => {
    expect(REVISION_RULES.LARGE_REVISION_COUNT.autoFixable).toBe(false);
    expect(REVISION_RULES.OLD_REVISION_DATE.autoFixable).toBe(false);
  });

  it('should mark error and warning rules as auto-fixable', () => {
    const fixable = Object.values(REVISION_RULES).filter(
      (r) => r.severity === 'error' || r.severity === 'warning'
    );
    expect(fixable.every((r) => r.autoFixable)).toBe(true);
  });

  it('should have suggested fixes for all auto-fixable rules', () => {
    const fixable = Object.values(REVISION_RULES).filter((r) => r.autoFixable);
    expect(fixable.every((r) => r.suggestedFix)).toBe(true);
  });
});

describe('getRuleByCode', () => {
  it('should find rules by code', () => {
    const rule = getRuleByCode('REV001');
    expect(rule).toBeDefined();
    expect(rule!.code).toBe('REV001');
    expect(rule!.severity).toBe('error');
  });

  it('should return undefined for unknown codes', () => {
    expect(getRuleByCode('REV999')).toBeUndefined();
    expect(getRuleByCode('')).toBeUndefined();
  });
});

describe('getRulesBySeverity', () => {
  it('should return error rules', () => {
    const errors = getRulesBySeverity('error');
    expect(errors.length).toBe(4);
    expect(errors.every((r) => r.severity === 'error')).toBe(true);
  });

  it('should return warning rules', () => {
    const warnings = getRulesBySeverity('warning');
    expect(warnings.length).toBe(4);
  });

  it('should return info rules', () => {
    const infos = getRulesBySeverity('info');
    expect(infos.length).toBe(2);
  });
});

describe('getAutoFixableRules', () => {
  it('should return 8 auto-fixable rules', () => {
    const fixable = getAutoFixableRules();
    expect(fixable.length).toBe(8);
    expect(fixable.every((r) => r.autoFixable)).toBe(true);
  });
});

describe('createIssueFromRule', () => {
  it('should create issue with default message', () => {
    const rule = REVISION_RULES.DUPLICATE_ID;
    const issue = createIssueFromRule(rule);

    expect(issue.code).toBe('REV001');
    expect(issue.severity).toBe('error');
    expect(issue.message).toBe(rule.message);
    expect(issue.autoFixable).toBe(true);
    expect(issue.suggestedFix).toBe(rule.suggestedFix);
  });

  it('should create issue with custom message', () => {
    const rule = REVISION_RULES.DUPLICATE_ID;
    const issue = createIssueFromRule(rule, undefined, 'Custom: ID 42 is duplicated');
    expect(issue.message).toBe('Custom: ID 42 is duplicated');
  });

  it('should include location when provided', () => {
    const rule = REVISION_RULES.MISSING_AUTHOR;
    const issue = createIssueFromRule(rule, { paragraphIndex: 5, revisionId: 10 });
    expect(issue.location).toEqual({ paragraphIndex: 5, revisionId: 10 });
  });

  it('should omit location when not provided', () => {
    const issue = createIssueFromRule(REVISION_RULES.EMPTY_REVISION);
    expect(issue.location).toBeUndefined();
  });
});
