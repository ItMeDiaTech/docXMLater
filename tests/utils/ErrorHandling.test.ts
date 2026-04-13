import { isError, toError, wrapError, getErrorMessage } from '../../src/utils/errorHandling';

describe('isError', () => {
  it('should return true for Error instances', () => {
    expect(isError(new Error('test'))).toBe(true);
    expect(isError(new TypeError('type'))).toBe(true);
    expect(isError(new RangeError('range'))).toBe(true);
  });

  it('should return false for non-Error values', () => {
    expect(isError('string error')).toBe(false);
    expect(isError(42)).toBe(false);
    expect(isError(null)).toBe(false);
    expect(isError(undefined)).toBe(false);
    expect(isError({ message: 'fake' })).toBe(false);
  });
});

describe('toError', () => {
  it('should return Error instances unchanged', () => {
    const err = new Error('original');
    expect(toError(err)).toBe(err);
  });

  it('should convert strings to Error', () => {
    const result = toError('something failed');
    expect(result).toBeInstanceOf(Error);
    expect(result.message).toBe('something failed');
  });

  it('should convert objects with message property', () => {
    const result = toError({ message: 'obj error' });
    expect(result).toBeInstanceOf(Error);
    expect(result.message).toBe('obj error');
  });

  it('should handle other types', () => {
    expect(toError(42).message).toBe('42');
    expect(toError(null).message).toBe('null');
    expect(toError(undefined).message).toBe('undefined');
  });
});

describe('wrapError', () => {
  it('should add context to error message', () => {
    const original = new Error('file not found');
    const wrapped = wrapError(original, 'Failed to load document');

    expect(wrapped.message).toBe('Failed to load document: file not found');
  });

  it('should preserve original stack in caused-by chain', () => {
    const original = new Error('root cause');
    const wrapped = wrapError(original, 'Operation failed');

    expect(wrapped.stack).toContain('Caused by:');
  });

  it('should handle non-Error inputs', () => {
    const wrapped = wrapError('string error', 'Context');
    expect(wrapped.message).toBe('Context: string error');
  });
});

describe('getErrorMessage', () => {
  it('should extract message from Error', () => {
    expect(getErrorMessage(new Error('test msg'))).toBe('test msg');
  });

  it('should return strings directly', () => {
    expect(getErrorMessage('direct string')).toBe('direct string');
  });

  it('should extract from objects with message', () => {
    expect(getErrorMessage({ message: 'obj msg' })).toBe('obj msg');
  });

  it('should return fallback for unknown types', () => {
    expect(getErrorMessage(42)).toBe('Unknown error occurred');
    expect(getErrorMessage(null)).toBe('Unknown error occurred');
  });

  it('should use custom fallback', () => {
    expect(getErrorMessage(42, 'Custom fallback')).toBe('Custom fallback');
  });
});
