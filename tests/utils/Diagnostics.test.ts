import {
  enableDiagnostics,
  disableDiagnostics,
  getDiagnosticConfig,
  logParsing,
  logSerialization,
  logTextDirection,
  logVerbose,
  logParagraphContent,
  logTextComparison,
} from '../../src/utils/diagnostics';

describe('diagnostics', () => {
  afterEach(() => {
    disableDiagnostics();
  });

  describe('enableDiagnostics / disableDiagnostics', () => {
    it('should start disabled by default', () => {
      const config = getDiagnosticConfig();
      expect(config.enabled).toBe(false);
      expect(config.logParsing).toBe(false);
      expect(config.logSerialization).toBe(false);
    });

    it('should enable with defaults', () => {
      enableDiagnostics();
      const config = getDiagnosticConfig();
      expect(config.enabled).toBe(true);
    });

    it('should enable with specific options', () => {
      enableDiagnostics({ logParsing: true, verbose: true });
      const config = getDiagnosticConfig();
      expect(config.enabled).toBe(true);
      expect(config.logParsing).toBe(true);
      expect(config.verbose).toBe(true);
      expect(config.logSerialization).toBe(false);
    });

    it('should disable and reset all options', () => {
      enableDiagnostics({ logParsing: true, verbose: true });
      disableDiagnostics();
      const config = getDiagnosticConfig();
      expect(config.enabled).toBe(false);
      expect(config.logParsing).toBe(false);
      expect(config.verbose).toBe(false);
    });
  });

  describe('getDiagnosticConfig', () => {
    it('should return a copy (not the original)', () => {
      const config1 = getDiagnosticConfig();
      const config2 = getDiagnosticConfig();
      expect(config1).toEqual(config2);
      expect(config1).not.toBe(config2);
    });
  });

  describe('logging functions', () => {
    let consoleSpy: jest.SpyInstance;

    beforeEach(() => {
      consoleSpy = jest.spyOn(console, 'log').mockImplementation();
    });

    afterEach(() => {
      consoleSpy.mockRestore();
    });

    it('should not log when disabled', () => {
      logParsing('test message');
      logSerialization('test message');
      logTextDirection('test message');
      logVerbose('test message');
      expect(consoleSpy).not.toHaveBeenCalled();
    });

    it('should log parsing when enabled', () => {
      enableDiagnostics({ logParsing: true });
      logParsing('parsing element');
      expect(consoleSpy).toHaveBeenCalledWith('[PARSE] parsing element', '');
    });

    it('should log serialization when enabled', () => {
      enableDiagnostics({ logSerialization: true });
      logSerialization('serializing', { count: 5 });
      expect(consoleSpy).toHaveBeenCalledWith('[SERIALIZE] serializing', { count: 5 });
    });

    it('should log text direction when enabled', () => {
      enableDiagnostics({ logTextDirection: true });
      logTextDirection('RTL detected');
      expect(consoleSpy).toHaveBeenCalledWith('[TEXT-DIR] RTL detected', '');
    });

    it('should log verbose when enabled', () => {
      enableDiagnostics({ verbose: true });
      logVerbose('detail info');
      expect(consoleSpy).toHaveBeenCalledWith('[VERBOSE] detail info', '');
    });

    it('should not log parsing when only serialization is enabled', () => {
      enableDiagnostics({ logSerialization: true });
      logParsing('should not appear');
      expect(consoleSpy).not.toHaveBeenCalled();
    });
  });

  describe('logTextComparison', () => {
    let consoleSpy: jest.SpyInstance;

    beforeEach(() => {
      consoleSpy = jest.spyOn(console, 'log').mockImplementation();
    });

    afterEach(() => {
      consoleSpy.mockRestore();
    });

    it('should log mismatch when texts differ', () => {
      enableDiagnostics();
      logTextComparison('test', 'before', 'after');
      expect(consoleSpy).toHaveBeenCalledWith('[TEXT-CHANGE] test:');
    });

    it('should not log when disabled', () => {
      logTextComparison('test', 'before', 'after');
      expect(consoleSpy).not.toHaveBeenCalled();
    });

    it('should not log when texts match and not verbose', () => {
      enableDiagnostics();
      logTextComparison('test', 'same', 'same');
      expect(consoleSpy).not.toHaveBeenCalled();
    });
  });

  describe('logParagraphContent', () => {
    let consoleSpy: jest.SpyInstance;

    beforeEach(() => {
      consoleSpy = jest.spyOn(console, 'log').mockImplementation();
    });

    afterEach(() => {
      consoleSpy.mockRestore();
    });

    it('should not log when disabled', () => {
      logParagraphContent('parsing', 0, [{ text: 'Hello' }]);
      expect(consoleSpy).not.toHaveBeenCalled();
    });

    it('should log paragraph details when parsing enabled', () => {
      enableDiagnostics({ logParsing: true });
      logParagraphContent(
        'parsing',
        0,
        [
          { text: 'Hello', rtl: false },
          { text: 'World', rtl: true },
        ],
        true
      );
      expect(consoleSpy).toHaveBeenCalled();
    });
  });
});
