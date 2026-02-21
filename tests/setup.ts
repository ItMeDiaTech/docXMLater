/**
 * Jest Test Setup File
 * Configures the test environment and global utilities
 */

import { createHash } from 'crypto';
import { setGlobalLogger, SilentLogger } from '../src/utils/logger';
import { Document } from '../src/core/Document';
import { validateOoxml } from './helpers/validateOoxml';

// Increase timeout for async operations
jest.setTimeout(30000);

// Suppress library warnings during tests (set DOCXMLATER_LOG_LEVEL=warn to enable)
if (!process.env.DOCXMLATER_LOG_LEVEL) {
  setGlobalLogger(new SilentLogger());
}

// Global test utilities
beforeEach(() => {
  // Reset any global state if needed
});

afterEach(() => {
  // Cleanup after each test if needed
});

// Clean up after all tests
afterAll(() => {
  // Ensure any async operations are cleaned up
  jest.clearAllTimers();
});

// OOXML Validation: monkey-patch Document.toBuffer/save to validate every output buffer
if (!process.env.SKIP_OOXML_VALIDATION) {
  const validatedHashes = new Set<string>();

  const originalToBuffer = Document.prototype.toBuffer;
  const originalSave = Document.prototype.save;

  Document.prototype.toBuffer = async function (): Promise<Buffer> {
    const buffer = await originalToBuffer.call(this);
    const hash = createHash('sha256').update(buffer).digest('hex');
    if (!validatedHashes.has(hash)) {
      await validateOoxml(buffer);
      validatedHashes.add(hash);
    }
    return buffer;
  };

  Document.prototype.save = async function (filePath: string): Promise<void> {
    await originalSave.call(this, filePath);
    const { promises: fs } = await import('fs');
    const buffer = await fs.readFile(filePath);
    const hash = createHash('sha256').update(buffer).digest('hex');
    if (!validatedHashes.has(hash)) {
      await validateOoxml(buffer);
      validatedHashes.add(hash);
    }
  };
}
