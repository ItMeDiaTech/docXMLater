/**
 * Jest Test Setup File
 * Configures the test environment and global utilities
 */

// Increase timeout for async operations
jest.setTimeout(30000);

// Suppress console output during tests (optional - can be commented out for debugging)
// global.console = {
//   ...console,
//   log: jest.fn(),
//   debug: jest.fn(),
//   info: jest.fn(),
//   warn: jest.fn(),
// };

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
