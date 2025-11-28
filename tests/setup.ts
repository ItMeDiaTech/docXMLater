/**
 * Jest setup file
 * Runs before all tests
 */

// Increase timeout for long-running tests
jest.setTimeout(30000);

// Suppress console output during tests (optional - can be commented out for debugging)
// global.console = {
//   ...console,
//   log: jest.fn(),
//   debug: jest.fn(),
//   info: jest.fn(),
//   warn: jest.fn(),
// };

// Clean up after all tests
afterAll(() => {
  // Ensure any async operations are cleaned up
  jest.clearAllTimers();
});
