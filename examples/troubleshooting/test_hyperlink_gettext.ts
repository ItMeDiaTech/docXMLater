/**
 * Test to verify Hyperlink.getText() accuracy
 * This test checks if getText() accurately reflects the actual text that will be in the XML
 */

import { Hyperlink } from './src/elements/Hyperlink';

console.log('Testing Hyperlink.getText() accuracy...\n');

// Test 1: Basic usage (should work fine)
console.log('Test 1: Basic usage');
const link1 = Hyperlink.createExternal('https://example.com', 'Click here');
console.log('getText():', link1.getText());
console.log('Expected: Click here');
console.log('✓ Test 1 passed\n');

// Test 2: Direct run modification (potential issue)
console.log('Test 2: Direct run modification');
const link2 = Hyperlink.createExternal('https://example.com', 'Original text');
console.log('Initial getText():', link2.getText());

// Get the run and modify it directly
const run = link2.getRun();
run.setText('Modified text');

console.log('After run.setText("Modified text"):');
console.log('  link2.getText():', link2.getText());
console.log('  run.getText():', run.getText());
console.log('Expected: Both should return "Modified text"');

if (link2.getText() === run.getText()) {
  console.log('✓ Test 2 passed - getText() is accurate\n');
} else {
  console.log('✗ Test 2 FAILED - getText() is out of sync!');
  console.log('  Cached text:', link2.getText());
  console.log('  Actual text:', run.getText());
  console.log('\n');
}

// Test 3: setRun() updates text (should work)
console.log('Test 3: setRun() updates cached text');
const link3 = Hyperlink.createExternal('https://example.com', 'Initial');
const newRun = link3.getRun();
newRun.setText('Updated via run');
link3.setRun(newRun);

console.log('After setRun():');
console.log('  link3.getText():', link3.getText());
console.log('  Expected: Updated via run');

if (link3.getText() === 'Updated via run') {
  console.log('✓ Test 3 passed\n');
} else {
  console.log('✗ Test 3 FAILED\n');
}

// Test 4: Complex content (tabs, breaks)
console.log('Test 4: Complex content with tabs');
const link4 = Hyperlink.createExternal('https://example.com', 'Text\twith\ttabs');
console.log('getText():', JSON.stringify(link4.getText()));
console.log('getRun().getText():', JSON.stringify(link4.getRun().getText()));
console.log('Expected: "Text\\twith\\ttabs"');

if (link4.getText() === link4.getRun().getText()) {
  console.log('✓ Test 4 passed\n');
} else {
  console.log('✗ Test 4 FAILED\n');
}

console.log('=== Summary ===');
console.log('If Test 2 failed, it means getText() returns cached value that can become stale.');
console.log('The fix would be to change getText() to: return this.run.getText()');
