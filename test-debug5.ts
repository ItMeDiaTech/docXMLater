import { NumberingManager } from './src/formatting/NumberingManager';

const manager = new NumberingManager();
const numId = manager.createBulletList();
const instance = manager.getNumberingInstance(numId);
const abstractId = instance?.getAbstractNumId();

console.log("abstractId:", abstractId);

// Test the method directly
const result = manager.getAbstractNumbering(abstractId!);
console.log("getAbstractNumbering result:", result);
console.log("result !== undefined:", result !== undefined);

// Test with explicit 0
const result2 = manager.getAbstractNumbering(0);
console.log("getAbstractNumbering(0) result:", result2);
