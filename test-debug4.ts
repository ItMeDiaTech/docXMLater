import { NumberingManager } from './src/formatting/NumberingManager';

const manager = new NumberingManager();
const numId = manager.createBulletList();
const instance = manager.getNumberingInstance(numId);
const abstractId = instance?.getAbstractNumId();

console.log("abstractId value:", abstractId);
console.log("abstractId type:", typeof abstractId);
console.log("abstractId === 0:", abstractId === 0);
console.log("abstractId == 0:", abstractId == 0);

// Try direct access
console.log("\nDirect access test:");
const map = (manager as any).abstractNumberings;
console.log("Map size:", map.size);
console.log("Map keys:", Array.from(map.keys()));
console.log("Get 0:", map.get(0));
if (abstractId !== undefined) {
  console.log("Get abstractId:", map.get(abstractId));
}
