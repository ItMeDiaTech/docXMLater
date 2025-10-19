import { NumberingManager } from './src/formatting/NumberingManager';

const manager = new NumberingManager();
console.log("Before createBulletList:");
console.log("nextAbstractNumId:", (manager as any).nextAbstractNumId);
console.log("nextNumId:", (manager as any).nextNumId);

const numId = manager.createBulletList();
console.log("\nAfter createBulletList:");
console.log("Returned numId:", numId);
console.log("nextAbstractNumId:", (manager as any).nextAbstractNumId);
console.log("nextNumId:", (manager as any).nextNumId);

const instance = manager.getNumberingInstance(numId);
console.log("\nInstance retrieved:", instance !== undefined);
console.log("Instance abstractNumId:", instance?.getAbstractNumId());

if (instance) {
  const abstractId = instance.getAbstractNumId();
  const abstractNum = manager.getAbstractNumbering(abstractId);
  console.log("Abstract numbering found:", abstractNum !== undefined);
  console.log("Abstract numbering ID:", abstractNum?.getId());
}
