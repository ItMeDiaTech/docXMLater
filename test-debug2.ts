// Simulate the exact test
import { NumberingManager } from './src/formatting/NumberingManager';

function test() {
  const manager = new NumberingManager();
  const numId = manager.createBulletList();

  console.log("numId defined:", numId !== undefined);
  const instance = manager.getNumberingInstance(numId);
  console.log("instance defined:", instance !== undefined);

  const abstractId = instance?.getAbstractNumId();
  console.log("abstractId:", abstractId);
  console.log("abstractId defined:", abstractId !== undefined);
  
  const abstractNum = abstractId ? manager.getAbstractNumbering(abstractId) : undefined;
  console.log("abstractNum defined:", abstractNum !== undefined);
  console.log("abstractNum:", abstractNum);
  
  // Try to debug the map
  const allAbstracts = manager.getAllAbstractNumberings();
  console.log("All abstracts count:", allAbstracts.length);
  console.log("All abstracts:", allAbstracts.map(a => a.getId()));
}

test();
