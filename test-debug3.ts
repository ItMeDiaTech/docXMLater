const map = new Map<number, string>();
map.set(0, "zero");

console.log("Has 0:", map.has(0));
console.log("Get 0:", map.get(0));

// Test with actual ID
const id: number = 0;
console.log("Has id:", map.has(id));
console.log("Get id:", map.get(id));
console.log("ID is:", id);
console.log("ID type:", typeof id);
