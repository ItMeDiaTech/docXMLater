import { Document } from "../src/core/Document";

async function cleanDocument() {
  const doc = await Document.load('../pre-processed.docx');
  doc.clearCustom();
  if (await doc.partExists('customXml/item1.xml')) {
    await doc.removePart('customXml/item1.xml');
  }
  await doc.save('../Testing.docx');
  console.log('Cleaned document saved as Testing.docx');
}

cleanDocument().catch(console.error);
