/**
 * Logging Examples for docXMLater
 * Demonstrates how to control framework logging behavior
 */

import {
  Document,
  ILogger,
  ConsoleLogger,
  SilentLogger,
  CollectingLogger,
  LogLevel,
  createScopedLogger,
} from '../../src/index';
import * as path from 'path';

const OUTPUT_DIR = path.join(__dirname, 'output');

/**
 * Example 1: Default Logging Behavior
 * By default, docXMLater uses ConsoleLogger with WARN minimum level
 */
async function example1_DefaultLogging() {
  console.log('\n=== Example 1: Default Logging ===');

  // No logger specified - uses default ConsoleLogger(WARN)
  const doc = Document.create();

  // Add lots of content to trigger size warning
  for (let i = 0; i < 100; i++) {
    doc.createParagraph(`This is paragraph ${i} with some content`);
  }

  // This will show warnings in console (if document is large)
  await doc.save(path.join(OUTPUT_DIR, 'default-logging.docx'));
  console.log('✓ Saved with default logging (console warnings shown)');
}

/**
 * Example 2: Silent Logging
 * Suppress all framework output using SilentLogger
 */
async function example2_SilentLogging() {
  console.log('\n=== Example 2: Silent Logging ===');

  // Use SilentLogger to suppress all framework output
  const doc = Document.create({
    logger: new SilentLogger(),
  });

  for (let i = 0; i < 100; i++) {
    doc.createParagraph(`Silent paragraph ${i}`);
  }

  // No warnings will be shown, even for large documents
  await doc.save(path.join(OUTPUT_DIR, 'silent-logging.docx'));
  console.log('✓ Saved with silent logging (no framework warnings)');
}

/**
 * Example 3: Verbose Logging
 * Use ConsoleLogger with DEBUG level for detailed output
 */
async function example3_VerboseLogging() {
  console.log('\n=== Example 3: Verbose Logging ===');

  // Enable DEBUG level for maximum verbosity
  const doc = Document.create({
    logger: new ConsoleLogger(LogLevel.DEBUG),
  });

  doc.createParagraph('Verbose logging example');

  await doc.save(path.join(OUTPUT_DIR, 'verbose-logging.docx'));
  console.log('✓ Saved with verbose logging');
}

/**
 * Example 4: Collecting Logs for Analysis
 * Use CollectingLogger to capture logs in memory
 */
async function example4_CollectingLogs() {
  console.log('\n=== Example 4: Collecting Logs ===');

  const logger = new CollectingLogger();
  const doc = Document.create({ logger });

  for (let i = 0; i < 150; i++) {
    doc.createParagraph(`Paragraph ${i}`);
  }

  await doc.save(path.join(OUTPUT_DIR, 'collected-logging.docx'));

  // Analyze collected logs
  const logs = logger.getLogs();
  console.log(`Total logs collected: ${logs.length}`);

  const warnings = logger.getLogsByLevel(LogLevel.WARN);
  console.log(`Warnings: ${warnings.length}`);

  warnings.forEach(log => {
    console.log(`  [${log.timestamp.toISOString()}] ${log.message}`);
    if (log.context) {
      console.log(`  Context:`, log.context);
    }
  });

  console.log('✓ Logs collected and analyzed');
}

/**
 * Example 5: Custom Logger Implementation
 * Implement ILogger for custom logging behavior
 */
class FileLogger implements ILogger {
  private logs: string[] = [];

  debug(message: string, context?: Record<string, any>): void {
    this.addLog('DEBUG', message, context);
  }

  info(message: string, context?: Record<string, any>): void {
    this.addLog('INFO', message, context);
  }

  warn(message: string, context?: Record<string, any>): void {
    this.addLog('WARN', message, context);
  }

  error(message: string, context?: Record<string, any>): void {
    this.addLog('ERROR', message, context);
  }

  private addLog(level: string, message: string, context?: Record<string, any>): void {
    const timestamp = new Date().toISOString();
    let logLine = `[${timestamp}] [${level}] ${message}`;

    if (context && Object.keys(context).length > 0) {
      logLine += ` | Context: ${JSON.stringify(context)}`;
    }

    this.logs.push(logLine);
  }

  saveToFile(filePath: string): void {
    const fs = require('fs');
    fs.writeFileSync(filePath, this.logs.join('\n'), 'utf-8');
  }

  getLogs(): string[] {
    return [...this.logs];
  }
}

async function example5_CustomLogger() {
  console.log('\n=== Example 5: Custom Logger ===');

  const fileLogger = new FileLogger();
  const doc = Document.create({ logger: fileLogger });

  for (let i = 0; i < 100; i++) {
    doc.createParagraph(`Custom logging paragraph ${i}`);
  }

  await doc.save(path.join(OUTPUT_DIR, 'custom-logging.docx'));

  // Save logs to file
  fileLogger.saveToFile(path.join(OUTPUT_DIR, 'document-generation.log'));
  console.log('✓ Saved with custom logger - logs written to file');
  console.log(`  Logs captured: ${fileLogger.getLogs().length}`);
}

/**
 * Example 6: Scoped Logger
 * Add source context to all log messages
 */
async function example6_ScopedLogger() {
  console.log('\n=== Example 6: Scoped Logger ===');

  const baseLogger = new CollectingLogger();
  const scopedLogger = createScopedLogger(baseLogger, 'DocumentBuilder');

  const doc = Document.create({ logger: scopedLogger });

  for (let i = 0; i < 120; i++) {
    doc.createParagraph(`Scoped logging ${i}`);
  }

  await doc.save(path.join(OUTPUT_DIR, 'scoped-logging.docx'));

  // Check logs have source context
  const logs = baseLogger.getLogs();
  console.log('Sample log entries with source:');
  logs.slice(0, 3).forEach(log => {
    console.log(`  ${log.message}`);
    console.log(`  Source: ${log.context?.source}`);
  });

  console.log('✓ Scoped logger adds source context to all messages');
}

/**
 * Example 7: Conditional Logging
 * Only log in development mode
 */
async function example7_ConditionalLogging() {
  console.log('\n=== Example 7: Conditional Logging ===');

  const isDevelopment = process.env.NODE_ENV !== 'production';

  const doc = Document.create({
    logger: isDevelopment
      ? new ConsoleLogger(LogLevel.DEBUG)
      : new SilentLogger(),
  });

  doc.createParagraph('Conditional logging example');

  await doc.save(path.join(OUTPUT_DIR, 'conditional-logging.docx'));
  console.log(`✓ Saved with ${isDevelopment ? 'verbose' : 'silent'} logging`);
}

/**
 * Example 8: Logging with Multiple Documents
 * Use different loggers for different documents
 */
async function example8_MultipleDocuments() {
  console.log('\n=== Example 8: Multiple Documents with Different Loggers ===');

  // Important document - collect all logs
  const importantLogger = new CollectingLogger();
  const importantDoc = Document.create({ logger: importantLogger });
  importantDoc.createParagraph('Important document content');
  await importantDoc.save(path.join(OUTPUT_DIR, 'important-doc.docx'));

  // Routine document - silent logging
  const routineDoc = Document.create({ logger: new SilentLogger() });
  routineDoc.createParagraph('Routine document content');
  await routineDoc.save(path.join(OUTPUT_DIR, 'routine-doc.docx'));

  console.log(`Important doc logs: ${importantLogger.getCount()}`);
  console.log('✓ Multiple documents with independent logging');
}

/**
 * Main execution
 */
async function main() {
  console.log('docXMLater - Logging Examples\n');

  // Create output directory
  const fs = require('fs');
  if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  }

  try {
    await example1_DefaultLogging();
    await example2_SilentLogging();
    await example3_VerboseLogging();
    await example4_CollectingLogs();
    await example5_CustomLogger();
    await example6_ScopedLogger();
    await example7_ConditionalLogging();
    await example8_MultipleDocuments();

    console.log('\n✓ All logging examples completed successfully!');
    console.log(`Output directory: ${OUTPUT_DIR}`);
  } catch (error) {
    console.error('Error running examples:', error);
    process.exit(1);
  }
}

// Run if executed directly
if (require.main === module) {
  main();
}

export {
  example1_DefaultLogging,
  example2_SilentLogging,
  example3_VerboseLogging,
  example4_CollectingLogs,
  example5_CustomLogger,
  example6_ScopedLogger,
  example7_ConditionalLogging,
  example8_MultipleDocuments,
};
