/**
 * Logging interface for docXMLater
 * Allows library consumers to control logging behavior
 */

/**
 * Log severity levels
 */
export enum LogLevel {
  /** Debugging information (most verbose) */
  DEBUG = 'debug',
  /** Informational messages */
  INFO = 'info',
  /** Warning messages - potential issues that don't prevent operation */
  WARN = 'warn',
  /** Error messages - serious issues that may cause failures */
  ERROR = 'error',
}

/**
 * Log entry structure
 */
export interface LogEntry {
  /** Timestamp when log was created */
  timestamp: Date;
  /** Severity level */
  level: LogLevel;
  /** Log message */
  message: string;
  /** Optional context data */
  context?: Record<string, any>;
  /** Source component that generated the log */
  source?: string;
}

/**
 * Logger interface that consumers can implement
 * Provides full control over how logs are handled
 */
export interface ILogger {
  /**
   * Log a debug message
   * @param message - Debug message
   * @param context - Optional context data
   */
  debug(message: string, context?: Record<string, any>): void;

  /**
   * Log an informational message
   * @param message - Info message
   * @param context - Optional context data
   */
  info(message: string, context?: Record<string, any>): void;

  /**
   * Log a warning message
   * @param message - Warning message
   * @param context - Optional context data
   */
  warn(message: string, context?: Record<string, any>): void;

  /**
   * Log an error message
   * @param message - Error message
   * @param context - Optional context data
   */
  error(message: string, context?: Record<string, any>): void;
}

/**
 * Console-based logger implementation
 * Uses standard console methods for output
 */
export class ConsoleLogger implements ILogger {
  constructor(private minLevel: LogLevel = LogLevel.WARN) {}

  debug(message: string, context?: Record<string, any>): void {
    if (this.shouldLog(LogLevel.DEBUG)) {
      console.debug(this.formatMessage(message, context));
    }
  }

  info(message: string, context?: Record<string, any>): void {
    if (this.shouldLog(LogLevel.INFO)) {
      console.info(this.formatMessage(message, context));
    }
  }

  warn(message: string, context?: Record<string, any>): void {
    if (this.shouldLog(LogLevel.WARN)) {
      console.warn(this.formatMessage(message, context));
    }
  }

  error(message: string, context?: Record<string, any>): void {
    if (this.shouldLog(LogLevel.ERROR)) {
      console.error(this.formatMessage(message, context));
    }
  }

  private shouldLog(level: LogLevel): boolean {
    const levels = [LogLevel.DEBUG, LogLevel.INFO, LogLevel.WARN, LogLevel.ERROR];
    const minIndex = levels.indexOf(this.minLevel);
    const currentIndex = levels.indexOf(level);
    return currentIndex >= minIndex;
  }

  private formatMessage(message: string, context?: Record<string, any>): string {
    if (context && Object.keys(context).length > 0) {
      return `${message} ${JSON.stringify(context)}`;
    }
    return message;
  }
}

/**
 * Silent logger that discards all log messages
 * Useful for testing or when logging is not desired
 */
export class SilentLogger implements ILogger {
  debug(): void {
    // No-op
  }

  info(): void {
    // No-op
  }

  warn(): void {
    // No-op
  }

  error(): void {
    // No-op
  }
}

/**
 * Collecting logger that stores log entries in memory
 * Useful for testing and diagnostics
 */
export class CollectingLogger implements ILogger {
  private logs: LogEntry[] = [];

  debug(message: string, context?: Record<string, any>): void {
    this.addLog(LogLevel.DEBUG, message, context);
  }

  info(message: string, context?: Record<string, any>): void {
    this.addLog(LogLevel.INFO, message, context);
  }

  warn(message: string, context?: Record<string, any>): void {
    this.addLog(LogLevel.WARN, message, context);
  }

  error(message: string, context?: Record<string, any>): void {
    this.addLog(LogLevel.ERROR, message, context);
  }

  private addLog(level: LogLevel, message: string, context?: Record<string, any>): void {
    this.logs.push({
      timestamp: new Date(),
      level,
      message,
      context,
    });
  }

  /**
   * Get all collected log entries
   */
  getLogs(): ReadonlyArray<LogEntry> {
    return [...this.logs];
  }

  /**
   * Get logs filtered by level
   */
  getLogsByLevel(level: LogLevel): ReadonlyArray<LogEntry> {
    return this.logs.filter(log => log.level === level);
  }

  /**
   * Clear all collected logs
   */
  clear(): void {
    this.logs = [];
  }

  /**
   * Get count of logs by level
   */
  getCount(level?: LogLevel): number {
    if (level) {
      return this.logs.filter(log => log.level === level).length;
    }
    return this.logs.length;
  }
}

/**
 * Default logger instance
 * Uses console output with WARN minimum level
 */
export const defaultLogger = new ConsoleLogger(LogLevel.WARN);

/**
 * Creates a scoped logger that adds source information
 * @param logger - Base logger
 * @param source - Source component name
 * @returns Scoped logger with source context
 */
export function createScopedLogger(logger: ILogger, source: string): ILogger {
  return {
    debug(message: string, context?: Record<string, any>): void {
      logger.debug(message, { ...context, source });
    },
    info(message: string, context?: Record<string, any>): void {
      logger.info(message, { ...context, source });
    },
    warn(message: string, context?: Record<string, any>): void {
      logger.warn(message, { ...context, source });
    },
    error(message: string, context?: Record<string, any>): void {
      logger.error(message, { ...context, source });
    },
  };
}
