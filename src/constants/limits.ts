/**
 * Document validation and processing limits
 * These constants define thresholds for memory usage, file sizes, and content limits
 * Based on Microsoft Word compatibility constraints and practical performance testing
 */

const PROPERTY_LIMITS = {
  MAX_STRING_LENGTH: 10000,
  MAX_REVISION: 999999,
} as const;

const FILE_SIZE_LIMITS = {
  WARNING_SIZE_MB: 50,
  ERROR_SIZE_MB: 150,
} as const;

const XML_LIMITS = {
  MAX_PARSE_SIZE_MB: 10,
  MAX_PARSE_SIZE_BYTES: 10 * 1024 * 1024,
} as const;

const MEMORY_LIMITS = {
  DEFAULT_MAX_HEAP_PERCENT: 80,
  DEFAULT_MAX_RSS_MB: 2048,
  DEFAULT_USE_ABSOLUTE_LIMIT: true,
} as const;

const IMAGE_LIMITS = {
  DEFAULT_MAX_IMAGE_COUNT: 20,
  DEFAULT_MAX_TOTAL_SIZE_MB: 100,
  DEFAULT_MAX_SINGLE_SIZE_MB: 20,
} as const;

const SIZE_ESTIMATES = {
  BYTES_PER_PARAGRAPH: 200,
  BYTES_PER_TABLE: 1000,
  BASE_STRUCTURE_BYTES: 50000,
} as const;

/**
 * All limits combined for easy import
 */
export const LIMITS = {
  ...PROPERTY_LIMITS,
  ...FILE_SIZE_LIMITS,
  ...XML_LIMITS,
  ...MEMORY_LIMITS,
  ...IMAGE_LIMITS,
  ...SIZE_ESTIMATES,
} as const;
