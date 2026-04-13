import { DocumentValidator } from '../../src/core/DocumentValidator';
import { Document, Paragraph, Table } from '../../src';

describe('DocumentValidator', () => {
  describe('constructor', () => {
    it('should create with default options', () => {
      const validator = new DocumentValidator();
      expect(validator).toBeDefined();
    });

    it('should reject invalid memory percentage', () => {
      expect(() => new DocumentValidator(0)).toThrow(
        'maxMemoryUsagePercent must be between 1 and 100'
      );
      expect(() => new DocumentValidator(101)).toThrow(
        'maxMemoryUsagePercent must be between 1 and 100'
      );
      expect(() => new DocumentValidator(NaN)).toThrow(
        'maxMemoryUsagePercent must be between 1 and 100'
      );
    });

    it('should accept valid memory percentage', () => {
      expect(() => new DocumentValidator(50)).not.toThrow();
      expect(() => new DocumentValidator(1)).not.toThrow();
      expect(() => new DocumentValidator(100)).not.toThrow();
    });

    it('should accept custom options', () => {
      const validator = new DocumentValidator(80, { maxRssMB: 4096, useAbsoluteLimit: false });
      expect(validator).toBeDefined();
    });
  });

  describe('validateProperties', () => {
    it('should accept valid properties', () => {
      const result = DocumentValidator.validateProperties({
        title: 'Test Document',
        creator: 'Test Author',
        revision: 5,
      });
      expect(result.title).toBe('Test Document');
      expect(result.creator).toBe('Test Author');
      expect(result.revision).toBe(5);
    });

    it('should accept empty properties', () => {
      const result = DocumentValidator.validateProperties({});
      expect(result).toEqual({});
    });

    it('should reject non-string title', () => {
      expect(() => DocumentValidator.validateProperties({ title: 123 as any })).toThrow(
        'title must be a string'
      );
    });

    it('should reject overly long strings', () => {
      const longString = 'x'.repeat(10001);
      expect(() => DocumentValidator.validateProperties({ title: longString })).toThrow(
        'exceeds maximum length'
      );
    });

    it('should accept maximum length strings', () => {
      const maxString = 'x'.repeat(10000);
      const result = DocumentValidator.validateProperties({ title: maxString });
      expect(result.title).toBe(maxString);
    });

    it('should reject non-integer revision', () => {
      expect(() => DocumentValidator.validateProperties({ revision: 1.5 })).toThrow(
        'revision must be an integer'
      );
    });

    it('should reject negative revision', () => {
      expect(() => DocumentValidator.validateProperties({ revision: -1 })).toThrow(
        'revision must be between 0'
      );
    });

    it('should accept zero revision', () => {
      const result = DocumentValidator.validateProperties({ revision: 0 });
      expect(result.revision).toBe(0);
    });

    it('should validate date properties', () => {
      const date = new Date('2024-01-15');
      const result = DocumentValidator.validateProperties({ created: date });
      expect(result.created).toBe(date);
    });

    it('should reject invalid date', () => {
      expect(() => DocumentValidator.validateProperties({ created: new Date('invalid') })).toThrow(
        'invalid date'
      );
    });

    it('should reject non-Date created', () => {
      expect(() => DocumentValidator.validateProperties({ created: '2024-01-15' as any })).toThrow(
        'must be a Date'
      );
    });

    it('should validate all string properties', () => {
      const result = DocumentValidator.validateProperties({
        title: 'Title',
        subject: 'Subject',
        creator: 'Creator',
        keywords: 'Keywords',
        description: 'Description',
        lastModifiedBy: 'Modifier',
      });
      expect(result.title).toBe('Title');
      expect(result.subject).toBe('Subject');
      expect(result.creator).toBe('Creator');
      expect(result.keywords).toBe('Keywords');
      expect(result.description).toBe('Description');
      expect(result.lastModifiedBy).toBe('Modifier');
    });
  });

  describe('estimateSize', () => {
    let doc: Document;
    let validator: DocumentValidator;

    beforeEach(() => {
      doc = Document.create();
      validator = new DocumentValidator();
    });

    afterEach(() => {
      doc?.dispose();
    });

    it('should estimate size for empty document', () => {
      const estimate = validator.estimateSize(doc.getBodyElements(), doc.getImageManager());
      expect(estimate.paragraphs).toBe(0);
      expect(estimate.tables).toBe(0);
      expect(estimate.images).toBe(0);
      expect(estimate.totalEstimatedBytes).toBeGreaterThan(0); // base structure
    });

    it('should count paragraphs', () => {
      doc.addParagraph(new Paragraph().addText('Hello'));
      doc.addParagraph(new Paragraph().addText('World'));

      const estimate = validator.estimateSize(doc.getBodyElements(), doc.getImageManager());
      expect(estimate.paragraphs).toBe(2);
    });

    it('should count tables', () => {
      doc.addTable(new Table(2, 3));

      const estimate = validator.estimateSize(doc.getBodyElements(), doc.getImageManager());
      expect(estimate.tables).toBe(1);
    });

    it('should return size in MB', () => {
      const estimate = validator.estimateSize(doc.getBodyElements(), doc.getImageManager());
      expect(typeof estimate.totalEstimatedMB).toBe('number');
      expect(estimate.totalEstimatedMB).toBeGreaterThanOrEqual(0);
    });
  });

  describe('getSizeStats', () => {
    let doc: Document;
    let validator: DocumentValidator;

    beforeEach(() => {
      doc = Document.create();
      validator = new DocumentValidator();
    });

    afterEach(() => {
      doc?.dispose();
    });

    it('should return formatted size stats', () => {
      doc.addParagraph(new Paragraph().addText('Test'));

      const stats = validator.getSizeStats(doc.getBodyElements(), doc.getImageManager());

      expect(stats.elements.paragraphs).toBe(1);
      expect(stats.elements.tables).toBe(0);
      expect(typeof stats.size.xml).toBe('string');
      expect(typeof stats.size.total).toBe('string');
      expect(Array.isArray(stats.warnings)).toBe(true);
    });
  });
});
