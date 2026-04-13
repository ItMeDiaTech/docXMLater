import { DrawingManager, PreservedDrawing } from '../../src/managers/DrawingManager';
import { Shape } from '../../src/elements/Shape';
import { TextBox } from '../../src/elements/TextBox';

describe('DrawingManager', () => {
  let manager: DrawingManager;

  beforeEach(() => {
    manager = new DrawingManager();
  });

  describe('initial state', () => {
    it('should start empty', () => {
      expect(manager.isEmpty()).toBe(true);
      expect(manager.getCount()).toBe(0);
      expect(manager.getAllDrawings()).toEqual([]);
    });
  });

  describe('addShape / getAllShapes', () => {
    it('should add a shape and return an ID', () => {
      const shape = Shape.create('rect', 914400, 914400);
      const id = manager.addShape(shape);
      expect(id).toBeDefined();
      expect(typeof id).toBe('string');
      expect(manager.getCount()).toBe(1);
    });

    it('should retrieve shapes by type', () => {
      manager.addShape(Shape.create('rect', 914400, 914400));
      manager.addShape(Shape.create('ellipse', 914400, 914400));
      expect(manager.getAllShapes().length).toBe(2);
    });
  });

  describe('addTextBox / getAllTextBoxes', () => {
    it('should add a text box', () => {
      const tb = TextBox.create(914400, 914400);
      const id = manager.addTextBox(tb);
      expect(id).toBeDefined();
      expect(manager.getAllTextBoxes().length).toBe(1);
    });
  });

  describe('addPreservedDrawing / getAllPreservedDrawings', () => {
    it('should add a preserved drawing', () => {
      const preserved: PreservedDrawing = {
        type: 'chart',
        xml: '<c:chart/>',
        relationshipIds: ['rId5'],
        id: '',
      };
      const id = manager.addPreservedDrawing(preserved);
      expect(id).toBeDefined();
      expect(manager.getAllPreservedDrawings().length).toBe(1);
    });

    it('should handle smartart type', () => {
      const preserved: PreservedDrawing = {
        type: 'smartart',
        xml: '<dgm:relIds/>',
        relationshipIds: ['rId1', 'rId2'],
        id: '',
      };
      manager.addPreservedDrawing(preserved);
      expect(manager.getAllPreservedDrawings()[0]!.type).toBe('smartart');
    });
  });

  describe('getDrawing', () => {
    it('should retrieve by ID', () => {
      const shape = Shape.create('rect', 914400, 914400);
      const id = manager.addShape(shape);
      expect(manager.getDrawing(id)).toBe(shape);
    });

    it('should return undefined for unknown ID', () => {
      expect(manager.getDrawing('nonexistent')).toBeUndefined();
    });
  });

  describe('removeDrawing', () => {
    it('should remove and return true', () => {
      const id = manager.addShape(Shape.create('rect', 914400, 914400));
      expect(manager.removeDrawing(id)).toBe(true);
      expect(manager.getCount()).toBe(0);
      expect(manager.getDrawing(id)).toBeUndefined();
    });

    it('should return false for unknown ID', () => {
      expect(manager.removeDrawing('nonexistent')).toBe(false);
    });

    it('should clean up type-specific indices', () => {
      const id = manager.addShape(Shape.create('rect', 914400, 914400));
      manager.removeDrawing(id);
      expect(manager.getAllShapes().length).toBe(0);
    });
  });

  describe('getDrawingType', () => {
    it('should identify shapes', () => {
      const shape = Shape.create('rect', 914400, 914400);
      expect(manager.getDrawingType(shape)).toBe('shape');
    });

    it('should identify text boxes', () => {
      const tb = TextBox.create(914400, 914400);
      expect(manager.getDrawingType(tb)).toBe('textbox');
    });

    it('should identify preserved drawings', () => {
      const preserved: PreservedDrawing = {
        type: 'chart',
        xml: '<c:chart/>',
        relationshipIds: [],
        id: 'p1',
      };
      expect(manager.getDrawingType(preserved)).toBe('preserved');
    });
  });

  describe('clear', () => {
    it('should remove all drawings', () => {
      manager.addShape(Shape.create('rect', 914400, 914400));
      manager.addTextBox(TextBox.create(914400, 914400));
      expect(manager.getCount()).toBe(2);

      manager.clear();
      expect(manager.getCount()).toBe(0);
      expect(manager.isEmpty()).toBe(true);
      expect(manager.getAllShapes()).toEqual([]);
      expect(manager.getAllTextBoxes()).toEqual([]);
    });
  });

  describe('mixed types', () => {
    it('should track different drawing types independently', () => {
      manager.addShape(Shape.create('rect', 914400, 914400));
      manager.addTextBox(TextBox.create(914400, 914400));
      manager.addPreservedDrawing({
        type: 'chart',
        xml: '<c:chart/>',
        relationshipIds: [],
        id: '',
      });

      expect(manager.getCount()).toBe(3);
      expect(manager.getAllShapes().length).toBe(1);
      expect(manager.getAllTextBoxes().length).toBe(1);
      expect(manager.getAllPreservedDrawings().length).toBe(1);
      expect(manager.getAllDrawings().length).toBe(3);
    });
  });
});
