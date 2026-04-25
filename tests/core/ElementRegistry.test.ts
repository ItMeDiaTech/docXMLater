import { ElementRegistry } from '../../src/core/ElementRegistry';

describe('ElementRegistry', () => {
  beforeEach(() => {
    ElementRegistry.clear();
  });

  it('registers and retrieves a handler by tag', () => {
    const handler = {
      parse: (xml: string) => ({ raw: xml }),
      serialize: (e: { raw: string }) => e.raw,
    };
    ElementRegistry.register('myco:test', handler);
    expect(ElementRegistry.has('myco:test')).toBe(true);
    const got = ElementRegistry.get<{ raw: string }>('myco:test');
    expect(got).toBe(handler);
    expect(got!.parse('<x/>')).toEqual({ raw: '<x/>' });
    expect(got!.serialize({ raw: '<x/>' })).toBe('<x/>');
  });

  it('throws on duplicate registration', () => {
    const handler = { parse: () => null, serialize: () => '' };
    ElementRegistry.register('myco:dup', handler);
    expect(() => ElementRegistry.register('myco:dup', handler)).toThrow(/already registered/);
  });

  it('unregister() removes a handler', () => {
    ElementRegistry.register('myco:rm', { parse: () => null, serialize: () => '' });
    expect(ElementRegistry.unregister('myco:rm')).toBe(true);
    expect(ElementRegistry.has('myco:rm')).toBe(false);
    expect(ElementRegistry.unregister('myco:rm')).toBe(false);
  });

  it('lists all registered tags', () => {
    ElementRegistry.register('a:one', { parse: () => null, serialize: () => '' });
    ElementRegistry.register('b:two', { parse: () => null, serialize: () => '' });
    const tags = ElementRegistry.registeredTags();
    expect(tags.sort()).toEqual(['a:one', 'b:two']);
  });

  it('clear() removes all handlers', () => {
    ElementRegistry.register('a:x', { parse: () => null, serialize: () => '' });
    ElementRegistry.register('b:y', { parse: () => null, serialize: () => '' });
    ElementRegistry.clear();
    expect(ElementRegistry.registeredTags()).toEqual([]);
  });
});
