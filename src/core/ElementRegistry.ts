/**
 * Plugin extension point for custom OOXML element handlers. When the
 * parser encounters an unknown qualified-name child of the body and a
 * handler is registered for that tag, the handler's `parse()` is invoked
 * and the result becomes a `PreservedElement` carrying the original raw
 * XML (so save round-trips byte-for-byte). Handlers that return objects
 * with a `toXml()` method allow consumers to roundtrip through their
 * own model — see DocumentParser.processBodyElement for the contract.
 *
 * Registration is process-global; clear via `clear()` between tests.
 *
 * @example
 * ```typescript
 * ElementRegistry.register('myco:specialBlock', {
 *   parse: (xml) => new MySpecialBlock(xml),
 *   serialize: (element) => (element as MySpecialBlock).toXML(),
 * });
 * ```
 */
import { KeyedRegistry } from '../utils/KeyedRegistry';

export interface ElementHandler<E = unknown> {
  /** Parse the raw XML fragment for this element into a model object. */
  parse(rawXml: string, context?: ElementParseContext): E;
  /** Serialize the model object back to raw XML. */
  serialize(element: E, context?: ElementSerializeContext): string;
}

export interface ElementParseContext {
  /** Path of the OOXML part being parsed (e.g., 'word/document.xml'). */
  partPath?: string;
}

export interface ElementSerializeContext {
  partPath?: string;
}

class ElementRegistryImpl {
  private readonly inner = new KeyedRegistry<ElementHandler<unknown>>('ElementRegistry');

  register<E>(tag: string, handler: ElementHandler<E>): void {
    this.inner.register(tag, handler as ElementHandler<unknown>);
  }

  unregister(tag: string): boolean {
    return this.inner.unregister(tag);
  }

  has(tag: string): boolean {
    return this.inner.has(tag);
  }

  get<E = unknown>(tag: string): ElementHandler<E> | undefined {
    return this.inner.get(tag) as ElementHandler<E> | undefined;
  }

  /** Returns the list of currently-registered tags (snapshot). */
  registeredTags(): string[] {
    return this.inner.keys();
  }

  /** Remove all handlers — primarily for test isolation. */
  clear(): void {
    this.inner.clear();
  }
}

/** Process-global element registry instance. */
export const ElementRegistry = new ElementRegistryImpl();
