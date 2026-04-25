/**
 * Plugin extension point for custom OOXML element handlers.
 *
 * Lets consumers register a parser + serializer pair for a tag the
 * framework does not natively understand (e.g., a vendor-specific
 * extension namespace). Registered handlers integrate with the existing
 * passthrough mechanism: when DocumentParser encounters the tag at a
 * known position, it delegates to the registered parser; on save,
 * the serializer reconstitutes the XML.
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
  private handlers = new Map<string, ElementHandler<unknown>>();

  /**
   * Register a handler for `tag` (e.g., `'myco:specialBlock'`). Throws
   * if a handler is already registered for the same tag — caller can
   * call `unregister()` first if intentional.
   */
  register<E>(tag: string, handler: ElementHandler<E>): void {
    if (this.handlers.has(tag)) {
      throw new Error(`ElementRegistry: handler for "${tag}" is already registered`);
    }
    this.handlers.set(tag, handler as ElementHandler<unknown>);
  }

  unregister(tag: string): boolean {
    return this.handlers.delete(tag);
  }

  has(tag: string): boolean {
    return this.handlers.has(tag);
  }

  get<E = unknown>(tag: string): ElementHandler<E> | undefined {
    return this.handlers.get(tag) as ElementHandler<E> | undefined;
  }

  /** Returns the list of currently-registered tags (snapshot). */
  registeredTags(): string[] {
    return [...this.handlers.keys()];
  }

  /** Remove all handlers — primarily for test isolation. */
  clear(): void {
    this.handlers.clear();
  }
}

/** Process-global element registry instance. */
export const ElementRegistry = new ElementRegistryImpl();
