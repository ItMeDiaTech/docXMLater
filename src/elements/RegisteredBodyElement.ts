/**
 * Body element produced by an `ElementRegistry` handler. Carries the
 * consumer's parsed model object together with the handler reference so
 * the document save flow can call `handler.serialize(model)` to round
 * the element back to XML — letting consumers mutate their own model and
 * have changes survive `doc.save()`.
 *
 * If the handler's `serialize()` throws, the original raw XML is emitted
 * unchanged so a buggy custom serializer cannot corrupt the document.
 */
import type { ElementHandler } from '../core/ElementRegistry.js';
import type { XMLElement } from '../xml/XMLBuilder.js';
import { getGlobalLogger } from '../utils/logger.js';

export class RegisteredBodyElement<E = unknown> {
  constructor(
    private readonly tag: string,
    private readonly model: E,
    private readonly handler: ElementHandler<E>,
    private readonly originalXml: string
  ) {}

  getTag(): string {
    return this.tag;
  }

  getModel(): E {
    return this.model;
  }

  getRawXml(): string {
    return this.originalXml;
  }

  getType(): string {
    return 'registeredBodyElement';
  }

  toXML(): XMLElement {
    let xml: string;
    try {
      xml = this.handler.serialize(this.model);
    } catch (err) {
      getGlobalLogger().warn('ElementRegistry serializer threw; emitting original XML', {
        tag: this.tag,
        error: String(err),
      });
      xml = this.originalXml;
    }
    return { name: '__rawXml', rawXml: xml } as XMLElement;
  }
}
