/**
 * Lightweight event system for Document lifecycle and content mutations.
 *
 * Designed for batch processing pipelines that need to react to programmatic
 * edits (e.g., audit logging a `paragraphAdded` for every paragraph the
 * pipeline appends). Not a full reactive store — there is no automatic
 * dirty tracking and listeners are synchronous.
 *
 * Event guarantees:
 * - `beforeSave` / `beforeLoad` fire **inside** the save/load lock, so
 *   listeners observe a stable snapshot of in-memory state.
 * - `afterSave` / `afterLoad` fire after the operation completes
 *   successfully. They do not fire on failure.
 * - Mutation events (`paragraphAdded`, `paragraphRemoved`, `tableAdded`)
 *   fire synchronously after the mutation is applied.
 * - Listener errors do not abort the operation — they are caught and
 *   logged via the global logger so a misbehaving listener cannot corrupt
 *   document state. (Throwing from a listener is a programming error.)
 */
import type { Paragraph } from '../elements/Paragraph';
import type { Table } from '../elements/Table';
import { getGlobalLogger } from '../utils/logger';

export interface DocumentEventMap {
  /** Fired after a paragraph is added to the body or a table cell. */
  paragraphAdded: { paragraph: Paragraph; index?: number };
  /** Fired after a paragraph is removed from the document. */
  paragraphRemoved: { paragraph: Paragraph };
  /** Fired after a table is added to the body. */
  tableAdded: { table: Table };
  /** Fired immediately before save() / toBuffer() begins generation. */
  beforeSave: { filePath?: string };
  /** Fired immediately after save() / toBuffer() completes successfully. */
  afterSave: { filePath?: string; bufferSize?: number };
  /** Fired immediately before a load() / loadFromBuffer() finishes parsing. */
  beforeLoad: { source: 'file' | 'buffer'; path?: string };
  /** Fired immediately after a load() / loadFromBuffer() completes. */
  afterLoad: { source: 'file' | 'buffer'; path?: string };
}

export type DocumentEventType = keyof DocumentEventMap;

export type DocumentEventListener<T extends DocumentEventType> = (
  payload: DocumentEventMap[T]
) => void;

/**
 * Internal emitter implementation. Stored on each Document instance.
 */
export class DocumentEventEmitter {
  private listeners = new Map<DocumentEventType, Set<DocumentEventListener<DocumentEventType>>>();

  on<T extends DocumentEventType>(event: T, listener: DocumentEventListener<T>): () => void {
    let set = this.listeners.get(event);
    if (!set) {
      set = new Set();
      this.listeners.set(event, set);
    }
    set.add(listener as DocumentEventListener<DocumentEventType>);
    return () => this.off(event, listener);
  }

  off<T extends DocumentEventType>(event: T, listener: DocumentEventListener<T>): void {
    const set = this.listeners.get(event);
    if (set) {
      set.delete(listener as DocumentEventListener<DocumentEventType>);
    }
  }

  emit<T extends DocumentEventType>(event: T, payload: DocumentEventMap[T]): void {
    const set = this.listeners.get(event);
    if (!set || set.size === 0) return;
    for (const listener of set) {
      try {
        (listener as DocumentEventListener<T>)(payload);
      } catch (err) {
        getGlobalLogger().warn('Document event listener threw', {
          event,
          error: String(err),
        });
      }
    }
  }

  removeAllListeners(event?: DocumentEventType): void {
    if (event) {
      this.listeners.delete(event);
    } else {
      this.listeners.clear();
    }
  }

  listenerCount(event: DocumentEventType): number {
    return this.listeners.get(event)?.size ?? 0;
  }
}
