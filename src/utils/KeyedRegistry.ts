/**
 * Map-backed keyed registry shared by ElementRegistry and ValidationRuleRegistry.
 * Throws on duplicate `register()` so callers cannot silently overwrite a
 * pre-existing entry.
 */
export class KeyedRegistry<V> {
  private readonly entries = new Map<string, V>();

  constructor(private readonly label: string) {}

  register(key: string, value: V): void {
    if (this.entries.has(key)) {
      throw new Error(`${this.label}: "${key}" is already registered`);
    }
    this.entries.set(key, value);
  }

  unregister(key: string): boolean {
    return this.entries.delete(key);
  }

  has(key: string): boolean {
    return this.entries.has(key);
  }

  get(key: string): V | undefined {
    return this.entries.get(key);
  }

  keys(): string[] {
    return [...this.entries.keys()];
  }

  values(): V[] {
    return [...this.entries.values()];
  }

  clear(): void {
    this.entries.clear();
  }
}
