import { describe, expect, it } from 'vitest';
import { pagekit } from '../src/index.js';
import { mount } from './helpers.js';

describe('no network and no scripts', () => {
  it('does not write innerHTML or evaluate script text from the input', () => {
    const setter = Object.getOwnPropertyDescriptor(Element.prototype, 'innerHTML');
    if (!setter?.set || !setter.get) throw new Error('innerHTML descriptor missing');
    let writes = 0;
    Object.defineProperty(Element.prototype, 'innerHTML', {
      configurable: true,
      get() {
        return setter.get?.call(this);
      },
      set(value: string) {
        writes += 1;
        setter.set?.call(this, value);
      },
    });
    try {
      mount(`
        <div id="safe">
          <p>Visible line</p>
          <script>globalThis.__pagekitRan = true</script>
          <img alt="sprig" src="https://example.invalid/sprig.png" onerror="globalThis.__pagekitRan = true">
          <a href="javascript:globalThis.__pagekitRan = true">do not run</a>
        </div>
      `);
      const before = writes;
      const tree = pagekit.parse('safe');
      expect(writes).toBe(before);
      const text = JSON.stringify(tree);
      expect(text).toContain('Visible line');
      expect(text).toContain('sprig');
      expect(text).toContain('do not run');
      expect(text).not.toContain('__pagekitRan = true');
      expect(tree.skipped.some((item) => item.kind === 'script')).toBe(true);
      expect(tree.skipped.some((item) => item.kind === 'unsafe-link')).toBe(true);
      expect(tree.skipped.some((item) => item.kind === 'image')).toBe(true);
      expect((globalThis as { __pagekitRan?: boolean }).__pagekitRan).toBeUndefined();
    } finally {
      Object.defineProperty(Element.prototype, 'innerHTML', setter);
    }
  });
});
