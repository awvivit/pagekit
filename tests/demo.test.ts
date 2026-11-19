import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { pagekit } from '../src/index.js';
import { documentText, phrasesInOrder } from './helpers.js';

describe('demo bulletin', () => {
  it('exports the sample page with text in order and reports the unsupported bits', async () => {
    const css = readFileSync('demo/styles.css', 'utf8');
    const html = readFileSync('demo/index.html', 'utf8').replace(
      '<link rel="stylesheet" href="./styles.css" />',
      `<style>${css}</style>`,
    );
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    document.head.replaceChildren();
    document.body.replaceChildren();
    const style = parsed.querySelector('style');
    const page = parsed.getElementById('page');
    if (!style || !page) throw new Error('demo markup missing');
    document.head.appendChild(document.adoptNode(style));
    document.body.appendChild(document.adoptNode(page));

    const tree = pagekit.parse('page', {
      styles: { fontFamily: 'Georgia', fontSize: 12, color: '#292524' },
    });
    const result = await pagekit.write(tree, 'docx', {
      template: {
        styles: { fontFamily: 'Calibri', fontSize: 11, color: '#1c1917' },
        headings: {
          1: { fontFamily: 'Calibri', color: '#1F4E3D' },
          2: { fontFamily: 'Calibri', color: '#1F4E3D' },
        },
      },
    });
    const text = await documentText(result.blob);
    phrasesInOrder(text, [
      'Harbor & Pine Library',
      'Mara Ellison',
      'quiet hour',
      'example.com/harbor-pine/hours',
      'Cedar loop',
      'painted stone',
      'Sign the ledger',
      'Monday',
      'Jonah Pell',
      'map drawer',
      'pine sprig',
    ]);
    const kinds = result.skipped.map((item) => item.kind);
    expect(kinds).toContain('flex');
    expect(kinds).toContain('image');
    expect(text).not.toContain('data:image');
  });
});
