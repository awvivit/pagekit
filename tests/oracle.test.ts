import { describe, expect, it } from 'vitest';
import { pagekit } from '../src/index.js';
import { assertWellFormedOfficeXml, documentText, libreOfficeText, mount, packageParts, phrasesInOrder } from './helpers.js';

const story = `
  <article id="story">
    <h1>Harbor &amp; Pine bulletin</h1>
    <p>The reading room opens onto the cedar trail.</p>
    <ul><li>Cedar loop<ul><li>Painted stone</li></ul></li></ul>
    <ol><li>Sign the ledger</li><li>Choose a book</li></ol>
    <table><tr><th>Day</th><th>Desk</th></tr><tr><td>Monday</td><td>Mara</td></tr></table>
    <p>See <a href="https://example.com/harbor-pine">the made-up hours page</a>.</p>
  </article>
`;

describe('docx opens and validates', () => {
  it('matches the ECMA-376 transitional schema', async () => {
    mount(story);
    const result = await pagekit.write(pagekit.parse('story'), 'docx');
    const parts = await packageParts(result.blob);
    expect(() => assertWellFormedOfficeXml(parts)).not.toThrow();
    expect(await documentText(result.blob)).toContain('Harbor & Pine bulletin');
  });

  it('rejects markup that breaks the schema, so the check is real', () => {
    const broken = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:notARealElement/>
      </w:document>`;
    expect(() => assertWellFormedOfficeXml(new Map([['word/document.xml', broken], ['word/styles.xml', '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'], ['word/numbering.xml', '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'], ['word/settings.xml', '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>']]))).toThrow();
  });

  it('converts in LibreOffice with the text in HTML order', async () => {
    mount(story);
    const result = await pagekit.write(pagekit.parse('story'), 'docx');
    const text = libreOfficeText(Buffer.from(await result.blob.arrayBuffer()), 'harbor-pine.docx');
    phrasesInOrder(text, [
      'Harbor & Pine bulletin',
      'reading room opens onto the cedar trail',
      'Cedar loop',
      'Painted stone',
      'Sign the ledger',
      'Choose a book',
      'Monday',
      'Mara',
      'made-up hours page',
    ]);
    expect(text.length).toBeGreaterThan(0);
  }, 180000);
});
