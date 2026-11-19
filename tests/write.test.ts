import { describe, expect, it, vi } from 'vitest';
import { PagekitError, pagekit } from '../src/index.js';
import { documentText, mount, packageParts, phrasesInOrder } from './helpers.js';

const sample = `
  <article id="bulletin">
    <h1>Harbor &amp; Pine</h1>
    <p style="text-align: center; color: #1F4E79; font-family: Georgia; font-size: 14pt; margin-top: 10pt; margin-bottom: 6pt; line-height: 20pt">
      The <b>reading room</b> serves <i>cedar tea</i> and keeps a <u>quiet hour</u>.
      See the <a href="https://example.com/harbor-pine/hours">hours</a>.
    </p>
    <h2>Trails</h2>
    <ul>
      <li>Cedar loop
        <ul><li>Look for the painted stone</li></ul>
      </li>
      <li>North meadow</li>
    </ul>
    <ol>
      <li>Sign in
        <ol><li>Leave a note for Mara</li></ol>
      </li>
      <li>Choose a book</li>
    </ol>
    <table>
      <tr><th>Day</th><th>Desk</th></tr>
      <tr><td>Monday</td><td>Mara</td></tr>
      <tr><td>Tuesday</td><td>Jonah</td></tr>
    </table>
    <p>Fish &amp; chips &lt;fresh&gt;.</p>
  </article>
`;

describe('write docx', () => {
  it('maps headings, runs, links, lists, tables, and CSS into Word XML', async () => {
    const root = mount(sample);
    const tree = pagekit.parse(root, { styles: { fontFamily: 'Palatino', fontSize: 13 } });
    const result = await pagekit.write(tree, 'docx', {
      template: {
        styles: { fontFamily: 'Courier New', fontSize: 22, color: '#445566', alignment: 'right', spaceAfter: 3 },
        headings: { 1: { fontSize: 28, color: '#102030' } },
      },
    });

    const parts = await packageParts(result.blob);
    const document = parts.get('word/document.xml') ?? '';
    const rels = parts.get('word/_rels/document.xml.rels') ?? '';
    const text = await documentText(result.blob);

    expect(document).toContain("w:val='Heading1'");
    expect(document).toContain("w:val='Heading2'");
    expect(document).toContain('<w:b');
    expect(document).toContain('<w:i');
    expect(document).toContain("w:val='single'");
    expect(rels).toContain("Target='https://example.com/harbor-pine/hours'");
    expect(rels).toContain("TargetMode='External'");
    expect(document).toContain('<w:hyperlink ');
    expect(document).toContain('<w:numPr>');
    expect(document).toContain('<w:tbl>');
    expect(document).not.toContain('w:gridSpan');
    expect(text).toContain('Harbor & Pine');
    expect(text).toContain('Fish & chips <fresh>.');
    phrasesInOrder(text, [
      'Harbor & Pine',
      'reading room',
      'cedar tea',
      'quiet hour',
      'hours',
      'Trails',
      'Cedar loop',
      'painted stone',
      'North meadow',
      'Sign in',
      'Leave a note for Mara',
      'Choose a book',
      'Day',
      'Desk',
      'Monday',
      'Mara',
      'Tuesday',
      'Jonah',
      'Fish & chips <fresh>.',
    ]);

    const readingRoom = document.indexOf('reading room');
    const centered = document.slice(Math.max(0, readingRoom - 1500), readingRoom);
    expect(centered).toContain("w:val='center'");
    expect(centered).toContain("w:ascii='Georgia'");
    expect(centered).toContain("w:val='1F4E79'");
    expect(centered).toContain("w:val='28'");
    expect(centered).toContain("w:before='200'");
    expect(centered).toContain("w:after='120'");
    expect(centered).toContain("w:line='400'");

    const fish = document.indexOf('Fish');
    const plain = document.slice(Math.max(0, fish - 600), fish);
    expect(plain).toContain("w:ascii='Palatino'");
    expect(plain).toContain("w:val='26'");
  });

  it('lets the style object beat the template when the element does not set the property', async () => {
    mount('<p id="plain">Only defaults.</p>');
    const tree = pagekit.parse('plain', { styles: { fontFamily: 'Georgia', color: '#ABCDEF' } });
    const result = await pagekit.write(tree, 'docx', {
      template: { styles: { fontFamily: 'Courier New', color: '#111111', alignment: 'right' } },
    });
    const document = (await packageParts(result.blob)).get('word/document.xml') ?? '';
    expect(document).toContain("w:ascii='Georgia'");
    expect(document).toContain("w:val='ABCDEF'");
    expect(document).toContain("w:val='right'");
    expect(document).not.toContain('Courier New');
  });

  it('rejects an unknown format before producing a file', async () => {
    mount('<p id="plain">Hello</p>');
    const tree = pagekit.parse('plain');
    let blob: Blob | undefined;
    await expect(pagekit.write(tree, 'pdf').then((result) => {
      blob = result.blob;
    })).rejects.toBeInstanceOf(PagekitError);
    expect(blob).toBeUndefined();
  });

  it('flattens merged cells, keeps unsupported text, and lists what it skipped', async () => {
    mount(`
      <div id="mixed">
        <p>Before</p>
        <img alt="Pine sprig">
        <div style="display: flex"><span>Flex note</span><span>kept</span></div>
        <div style="display: grid">Grid note</div>
        <p style="page-break-before: always">After the break</p>
        <table>
          <tr><td colspan="2">Wide cell</td></tr>
          <tr><td>Left</td><td rowspan="1">Right</td></tr>
        </table>
        <header>Masthead</header>
        <footer>Printed in town</footer>
        <!-- desk note -->
        <del>old hours</del><ins>new hours</ins>
        <form><label>Name <input value="Mara"></label><button>Save</button></form>
        <script>globalThis.__pagekitRan = true</script>
        <p>After</p>
      </div>
    `);
    const fetchSpy = vi.spyOn(globalThis, 'fetch');
    const tree = pagekit.parse('mixed');
    const result = await pagekit.write(tree, 'docx');
    const text = await documentText(result.blob);
    phrasesInOrder(text, ['Before', 'Pine sprig', 'Flex note', 'kept', 'Grid note', 'After the break', 'Wide cell', 'Left', 'Right', 'Masthead', 'Printed in town', 'old hours', 'new hours', 'Mara', 'Save', 'After']);
    expect(text).not.toContain('__pagekitRan');
    expect(text).not.toContain('globalThis');
    const kinds = result.skipped.map((item) => item.kind);
    for (const kind of ['image', 'flex', 'grid', 'page-break', 'merged-cells', 'header', 'footer', 'comment', 'tracked-change', 'form', 'script']) {
      expect(kinds).toContain(kind);
    }
    expect(fetchSpy).not.toHaveBeenCalled();
    fetchSpy.mockRestore();
    expect((globalThis as { __pagekitRan?: boolean }).__pagekitRan).toBeUndefined();
  });
});
