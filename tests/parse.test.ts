import { describe, expect, it } from 'vitest';
import { PagekitError, pagekit } from '../src/index.js';
import { mount } from './helpers.js';

describe('parse', () => {
  it('accepts an element or its id', () => {
    const root = mount('<section id="bulletin"><h1>Harbor notes</h1><p>Cedar tea at four.</p></section>');
    const fromElement = pagekit.parse(root);
    const fromId = pagekit.parse('bulletin');
    expect(fromElement.blocks.map((block) => block.type)).toEqual(['paragraph', 'paragraph']);
    expect(fromId.blocks).toEqual(fromElement.blocks);
    const heading = fromElement.blocks[0];
    expect(heading?.type).toBe('paragraph');
    if (heading?.type === 'paragraph') expect(heading.headingLevel).toBe(1);
  });

  it('throws a clear error for an unknown id and does not produce a file', async () => {
    let blob: Blob | undefined;
    let caught: unknown;
    try {
      const tree = pagekit.parse('missing-bulletin');
      blob = (await pagekit.write(tree, 'docx')).blob;
    } catch (error) {
      caught = error;
    }
    expect(caught).toBeInstanceOf(PagekitError);
    expect((caught as Error).message).toMatch(/missing-bulletin/);
    expect((caught as Error).message).toMatch(/No document was produced/);
    expect(blob).toBeUndefined();
  });

  it('resolves author CSS to points and hex, with element CSS above the style object', () => {
    const root = mount(
      '<p id="sample" style="color: #112233; margin-bottom: 6pt">Hello</p>',
      '#sample { font-size: 18pt; text-align: center; line-height: 22pt; }',
    );
    const tree = pagekit.parse(root, {
      styles: {
        color: '#00ff00',
        fontFamily: 'Georgia',
        fontSize: 15,
        alignment: 'right',
        spaceAfter: 2,
      },
    });
    expect(tree.styles).toMatchObject({
      color: '#00FF00',
      fontFamily: 'Georgia',
      fontSizePt: 15,
      alignment: 'right',
      spaceAfterPt: 2,
    });
    const paragraph = tree.blocks[0];
    expect(paragraph?.type).toBe('paragraph');
    if (paragraph?.type !== 'paragraph') return;
    expect(paragraph.style).toMatchObject({
      alignment: 'center',
      spaceAfterPt: 6,
      lineSpacingPt: 22,
    });
    const text = paragraph.inlines[0];
    expect(text?.type).toBe('text');
    if (text?.type !== 'text') return;
    expect(text.style.color).toBe('#112233');
    expect(text.style.fontSizePt).toBe(18);
    expect(text.style.fontFamily).toBeUndefined();
    expect(JSON.parse(JSON.stringify(tree))).toEqual(tree);
  });

  it('keeps semantic bold, italic, underline, and link color unless CSS overrides them', () => {
    const root = mount('<p>Say <b>yes</b>, <i>maybe</i>, <u>done</u>, and <a href="https://example.com/harbor">go</a>. <b style="font-weight: 400">plain</b></p>');
    const paragraph = pagekit.parse(root).blocks[0];
    if (paragraph?.type !== 'paragraph') throw new Error('expected a paragraph');
    const texts = paragraph.inlines.filter((inline) => inline.type === 'text' || inline.type === 'hyperlink');
    const bold = paragraph.inlines.find((inline) => inline.type === 'text' && inline.text === 'yes');
    const italic = paragraph.inlines.find((inline) => inline.type === 'text' && inline.text === 'maybe');
    const underline = paragraph.inlines.find((inline) => inline.type === 'text' && inline.text === 'done');
    const link = paragraph.inlines.find((inline) => inline.type === 'hyperlink');
    const plain = paragraph.inlines.find((inline) => inline.type === 'text' && inline.text === 'plain');
    expect(bold?.type === 'text' && bold.style.bold).toBe(true);
    expect(italic?.type === 'text' && italic.style.italic).toBe(true);
    expect(underline?.type === 'text' && underline.style.underline).toBe(true);
    expect(link?.type).toBe('hyperlink');
    if (link?.type === 'hyperlink') {
      expect(link.href).toBe('https://example.com/harbor');
      expect(link.style.color).toBe('#0563C1');
      expect(link.style.underline).toBe(true);
    }
    expect(plain?.type === 'text' && plain.style.bold).toBe(false);
    expect(texts.length).toBeGreaterThan(0);
  });

  it('builds nested lists and a rectangular table', () => {
    const root = mount(`
      <div id="notes">
        <ul>
          <li>Apples<ul><li>Honeycrisp</li></ul></li>
          <li>Pears</li>
        </ul>
        <ol start="3"><li>First step</li></ol>
        <table>
          <tr><th>Day</th><th>Hours</th></tr>
          <tr><td>Monday</td><td>9 to 5</td></tr>
        </table>
      </div>
    `);
    const tree = pagekit.parse('notes');
    const list = tree.blocks.find((block) => block.type === 'list');
    expect(list?.type).toBe('list');
    if (list?.type === 'list') {
      expect(list.ordered).toBe(false);
      expect(list.items).toHaveLength(2);
      const nested = list.items[0]?.blocks.find((block) => block.type === 'list');
      expect(nested?.type).toBe('list');
    }
    const ordered = tree.blocks.filter((block) => block.type === 'list')[1];
    if (ordered?.type === 'list') expect(ordered.start).toBe(3);
    const table = tree.blocks.find((block) => block.type === 'table');
    expect(table?.type).toBe('table');
    if (table?.type === 'table') {
      expect(table.rows).toHaveLength(2);
      expect(table.rows[0]?.cells[0]?.header).toBe(true);
      expect(table.rows[1]?.cells).toHaveLength(2);
    }
    expect(root.id).toBe('notes');
  });
});
