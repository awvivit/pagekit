# Pagekit

Pagekit turns an HTML element into a Word `.docx` file. It runs in the browser. It does not call a server, and it does not run scripts from the HTML you give it.

Phase 1 reads computed CSS and writes one `.docx`. The stack is [parse5](https://github.com/inikulin/parse5) 8.1.1 and [OpenXmlSdkTs](https://github.com/EricWhiteDev/OpenXmlSdkTs) 1.0.13, both pinned. OpenXmlSdkTs opens a small seed package and writes the WordprocessingML parts. JSZip (used under its MIT license) builds that seed. See `THIRD_PARTY_NOTICES.md`.

## API

```ts
import { pagekit } from 'pagekit';

const tree = pagekit.parse(document.getElementById('page') ?? 'page', {
  styles: { fontFamily: 'Georgia', fontSize: 12, color: '#292524' },
});

const { blob, skipped } = await pagekit.write(tree, 'docx', {
  template: {
    styles: { fontFamily: 'Calibri', fontSize: 11 },
    headings: { 1: { fontSize: 24, color: '#1F4E3D' } },
  },
});
```

`parse` accepts an element or the id of an element in the current document. An unknown id throws `PagekitError` and does not produce a file. The tree is plain data: lengths are in points, colors are `#RRGGBB`.

`write` accepts the format `'docx'` and returns a `Blob` plus a list of anything it could not represent. Other formats throw, and no file is produced.

### Style precedence

For each property, the winner is the first layer that sets it, from highest to lowest:

1. Element CSS. Pagekit reads computed styles and keeps a property when an author stylesheet or inline style sets it (including inheritance, unless a user-agent rule such as link blue overrides that inheritance). `b`, `strong`, `i`, `em`, `u`, and links also carry their usual HTML formatting at this layer.
2. The `styles` object passed to `parse`.
3. The `template` passed to `write`. Built-in Calibri body text and Heading 1–6 definitions fill anything the template leaves out.

Headings `h1`–`h6` become Word Heading 1–6. Paragraphs, bold, italic, underline, hyperlinks, nested bulleted and numbered lists, and simple tables are written. Merged cells are not written; their text is kept as paragraphs and reported in `skipped`.

Images, page breaks, flex and grid layout, headers and footers, comments, tracked changes, forms, and scripts are not written as those features. Their visible text is kept, the call does not throw, and `skipped` names each one. Scripts are not executed and their source is not copied into the document.

## Demo

```bash
npm install
npm run demo
```

Open the printed local URL. The page is a fictional Harbor & Pine Library bulletin with an **Export to Word** button. `npm run build:demo` writes a static site to `demo/dist` for a later GitHub Pages deploy. This repo does not enable Pages.

## Checks

```bash
npm run lint
npm run typecheck
npm test
```

`npm test` builds a `.docx`, checks `word/document.xml`, `word/styles.xml`, `word/numbering.xml`, and `word/settings.xml` against the ECMA-376 transitional schemas with `xmllint`, and converts the file to text with headless LibreOffice to confirm the wording stays in HTML order. CI runs lint, typecheck, unit tests, and that validation. See `.github/workflows/ci.yml`.

Opening the file in desktop Word is a follow-up human check. LibreOffice and the schema check are what this phase automates.

## License

[MIT](LICENSE). Third-party notices are in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
