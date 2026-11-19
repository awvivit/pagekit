import { W, WmlPackage } from 'openxmlsdkts';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import JSZip from 'jszip';

const schemaDir = join(dirname(fileURLToPath(import.meta.url)), '../schemas/ecma-376');

export function mount(html: string, css = ''): HTMLElement {
  document.head.replaceChildren();
  document.body.replaceChildren();
  if (css) {
    const style = document.createElement('style');
    style.textContent = css;
    document.head.appendChild(style);
  }
  const host = document.createElement('div');
  host.innerHTML = html;
  document.body.appendChild(host);
  const root = host.firstElementChild;
  if (!root) throw new Error('fixture needs a root element');
  return root as HTMLElement;
}

export async function packageParts(blob: Blob): Promise<Map<string, string>> {
  const zip = await JSZip.loadAsync(await blob.arrayBuffer());
  const parts = new Map<string, string>();
  for (const [name, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    parts.set(name, await file.async('string'));
  }
  return parts;
}

export async function documentText(blob: Blob): Promise<string> {
  const pkg = await WmlPackage.open(blob);
  const main = await pkg.mainDocumentPart();
  if (!main) throw new Error('missing main document');
  const xml = await main.getXDocument();
  return xml.descendants(W.t).map((node) => node.value).join('');
}

export function assertWellFormedOfficeXml(parts: Map<string, string>): void {
  const dir = mkdtempSync(join(tmpdir(), 'pagekit-xsd-'));
  for (const name of ['word/document.xml', 'word/styles.xml', 'word/numbering.xml', 'word/settings.xml']) {
    const xml = parts.get(name);
    if (!xml) throw new Error(`missing ${name}`);
    const path = join(dir, name.replaceAll('/', '_'));
    writeFileSync(path, xml);
    execFileSync('xmllint', ['--noout', '--nonet', '--schema', join(schemaDir, 'wml.xsd'), path], {
      env: { ...process.env, XML_CATALOG_FILES: join(schemaDir, 'catalog.xml') },
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
    });
  }
}

export function libreOfficeText(blob: Buffer, filename: string): string {
  const dir = mkdtempSync(join(tmpdir(), 'pagekit-lo-'));
  const profile = mkdtempSync(join(tmpdir(), 'pagekit-lo-profile-'));
  const docxPath = join(dir, filename);
  writeFileSync(docxPath, blob);
  execFileSync(
    'soffice',
    [
      `-env:UserInstallation=file://${profile}`,
      '--headless',
      '--norestore',
      '--convert-to',
      'txt:Text',
      '--outdir',
      dir,
      docxPath,
    ],
    { encoding: 'utf8', timeout: 120000 },
  );
  return readFileSync(join(dir, filename.replace(/\.docx$/, '.txt')), 'utf8');
}

export function phrasesInOrder(text: string, phrases: string[]): void {
  let from = 0;
  for (const phrase of phrases) {
    const found = text.indexOf(phrase, from);
    if (found < 0) {
      throw new Error(`Missing "${phrase}" after index ${from} in:\n${text}`);
    }
    from = found + phrase.length;
  }
}
