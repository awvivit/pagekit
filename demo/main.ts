import { pagekit } from '../src/index.js';

const button = document.querySelector<HTMLButtonElement>('#export');
const status = document.querySelector<HTMLElement>('#status');

button?.addEventListener('click', () => {
  void exportBulletin();
});

async function exportBulletin(): Promise<void> {
  if (!button || !status) return;
  button.disabled = true;
  status.textContent = 'Building the document…';
  try {
    const tree = pagekit.parse('page', {
      styles: {
        fontFamily: 'Georgia',
        fontSize: 12,
        color: '#292524',
      },
    });
    const result = await pagekit.write(tree, 'docx', {
      template: {
        styles: {
          fontFamily: 'Calibri',
          fontSize: 11,
          color: '#1c1917',
        },
        headings: {
          1: { fontFamily: 'Calibri', color: '#1F4E3D' },
          2: { fontFamily: 'Calibri', color: '#1F4E3D' },
        },
      },
    });
    const url = URL.createObjectURL(result.blob);
    const download = document.createElement('a');
    download.href = url;
    download.download = 'harbor-pine-bulletin.docx';
    download.click();
    URL.revokeObjectURL(url);
    const skipped = result.skipped.length
      ? `Skipped: ${result.skipped.map((item) => item.kind).join(', ')}.`
      : 'Nothing was skipped.';
    status.textContent = `Saved harbor-pine-bulletin.docx. ${skipped}`;
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Export failed.';
    status.textContent = message;
  } finally {
    button.disabled = false;
  }
}
