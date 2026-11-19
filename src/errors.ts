/** Error thrown when Pagekit refuses to continue. No file is produced. */
export class PagekitError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'PagekitError';
  }
}
