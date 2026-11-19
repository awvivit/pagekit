import { defineConfig } from 'vitest/config';

export default defineConfig({
  ssr: {
    noExternal: ['openxmlsdkts', 'ltxmlts'],
  },
  test: {
    environment: 'jsdom',
    include: ['tests/**/*.test.ts'],
  },
});
