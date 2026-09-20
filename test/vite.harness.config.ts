import { defineConfig } from 'vite';
import { resolve } from 'node:path';

export default defineConfig({
  build: {
    lib: {
      entry: resolve(__dirname, 'harness-entry.ts'),
      formats: ['es'],
      fileName: () => 'harness.mjs',
    },
    outDir: resolve(__dirname, 'build'),
    emptyOutDir: true,
    target: 'node18',
    minify: false,
    rollupOptions: { external: [] },
  },
});
