import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

const here = path.dirname(fileURLToPath(import.meta.url));

export default defineConfig({
  plugins: [react()],
  // Câu chữ của vmai, một nguồn duy nhất (lát 1, Q133). Docker chép đúng file này vào cùng chỗ tương đối.
  resolve: { alias: { '@vmai-texts': path.resolve(here, '../vmai/src/vmai/data/doc_codes.json') } },
  build: { chunkSizeWarningLimit: 16000, assetsInlineLimit: 0 },
});
