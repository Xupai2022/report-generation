import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'node:path';

const rootDir = __dirname;
const outDir = path.resolve(rootDir, '../mss_ai_ppt_sample_assets/backend/frontend');

export default defineConfig({
  plugins: [react()],
  root: rootDir,
  publicDir: path.resolve(rootDir, 'public'),
  build: {
    outDir,
    emptyOutDir: false,
    sourcemap: false,
    rollupOptions: {
      input: {
        index: path.resolve(rootDir, 'index.html'),
        login: path.resolve(rootDir, 'login.html'),
        admin: path.resolve(rootDir, 'admin.html'),
      },
      output: {
        entryFileNames: 'assets/[name]-[hash].js',
        chunkFileNames: 'assets/[name]-[hash].js',
        assetFileNames: 'assets/[name]-[hash][extname]',
      },
    },
  },
  server: {
    fs: {
      allow: [path.resolve(rootDir, '..')],
    },
  },
});
