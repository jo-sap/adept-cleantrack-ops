import path from 'path';
import os from 'os';
import { defineConfig, loadEnv } from 'vite';
import react from '@vitejs/plugin-react';

// Default cache is node_modules/.vite; under OneDrive/SharePoint that folder is often locked during sync, which surfaces as EPERM when Vite clears deps.
const cacheDir = path.join(
  process.env.LOCALAPPDATA ?? os.tmpdir(),
  'vite-cache',
  'adept-timesheet-application',
);

export default defineConfig(({ mode }) => {
    const env = loadEnv(mode, '.', '');
    return {
      cacheDir,
      server: {
        port: 3000,
        host: '0.0.0.0',
      },
      plugins: [react()],
      define: {
        'process.env.API_KEY': JSON.stringify(env.GEMINI_API_KEY),
        'process.env.GEMINI_API_KEY': JSON.stringify(env.GEMINI_API_KEY)
      },
      resolve: {
        alias: {
          '@': path.resolve(__dirname, '.'),
        }
      },
      optimizeDeps: {
        // CJS + large bundle: force stable prebundle in dev (avoids broken dynamic import chunks).
        include: ['xlsx-js-style'],
      },
    };
});
