import { defineConfig } from 'vite';
import { resolve } from 'path';
import fs from 'fs';

export default defineConfig({
  server: {
    port: 3000,
    https: {
      key: fs.readFileSync('localhost-key.pem'),
      cert: fs.readFileSync('localhost.pem'),
    },
  },
  build: {
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, 'src/taskpane/taskpane.html'),
      },
    },
  },
});
