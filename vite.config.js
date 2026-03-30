const { defineConfig } = require('vite');
const react = require('@vitejs/plugin-react');
const path = require('path');
const fs = require('fs');

// Read SSL certificates
const certPath = path.resolve(__dirname, 'certs/localhost.pem');
const keyPath = path.resolve(__dirname, 'certs/localhost-key.pem');

if (!fs.existsSync(certPath) || !fs.existsSync(keyPath)) {
  throw new Error(`SSL certificates not found. Expected:\n  ${certPath}\n  ${keyPath}`);
}

const httpsConfig = {
  cert: fs.readFileSync(certPath),
  key: fs.readFileSync(keyPath),
};

module.exports = defineConfig({
  plugins: [react.default()],
  root: 'src/ui',
  publicDir: 'public',
  build: {
    outDir: '../../dist',
    emptyOutDir: true,
    rollupOptions: {
      input: {
        index: path.resolve(__dirname, 'src/ui/index.html'),
        commands: path.resolve(__dirname, 'src/ui/commands.html'),
      },
      output: {
        manualChunks(id) {
          if (!id.includes('node_modules')) return;
          if (id.includes('/react/') || id.includes('/react-dom/')) return 'react';
          if (id.includes('/@fluentui/')) return 'fluentui';
          if (id.includes('/react-markdown/') || id.includes('/remark-gfm/') || id.includes('/mdast-') || id.includes('/micromark') || id.includes('/unified/') || id.includes('/remark-') || id.includes('/rehype-')) return 'markdown';
          if (id.includes('/zod/')) return 'zod';
          if (id.includes('/@opencode-ai/sdk/')) return 'opencode-sdk';
          return 'vendor';
        },
      },
    },
  },
  server: {
    port: 52390,
    strictPort: true,
    https: httpsConfig,
    hmr: {
      protocol: 'wss',
      host: 'localhost',
      port: 52390,
    },
  },
});
