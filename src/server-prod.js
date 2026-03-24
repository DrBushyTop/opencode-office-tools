/**
 * Production server for Office Add-in
 * This serves the pre-built static files (no Vite dev server)
 */
const express = require('express');
const https = require('https');
const http = require('http');
const path = require('path');
const fs = require('fs');
const { createApiRouter } = require('./server/api');
const { OpencodeRuntime } = require('./server/opencodeRuntime');
const { OfficeToolBridge } = require('./server/officeToolBridge');

// Determine if we're running from pkg bundle
const isPkg = typeof process.pkg !== 'undefined';

// Get the base directory (works both in dev and when packaged)
function getBasePath() {
  // Check if running from Electron tray app
  if (process.env.OPENCODE_OFFICE_BASE_PATH) {
    return process.env.OPENCODE_OFFICE_BASE_PATH;
  }
  if (isPkg) {
    // When packaged, __dirname points to snapshot filesystem
    // The actual files are next to the executable
    return path.dirname(process.execPath);
  }
  return path.resolve(__dirname, '..');
}

const BASE_PATH = getBasePath();

async function createServer() {
  const app = express();
  const runtime = new OpencodeRuntime();
  const bridge = new OfficeToolBridge();
  const apiRouter = createApiRouter(runtime, bridge);
  app.use('/api', apiRouter);

  // ========== Static File Serving ==========
  const distPath = path.join(BASE_PATH, 'dist');
  app.use(express.static(distPath));
  
  // Fallback to index.html for SPA routing
  app.get('*path', (req, res) => {
    res.sendFile(path.join(distPath, 'index.html'));
  });

  // ========== HTTPS Server ==========
  const certPath = path.join(BASE_PATH, 'certs', 'localhost.pem');
  const keyPath = path.join(BASE_PATH, 'certs', 'localhost-key.pem');
  
  if (!fs.existsSync(certPath) || !fs.existsSync(keyPath)) {
    console.error('SSL certificates not found!');
    console.error('Expected:', certPath);
    console.error('Expected:', keyPath);
    process.exit(1);
  }
  
  const httpsConfig = {
    cert: fs.readFileSync(certPath),
    key: fs.readFileSync(keyPath),
  };
  
  const PORT = process.env.PORT || 52390;
  const BRIDGE_PORT = process.env.BRIDGE_PORT || 52391;
  const httpsServer = https.createServer(httpsConfig, app);
  const bridgeServer = http.createServer(app);

  httpsServer.listen(PORT, () => {
    console.log(`OpenCode Office Add-in Server running on https://localhost:${PORT}`);
  });
  bridgeServer.listen(BRIDGE_PORT, '127.0.0.1', () => {
    console.log(`Office bridge available at http://127.0.0.1:${BRIDGE_PORT}/api`);
  });

  return httpsServer;
}

// Export for use by tray app
module.exports = { createServer };

// Run directly if not required as a module
if (require.main === module) {
  createServer().catch(console.error);
}
