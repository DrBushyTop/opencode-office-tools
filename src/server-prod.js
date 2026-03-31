/**
 * Production server for Office Add-in
 * This serves the pre-built static files (no Vite dev server)
 */
const express = require('express');
const https = require('https');
const http = require('http');
const path = require('path');
const fs = require('fs');
const { z } = require('zod');
const { createApiRouter, createBridgeRouter } = require('./server/api');
const { OpencodeRuntime } = require('./server/opencodeRuntime');
const { OfficeToolBridge } = require('./server/officeToolBridge');
const { writeBridgeToken, removeBridgeToken } = require('./server/bridgeTokenPath');
const { logInfo, logError } = require('./server/devLogger');

// Determine if we're running from pkg bundle
const isPkg = typeof process.pkg !== 'undefined';

const productionServerConfigSchema = z.object({
  PORT: z.union([z.string(), z.number()]),
  BRIDGE_PORT: z.union([z.string(), z.number()]),
  BASE_PATH: z.string().min(1),
  cert: z.instanceof(Buffer),
  key: z.instanceof(Buffer),
});

// Get the base directory (works both in dev and when packaged)
function getBasePath() {
  // Check if running from Electron tray app
  const configuredBasePath = z.string().min(1).safeParse(process.env.OPENCODE_OFFICE_BASE_PATH);
  if (configuredBasePath.success) {
    return configuredBasePath.data;
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
  const certPath = path.join(BASE_PATH, 'certs', 'localhost.pem');
  const keyPath = path.join(BASE_PATH, 'certs', 'localhost-key.pem');

  if (!fs.existsSync(certPath) || !fs.existsSync(keyPath)) {
    console.error('SSL certificates not found!');
    console.error('Expected:', certPath);
    console.error('Expected:', keyPath);
    logError('server', 'SSL certificates not found', { certPath, keyPath });
    process.exit(1);
  }

  const { PORT, BRIDGE_PORT, cert, key } = productionServerConfigSchema.parse({
    PORT: process.env.PORT || 52390,
    BRIDGE_PORT: process.env.BRIDGE_PORT || 52391,
    BASE_PATH,
    cert: fs.readFileSync(certPath),
    key: fs.readFileSync(keyPath),
  });
  const app = express();
  const bridgeApp = express();
  const runtime = new OpencodeRuntime();
  const bridge = new OfficeToolBridge();
  writeBridgeToken(BRIDGE_PORT, bridge.bridgeToken);
  const apiRouter = createApiRouter(runtime, bridge);
  app.use('/api', apiRouter);
  bridgeApp.use('/api', createBridgeRouter(bridge));

  app.use((req, res, next) => {
    const startedAt = Date.now();
    res.on('finish', () => {
      logInfo('static', `${req.method} ${req.originalUrl}`, {
        statusCode: res.statusCode,
        durationMs: Date.now() - startedAt,
      });
    });
    next();
  });

  // ========== Static File Serving ==========
  const distPath = path.join(BASE_PATH, 'dist');
  app.use(express.static(distPath));
  
  // Fallback to index.html for SPA routing
  app.get('*path', (req, res) => {
    logInfo('static', 'Serving SPA fallback', { url: req.originalUrl, distPath });
    res.sendFile(path.join(distPath, 'index.html'));
  });

  // ========== HTTPS Server ==========
  const httpsConfig = { cert, key };

  const httpsServer = https.createServer(httpsConfig, app);
  const bridgeServer = http.createServer(bridgeApp);

  httpsServer.listen(PORT, () => {
    console.log(`OpenCode Office Add-in Server running on https://localhost:${PORT}`);
    logInfo('server', 'HTTPS server started', { port: PORT, basePath: BASE_PATH });
  });
  bridgeServer.listen(BRIDGE_PORT, '127.0.0.1', () => {
    console.log(`Office bridge available at http://127.0.0.1:${BRIDGE_PORT}/api`);
    logInfo('server', 'Bridge server started', { port: BRIDGE_PORT });
  });

  const close = () => {
    runtime.close();
    httpsServer.close();
    bridgeServer.close();
    removeBridgeToken(BRIDGE_PORT);
  };

  process.once('SIGINT', () => {
    logInfo('server', 'Received SIGINT');
    close();
    process.exit(0);
  });
  process.once('SIGTERM', () => {
    logInfo('server', 'Received SIGTERM');
    close();
    process.exit(0);
  });

  return { httpsServer, close };
}

// Export for use by tray app
module.exports = { createServer };

// Run directly if not required as a module
if (require.main === module) {
  createServer().catch((error) => {
    console.error(error);
    logError('server', 'Failed to create server', error);
  });
}
