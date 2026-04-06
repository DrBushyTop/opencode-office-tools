const path = require('path');
const fs = require('fs');
const express = require('express');
const { createHttpRuntime, registerShutdown } = require('./server/httpRuntime');
const { logInfo, logError } = require('./server/devLogger');

const isPkg = typeof process.pkg !== 'undefined';

function getBasePath() {
  if (typeof process.env.OPENCODE_OFFICE_BASE_PATH === 'string' && process.env.OPENCODE_OFFICE_BASE_PATH.trim()) {
    return process.env.OPENCODE_OFFICE_BASE_PATH;
  }
  if (isPkg) {
    return path.dirname(process.execPath);
  }
  return path.resolve(__dirname, '..');
}

function parsePort(value, fallback) {
  const port = Number.parseInt(String(value ?? fallback), 10);
  if (!Number.isFinite(port) || port <= 0) {
    throw new Error(`Invalid port: ${value}`);
  }
  return port;
}

function readCertificates(basePath) {
  const certPath = path.join(basePath, 'certs', 'localhost.pem');
  const keyPath = path.join(basePath, 'certs', 'localhost-key.pem');

  if (!fs.existsSync(certPath) || !fs.existsSync(keyPath)) {
    console.error('SSL certificates not found!');
    console.error('Expected:', certPath);
    console.error('Expected:', keyPath);
    logError('server', 'SSL certificates not found', { certPath, keyPath });
    process.exit(1);
  }

  return {
    cert: fs.readFileSync(certPath),
    key: fs.readFileSync(keyPath),
  };
}

function attachBuiltFrontend(app, basePath) {
  const distPath = path.join(basePath, 'dist');

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

  app.use(express.static(distPath));
  app.get('*path', (req, res) => {
    logInfo('static', 'Serving SPA fallback', { url: req.originalUrl, distPath });
    res.sendFile(path.join(distPath, 'index.html'));
  });
}

async function createServer() {
  const basePath = getBasePath();
  const port = parsePort(process.env.PORT, 52390);
  const bridgePort = parsePort(process.env.BRIDGE_PORT, 52391);
  const runtime = createHttpRuntime({
    port,
    bridgePort,
    ...readCertificates(basePath),
  });

  try {
    attachBuiltFrontend(runtime.app, basePath);
    registerShutdown(runtime.close);
    await runtime.listen({
      onHttpsListening: () => {
        console.log(`OpenCode Office Add-in Server running on https://localhost:${port}`);
        logInfo('server', 'HTTPS server started', { port, basePath });
      },
      onBridgeListening: () => {
        console.log(`Office bridge available at http://127.0.0.1:${bridgePort}/api`);
        logInfo('server', 'Bridge server started', { port: bridgePort });
      },
    });

    return { httpsServer: runtime.httpsServer, close: runtime.close };
  } catch (error) {
    await runtime.close();
    throw error;
  }
}

module.exports = { createServer };

if (require.main === module) {
  createServer().catch((error) => {
    console.error(error);
    logError('server', 'Failed to create server', error);
  });
}
