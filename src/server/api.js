const express = require('express');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { validateOfficeToolCall } = require('./officeToolValidation');
const { getLogFilePath, logInfo, logWarn, logError } = require('./devLogger');

const MODEL_FALLBACK = [
  {
    key: 'anthropic/claude-sonnet-4-5',
    label: 'Anthropic / Claude Sonnet 4.5',
    providerID: 'anthropic',
    modelID: 'claude-sonnet-4-5',
  },
];

function hostPrefix(host) {
  if (host === 'powerpoint') return 'PowerPoint: ';
  if (host === 'excel') return 'Excel: ';
  if (host === 'onenote') return 'OneNote: ';
  return 'Word: ';
}

function sessionEventSessionId(event) {
  return event?.properties?.sessionID || event?.properties?.info?.sessionID || event?.properties?.part?.sessionID || null;
}

function isSecureRequest(req) {
  return Boolean(req.secure || req.socket?.encrypted || req.get('x-forwarded-proto') === 'https');
}

function bridgeSessionToken(req) {
  return String(req.get('x-office-bridge-session') || '');
}

function bridgeExecutorId(req) {
  return String(req.get('x-office-executor-id') || req.query.executorId || req.body?.executorId || '');
}

function bridgeAccessToken(req) {
  return String(req.get('x-office-bridge-token') || '');
}

function sendExecuteResponse(res, bridge, body, token) {
  return (async () => {
    const args = body && Object.prototype.hasOwnProperty.call(body, 'args') ? body.args : {};
    validateOfficeToolCall(body.host, body.toolName, args);
    res.json(await bridge.execute(body.host, body.toolName, args, token));
  })();
}

function readRecentLogs(limit = 200) {
  const filePath = getLogFilePath();
  if (!fs.existsSync(filePath)) return [];
  const lines = fs.readFileSync(filePath, 'utf8').split(/\r?\n/).filter(Boolean);
  return lines.slice(-limit);
}

function createApiRouter(runtime, bridge) {
  const apiRouter = express.Router();
  apiRouter.use(express.json({ limit: '50mb' }));
  apiRouter.use((req, res, next) => {
    const startedAt = Date.now();
    const requestId = Math.random().toString(36).slice(2, 8);
    logInfo('http', `${req.method} ${req.originalUrl} started`, {
      requestId,
      host: req.get('host'),
      origin: req.get('origin'),
      referer: req.get('referer'),
      userAgent: req.get('user-agent'),
    });

    res.on('finish', () => {
      const level = res.statusCode >= 500 ? logError : res.statusCode >= 400 ? logWarn : logInfo;
      level('http', `${req.method} ${req.originalUrl} completed`, {
        requestId,
        statusCode: res.statusCode,
        durationMs: Date.now() - startedAt,
      });
    });

    next();
  });

  apiRouter.get('/hello', (req, res) => {
    res.json({ message: 'Hello from backend!', timestamp: new Date().toISOString() });
  });

  apiRouter.post('/upload-image', async (req, res) => {
    try {
      const { dataUrl, name } = req.body;

      if (!dataUrl || !dataUrl.startsWith('data:image/')) {
        return res.status(400).json({ error: 'Invalid image data' });
      }

      const matches = dataUrl.match(/^data:image\/([a-zA-Z+]+);base64,(.+)$/);
      if (!matches || matches.length !== 3) {
        return res.status(400).json({ error: 'Invalid data URL format' });
      }

      const extension = matches[1] === 'svg+xml' ? 'svg' : matches[1];
      const base64Data = matches[2];
      const buffer = Buffer.from(base64Data, 'base64');

      const tempDir = path.join(os.tmpdir(), 'opencode-office-images');
      if (!fs.existsSync(tempDir)) {
        fs.mkdirSync(tempDir, { recursive: true });
      }

      const filename = name || `image-${Date.now()}.${extension}`;
      const filepath = path.join(tempDir, filename);
      fs.writeFileSync(filepath, buffer);

      res.json({ path: filepath, name: filename, mime: `image/${extension === 'svg' ? 'svg+xml' : extension}` });
    } catch (error) {
      console.error('Upload error:', error);
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.get('/fetch', async (req, res) => {
    const url = req.query.url;
    if (!url) {
      return res.status(400).json({ error: 'Missing url parameter' });
    }
    try {
      const https = require('https');
      const http = require('http');
      const parsedUrl = new URL(url);
      const client = parsedUrl.protocol === 'https:' ? https : http;

      const options = {
        hostname: parsedUrl.hostname,
        path: parsedUrl.pathname + parsedUrl.search,
        headers: {
          'User-Agent': 'WordAddinDemo/1.0 (https://github.com; contact@example.com)',
        },
      };

      client.get(options, (response) => {
        let data = '';
        response.on('data', (chunk) => (data += chunk));
        response.on('end', () => {
          res.type('text/plain').send(data);
        });
      }).on('error', (error) => {
        res.status(500).json({ error: error.message });
      });
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.post('/log', (req, res) => {
    const { level = 'error', tag = 'client', message, detail } = req.body || {};
    const scope = String(tag || 'client');
    if (level === 'error') {
      console.error(`[${scope}]`, message, detail || '');
      logError(scope, String(message || 'Client error'), detail);
    } else if (level === 'warn') {
      console.warn(`[${scope}]`, message, detail || '');
      logWarn(scope, String(message || 'Client warning'), detail);
    } else {
      console.log(`[${scope}]`, message, detail || '');
      logInfo(scope, String(message || 'Client log'), detail);
    }
    res.json({ ok: true, logFilePath: getLogFilePath() });
  });

  apiRouter.get('/debug/logs', (req, res) => {
    res.json({ logFilePath: getLogFilePath(), lines: readRecentLogs() });
  });

  apiRouter.get('/models', async (req, res) => {
    try {
      const status = await runtime.status();
      res.json({ models: status.models.length ? status.models : MODEL_FALLBACK });
    } catch {
      res.json({ models: MODEL_FALLBACK });
    }
  });

  apiRouter.get('/browse', (req, res) => {
    try {
      const dir = req.query.path || os.homedir();
      const resolved = path.resolve(String(dir));
      if (!fs.existsSync(resolved) || !fs.statSync(resolved).isDirectory()) {
        return res.status(400).json({ error: 'Not a directory', path: resolved });
      }
      const entries = fs.readdirSync(resolved, { withFileTypes: true });
      const dirs = entries
        .filter((entry) => entry.isDirectory() && !entry.name.startsWith('.'))
        .map((entry) => entry.name)
        .sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
      const parent = path.dirname(resolved);
      res.json({ path: resolved, parent: parent !== resolved ? parent : null, dirs });
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.get('/env', (req, res) => {
    res.json({ cwd: process.cwd(), home: os.homedir() });
  });

  apiRouter.get('/opencode/status', async (req, res) => {
    try {
      res.json(await runtime.status());
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.get('/opencode/sessions', async (req, res) => {
    try {
      const host = String(req.query.host || 'word');
      const shared = String(req.query.shared || '0') === '1';
      const sessions = await runtime.request(`/session?roots=true&limit=100${shared ? '' : `&directory=${encodeURIComponent(runtime.directory())}`}`);
      const prefix = hostPrefix(host);
      const filtered = shared
        ? sessions
        : sessions.filter((item) => item.directory === runtime.directory() && String(item.title || '').startsWith(prefix));
      res.json(filtered);
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.get('/opencode/session/:id', async (req, res) => {
    try {
      res.json(await runtime.request(`/session/${req.params.id}`));
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.get('/opencode/session/:id/children', async (req, res) => {
    try {
      res.json(await runtime.request(`/session/${req.params.id}/children`));
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.patch('/opencode/session/:id', async (req, res) => {
    try {
      res.json(await runtime.request(`/session/${req.params.id}`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: req.body || {},
      }));
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.delete('/opencode/session/:id', async (req, res) => {
    try {
      res.json(await runtime.request(`/session/${req.params.id}`, { method: 'DELETE' }));
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.post('/opencode/session', async (req, res) => {
    try {
      const session = await runtime.request('/session', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: req.body || {},
      });
      res.json(session);
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.post('/opencode/session/:id/message', async (req, res) => {
    try {
      await runtime.request(`/session/${req.params.id}/prompt_async`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: req.body,
      });
      res.status(202).json({ ok: true });
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.post('/opencode/session/:id/abort', async (req, res) => {
    try {
      res.json(await runtime.request(`/session/${req.params.id}/abort`, { method: 'POST' }));
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.get('/opencode/session/:id/messages', async (req, res) => {
    try {
      res.json(await runtime.request(`/session/${req.params.id}/message`));
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.get('/opencode/permissions', async (req, res) => {
    try {
      res.json(await runtime.request('/permission'));
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.post('/opencode/permission/:id/reply', async (req, res) => {
    try {
      res.json(await runtime.request(`/permission/${req.params.id}/reply`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: req.body || {},
      }));
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.get('/opencode/events', async (req, res) => {
    try {
      const sessionId = req.query.sessionId ? String(req.query.sessionId) : null;
      const response = await runtime.request('/event', { raw: true });

      res.setHeader('Content-Type', 'text/event-stream');
      res.setHeader('Cache-Control', 'no-cache');
      res.setHeader('Connection', 'keep-alive');
      res.flushHeaders?.();

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer = '';

      const write = (event) => {
        const id = sessionEventSessionId(event);
        if (sessionId && id && id !== sessionId) return;
        if (sessionId && !id && !['server.connected', 'server.heartbeat'].includes(event.type)) return;
        res.write(`data: ${JSON.stringify(event)}\n\n`);
      };

      req.on('close', async () => {
        try {
          await reader.cancel();
        } catch {}
      });

      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        buffer += decoder.decode(value, { stream: true });
        buffer = buffer.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

        let boundary = buffer.indexOf('\n\n');
        while (boundary >= 0) {
          const raw = buffer.slice(0, boundary);
          buffer = buffer.slice(boundary + 2);
          const data = raw
            .split('\n')
            .filter((line) => line.startsWith('data:'))
            .map((line) => line.slice(5).trim())
            .join('\n');

          if (data) {
            try {
              write(JSON.parse(data));
            } catch {}
          }

          boundary = buffer.indexOf('\n\n');
        }
      }

      res.end();
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.post('/office-tools/register', (req, res) => {
    try {
      const sessionToken = bridgeSessionToken(req);
      const host = String(req.body.host || '');
      res.json(bridge.register(host, sessionToken));
    } catch (error) {
      const message = String(error.message || error);
      res.status(/already registered/.test(message) ? 409 : 401).json({ error: message });
    }
  });

  apiRouter.delete('/office-tools/register/:executorId', (req, res) => {
    try {
      bridge.unregister(req.params.executorId, bridgeSessionToken(req));
      res.json({ ok: true });
    } catch (error) {
      res.status(401).json({ error: error.message });
    }
  });

  apiRouter.get('/office-tools/session', (req, res) => {
    if (!isSecureRequest(req)) {
      return res.status(403).json({ error: 'Office bridge sessions must be requested over HTTPS' });
    }
    res.json({ sessionToken: bridge.issueClientSession() });
  });

  apiRouter.get('/office-tools/poll', (req, res) => {
    try {
      res.json({ request: bridge.poll(bridgeExecutorId(req), bridgeSessionToken(req)) });
    } catch (error) {
      res.status(401).json({ error: error.message });
    }
  });

  apiRouter.post('/office-tools/respond/:id', (req, res) => {
    try {
      bridge.respond(req.params.id, bridgeExecutorId(req), bridgeSessionToken(req), req.body || {});
      res.json({ ok: true });
    } catch (error) {
      res.status(401).json({ error: error.message });
    }
  });

  apiRouter.post('/office-tools/execute', async (req, res) => {
    try {
      await sendExecuteResponse(res, bridge, req.body, bridgeAccessToken(req));
    } catch (error) {
      const message = String(error.message || error);
      const status = /Invalid Office bridge token/.test(message)
        ? 401
        : /Unknown Office tool|not available for host|Missing required|Unexpected args\.|Invalid args/.test(message)
          ? 400
          : 500;
      res.status(status).json({ error: message });
    }
  });

  return apiRouter;
}

function createBridgeRouter(bridge) {
  const bridgeRouter = express.Router();
  bridgeRouter.use(express.json({ limit: '5mb' }));

  bridgeRouter.post('/office-tools/execute', async (req, res) => {
    try {
      await sendExecuteResponse(res, bridge, req.body, bridgeAccessToken(req));
    } catch (error) {
      const message = String(error.message || error);
      const status = /Invalid Office bridge token/.test(message)
        ? 401
        : /Unknown Office tool|not available for host|Missing required|Unexpected args\.|Invalid args/.test(message)
          ? 400
          : 500;
      res.status(status).json({ error: message });
    }
  });

  return bridgeRouter;
}

module.exports = {
  createApiRouter,
  createBridgeRouter,
  MODEL_FALLBACK,
};
