const express = require('express');
const path = require('path');
const fs = require('fs');
const os = require('os');

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
  return 'Word: ';
}

function sessionEventSessionId(event) {
  return event?.properties?.sessionID || event?.properties?.info?.sessionID || event?.properties?.part?.sessionID || null;
}

function createApiRouter(runtime, bridge) {
  const apiRouter = express.Router();
  apiRouter.use(express.json({ limit: '50mb' }));

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
    const prefix = `[${tag}]`;
    if (level === 'error') console.error(prefix, message, detail || '');
    else console.log(prefix, message, detail || '');
    res.sendStatus(204);
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
    bridge.register(req.body.host);
    res.json({ ok: true });
  });

  apiRouter.delete('/office-tools/register/:host', (req, res) => {
    bridge.unregister(req.params.host);
    res.json({ ok: true });
  });

  apiRouter.get('/office-tools/poll', (req, res) => {
    const host = String(req.query.host || '');
    try {
      res.json({ request: bridge.poll(host) });
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  apiRouter.post('/office-tools/respond/:id', (req, res) => {
    try {
      bridge.respond(req.params.id, req.body || {});
      res.json({ ok: true });
    } catch (error) {
      res.status(404).json({ error: error.message });
    }
  });

  apiRouter.post('/office-tools/execute', async (req, res) => {
    try {
      res.json(await bridge.execute(req.body.host, req.body.toolName, req.body.args));
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
  });

  return apiRouter;
}

module.exports = {
  createApiRouter,
  MODEL_FALLBACK,
};
