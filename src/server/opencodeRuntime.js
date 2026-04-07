const path = require('path');
const net = require('net');
const os = require('os');
const fs = require('fs');
const { z } = require('zod');

const DEFAULT_HOST = '127.0.0.1';
const DEFAULT_PORT = 4096;
const OFFICE_ROOT = path.resolve(__dirname, '..', '..');

const runtimeModelSchema = z.object({
  id: z.string(),
  name: z.string().optional(),
  limit: z.object({
    context: z.number().optional(),
  }).passthrough().optional(),
  variants: z.record(z.string(), z.record(z.string(), z.any())).optional(),
}).passthrough();

const runtimeProviderSchema = z.object({
  id: z.string(),
  name: z.string().optional(),
  models: z.record(z.string(), runtimeModelSchema).default({}),
}).passthrough();

const requestOptionsSchema = z.object({
  baseUrl: z.string().optional(),
  directory: z.string().optional(),
  method: z.string().optional(),
  headers: z.record(z.string(), z.string()).optional(),
  body: z.unknown().optional(),
  raw: z.boolean().optional(),
}).passthrough();

function officeConfigDirectory() {
  return process.env.OPENCODE_OFFICE_CONFIG_DIR
    ? path.resolve(process.env.OPENCODE_OFFICE_CONFIG_DIR)
    : path.join(officeDirectory(), '.opencode');
}

function trimSlash(value) {
  return String(value || '').replace(/\/+$/, '');
}

function officeDirectory() {
  return process.env.OPENCODE_OFFICE_DIRECTORY
    ? path.resolve(process.env.OPENCODE_OFFICE_DIRECTORY)
    : OFFICE_ROOT;
}

function configuredBaseUrl() {
  const value = process.env.OPENCODE_OFFICE_RUNTIME_URL || process.env.OPENCODE_RUNTIME_URL;
  return value ? trimSlash(value) : '';
}

function defaultOpencodeSearchPaths() {
  const home = os.homedir();
  const candidates = process.platform === 'win32'
    ? [
      path.join(home, '.opencode', 'bin'),
      path.join(home, '.local', 'bin'),
      path.join(home, 'AppData', 'Local', 'Programs', 'opencode', 'bin'),
    ]
    : [
      path.join(home, '.opencode', 'bin'),
      path.join(home, '.local', 'bin'),
      path.join(home, '.bun', 'bin'),
      '/opt/homebrew/bin',
      '/usr/local/bin',
      '/usr/bin',
      '/bin',
    ];

  return candidates.filter((candidate, index, all) => all.indexOf(candidate) === index);
}

function binaryPathExists(directory) {
  const binaryName = process.platform === 'win32' ? 'opencode.exe' : 'opencode';
  return fs.existsSync(path.join(directory, binaryName));
}

function runtimePathEnv(currentPath = process.env.PATH || '') {
  const existing = String(currentPath || '').split(path.delimiter).filter(Boolean);
  const additions = defaultOpencodeSearchPaths().filter((directory) => binaryPathExists(directory) && !existing.includes(directory));
  return [...additions, ...existing].join(path.delimiter);
}

function toModelInfo(provider, model) {
  const value = `${provider.id}/${model.id}`;
  const variants = model.variants && typeof model.variants === 'object'
    ? Object.keys(model.variants).filter((key) => !model.variants[key]?.disabled)
    : undefined;
  return {
    key: value,
    label: `${provider.name || provider.id} / ${model.name || model.id}`,
    providerID: provider.id,
    modelID: model.id,
    limitContext: model.limit?.context,
    variants: variants && variants.length > 0 ? variants : undefined,
  };
}

function configuredModels(configProviders) {
  const items = [];
  const seen = new Set();

  const providerItems = Array.isArray(configProviders?.providers) ? configProviders.providers : [];

  for (const providerItem of providerItems) {
    const parsedProvider = runtimeProviderSchema.safeParse(providerItem);
    if (!parsedProvider.success) continue;
    const provider = parsedProvider.data;

    for (const model of Object.values(provider.models || {})) {
      const modelInfo = toModelInfo(provider, model);
      if (seen.has(modelInfo.key)) continue;
      seen.add(modelInfo.key);
      items.push(modelInfo);
    }
  }

  return items;
}

async function readResponseBody(response) {
  if (response.status === 204 || response.status === 205 || response.status === 304) {
    return null;
  }

  const text = await response.text();
  if (!text) return null;

  const contentType = response.headers.get('content-type') || '';
  if (contentType.includes('application/json')) {
    return JSON.parse(text);
  }

  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

async function sdk() {
  return import('@opencode-ai/sdk');
}

function isReachablePort(port) {
  return new Promise((resolve) => {
    const server = net.createServer();
    server.once('error', () => resolve(false));
    server.listen(port, DEFAULT_HOST, () => {
      server.close(() => resolve(true));
    });
  });
}

async function getFreePort(start = DEFAULT_PORT) {
  for (let port = start; port < start + 50; port += 1) {
    if (await isReachablePort(port)) {
      return port;
    }
  }
  throw new Error('Could not find an available OpenCode port');
}

class OpencodeRuntime {
  constructor() {
    this.runtime = null;
    this.starting = null;
    this.cleanup = null;
  }

  directory(value) {
    return value ? path.resolve(value) : officeDirectory();
  }

  headers(extra = {}, directory) {
    return {
      ...extra,
      'x-opencode-directory': encodeURIComponent(this.directory(directory)),
    };
  }

  async ensureRuntime() {
    if (this.runtime) return this.runtime;
    if (this.starting) return this.starting;

    this.starting = (async () => {
      const attached = await this.tryAttach();
      if (attached) {
        this.runtime = attached;
        return attached;
      }

      const spawned = await this.spawn();
      this.runtime = spawned;
      return spawned;
    })();

    try {
      return await this.starting;
    } finally {
      this.starting = null;
    }
  }

  async tryAttach() {
    const baseUrl = configuredBaseUrl();
    if (!baseUrl) return null;

    try {
      await this.request('/provider', { baseUrl });
      return { mode: 'attached', baseUrl, child: null };
    } catch {
      return null;
    }
  }

  async spawn() {
    const port = process.env.OPENCODE_OFFICE_PORT
      ? Number(process.env.OPENCODE_OFFICE_PORT)
      : await getFreePort(DEFAULT_PORT);
    const dir = officeConfigDirectory();
    const env = process.env.OPENCODE_CONFIG_DIR;
    const previousPath = process.env.PATH;
    process.env.OPENCODE_CONFIG_DIR = dir;
    process.env.PATH = runtimePathEnv(previousPath);

    try {
      const mod = await sdk();
      const server = await mod.createOpencodeServer({
        hostname: DEFAULT_HOST,
        port,
        timeout: 15000,
      });
      const baseUrl = server.url.replace(/\/+$/, '');
      this.cleanup = () => server.close();
      return { mode: 'spawned', baseUrl, server };
    } finally {
      if (typeof previousPath === 'string') process.env.PATH = previousPath;
      else delete process.env.PATH;
      if (env) process.env.OPENCODE_CONFIG_DIR = env;
      else delete process.env.OPENCODE_CONFIG_DIR;
    }
  }

  close() {
    this.cleanup?.();
    this.cleanup = null;
  }

  async request(url, options = {}) {
    const parsedOptions = requestOptionsSchema.parse(options || {});
    const runtime = parsedOptions.baseUrl ? { baseUrl: parsedOptions.baseUrl } : await this.ensureRuntime();
    const response = await fetch(`${runtime.baseUrl}${url}`, {
      method: parsedOptions.method || 'GET',
      headers: this.headers(parsedOptions.headers, parsedOptions.directory),
      body: parsedOptions.body ? JSON.stringify(parsedOptions.body) : undefined,
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(text || `OpenCode request failed: ${response.status}`);
    }

    if (parsedOptions.raw) {
      return response;
    }

    return readResponseBody(response);
  }

  async listModels() {
    const configProviders = await this.request('/config/providers');
    const narrowed = configuredModels(configProviders);
    if (narrowed.length > 0) return narrowed;

    const providers = await this.request('/provider');
    const providerItems = Array.isArray(providers?.all) ? providers.all : [];
    const items = [];

    for (const providerItem of providerItems) {
      const parsedProvider = runtimeProviderSchema.safeParse(providerItem);
      if (!parsedProvider.success) continue;
      const provider = parsedProvider.data;

      for (const model of Object.values(provider.models || {})) {
        items.push(toModelInfo(provider, model));
      }
    }

    return items;
  }

  async status(directory) {
    const runtime = await this.ensureRuntime();
    let models = [];

    try {
      models = await this.listModels();
    } catch (error) {
      console.warn('Failed to list OpenCode models:', error.message);
    }

    return {
      mode: runtime.mode,
      baseUrl: runtime.baseUrl,
      directory: this.directory(directory),
      configDirectory: officeConfigDirectory(),
      models,
    };
  }
}

module.exports = {
  OpencodeRuntime,
  configuredModels,
  officeDirectory,
  officeConfigDirectory,
  readResponseBody,
  runtimePathEnv,
};
