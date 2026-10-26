import express from 'express';
import type { Request, Response } from 'express';
import { createServer } from 'http';
import { WebSocketServer, WebSocket } from 'ws';
import fs from 'fs';
import fsPromises from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = 3000;
const DEFAULT_TEAMS_COUNT = 14;
const DATA_DIR = path.resolve(__dirname, 'data');
const DATA_FILE = path.resolve(DATA_DIR, 'teams.json');

interface TeamRecord {
  id: number;
  teamNumber: string;
  countryName: string;
  sentence: string;
  updatedAt: string | null;
}

interface PersistedStore {
  teams: TeamRecord[];
  appsScriptUrl?: string;
  updatedAt: string;
}

function generateDefaultTeams(): TeamRecord[] {
  const list: TeamRecord[] = [];
  for (let i = 1; i <= DEFAULT_TEAMS_COUNT; i++) {
    list.push({
      id: i,
      teamNumber: `${i}모둠`,
      countryName: '',
      sentence: '',
      updatedAt: null,
    });
  }
  return list;
}

function sanitizeString(val: unknown, maxLen: number): string {
  if (typeof val !== 'string') return '';
  return val.trim().slice(0, maxLen);
}

function normalizeTeams(raw: unknown): TeamRecord[] {
  const defaults = generateDefaultTeams();
  if (!Array.isArray(raw)) return defaults;

  return defaults.map((def, idx) => {
    const found = raw.find(
      (item: unknown) =>
        item &&
        typeof item === 'object' &&
        Number((item as Record<string, unknown>).id) === def.id
    ) ?? raw[idx];

    if (!found || typeof found !== 'object') return def;
    const obj = found as Record<string, unknown>;
    return {
      id: def.id,
      teamNumber: def.teamNumber,
      countryName: sanitizeString(obj.countryName, 100),
      sentence: sanitizeString(obj.sentence, 2000),
      updatedAt: typeof obj.updatedAt === 'string' ? obj.updatedAt : null,
    };
  });
}

let store: PersistedStore = {
  teams: generateDefaultTeams(),
  appsScriptUrl: '',
  updatedAt: new Date().toISOString(),
};

async function loadStoreFromDisk(): Promise<void> {
  try {
    await fsPromises.mkdir(DATA_DIR, { recursive: true });
    if (fs.existsSync(DATA_FILE)) {
      const content = await fsPromises.readFile(DATA_FILE, 'utf-8');
      const parsed = JSON.parse(content) as Partial<PersistedStore>;
      store = {
        teams: normalizeTeams(parsed.teams),
        appsScriptUrl: sanitizeString(parsed.appsScriptUrl, 500),
        updatedAt:
          typeof parsed.updatedAt === 'string'
            ? parsed.updatedAt
            : new Date().toISOString(),
      };
    } else {
      await saveStoreToDisk();
    }
  } catch (err) {
    console.error('Failed to load teams store from disk, using defaults:', err);
  }
}

let writeQueue: Promise<void> = Promise.resolve();

function saveStoreToDisk(): Promise<void> {
  store.updatedAt = new Date().toISOString();
  const payload = JSON.stringify(store, null, 2);
  writeQueue = writeQueue
    .then(async () => {
      await fsPromises.mkdir(DATA_DIR, { recursive: true });
      const tempFile = `${DATA_FILE}.tmp`;
      await fsPromises.writeFile(tempFile, payload, 'utf-8');
      await fsPromises.rename(tempFile, DATA_FILE);
    })
    .catch((err) => {
      console.error('Failed to persist teams store:', err);
    });
  return writeQueue;
}

async function startServer() {
  await loadStoreFromDisk();

  const app = express();
  app.use(express.json({ limit: '256kb' }));
  app.use('/api', (req: Request, res: Response, next) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, PUT, POST, DELETE, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate');
    if (req.method === 'OPTIONS') {
      res.status(204).end();
      return;
    }
    next();
  });

  const httpServer = createServer(app);
  const wss = new WebSocketServer({ server: httpServer, path: '/ws' });

  function broadcast(event: Record<string, unknown>) {
    const msg = JSON.stringify(event);
    for (const client of wss.clients) {
      if (client.readyState === WebSocket.OPEN) {
        client.send(msg);
      }
    }
  }

  wss.on('connection', (ws: WebSocket) => {
    // Send authoritative state immediately upon connection
    ws.send(
      JSON.stringify({
        type: 'init',
        teams: store.teams,
        appsScriptUrl: store.appsScriptUrl || '',
        updatedAt: store.updatedAt,
      })
    );

    ws.on('message', (raw) => {
      try {
        const msg = JSON.parse(String(raw)) as { type?: string };
        if (msg.type === 'ping') {
          ws.send(JSON.stringify({ type: 'pong' }));
        }
      } catch {
        // Ignore malformed client frames
      }
    });
  });

  // Keep WebSocket connections warm behind proxies
  const heartbeatInterval = setInterval(() => {
    for (const client of wss.clients) {
      if (client.readyState === WebSocket.OPEN) {
        client.ping();
      }
    }
  }, 25000);

  wss.on('close', () => {
    clearInterval(heartbeatInterval);
  });

  // =========================================================================
  // REST API Routes (/api/*)
  // =========================================================================

  // Get full shared classroom state
  app.get('/api/teams', (_req: Request, res: Response) => {
    res.setHeader('Cache-Control', 'no-store');
    res.json({
      teams: store.teams,
      appsScriptUrl: store.appsScriptUrl || '',
      updatedAt: store.updatedAt,
    });
  });

  // Update a single team (atomic per-team update so concurrent groups never overwrite each other)
  app.put('/api/teams/:id', async (req: Request, res: Response) => {
    const teamId = Number(req.params.id);
    if (!Number.isInteger(teamId) || teamId < 1 || teamId > DEFAULT_TEAMS_COUNT) {
      res.status(400).json({ error: '유효하지 않은 모둠 번호입니다.' });
      return;
    }

    const team = store.teams.find((t) => t.id === teamId);
    if (!team) {
      res.status(404).json({ error: '모둠을 찾을 수 없습니다.' });
      return;
    }

    const body = (req.body ?? {}) as Record<string, unknown>;
    if ('countryName' in body) {
      team.countryName = sanitizeString(body.countryName, 100);
    }
    if ('sentence' in body) {
      team.sentence = sanitizeString(body.sentence, 2000);
    }
    team.updatedAt = new Date().toISOString();

    await saveStoreToDisk();

    broadcast({
      type: 'team:updated',
      team,
      updatedAt: store.updatedAt,
    });

    res.json({ ok: true, team, updatedAt: store.updatedAt });
  });

  // Clear a single team's content
  app.delete('/api/teams/:id', async (req: Request, res: Response) => {
    const teamId = Number(req.params.id);
    if (!Number.isInteger(teamId) || teamId < 1 || teamId > DEFAULT_TEAMS_COUNT) {
      res.status(400).json({ error: '유효하지 않은 모둠 번호입니다.' });
      return;
    }

    const team = store.teams.find((t) => t.id === teamId);
    if (!team) {
      res.status(404).json({ error: '모둠을 찾을 수 없습니다.' });
      return;
    }

    team.countryName = '';
    team.sentence = '';
    team.updatedAt = null;

    await saveStoreToDisk();

    broadcast({
      type: 'team:updated',
      team,
      updatedAt: store.updatedAt,
    });

    res.json({ ok: true, team, updatedAt: store.updatedAt });
  });

  // Batch operations (batch country names, sample data, sheet import, reset all)
  app.put('/api/teams', async (req: Request, res: Response) => {
    const body = (req.body ?? {}) as {
      mode?: 'batch-countries' | 'replace-all' | 'reset-all';
      teams?: unknown;
      countries?: Record<string, string>;
      appsScriptUrl?: string;
    };

    if (typeof body.appsScriptUrl === 'string') {
      store.appsScriptUrl = sanitizeString(body.appsScriptUrl, 500);
    }

    if (body.mode === 'reset-all') {
      store.teams = generateDefaultTeams();
    } else if (body.mode === 'batch-countries' && body.countries && typeof body.countries === 'object') {
      // Only update country names without erasing sentences students already wrote!
      for (const team of store.teams) {
        const key = String(team.id);
        if (key in body.countries) {
          team.countryName = sanitizeString(body.countries[key], 100);
          team.updatedAt = new Date().toISOString();
        }
      }
    } else if (Array.isArray(body.teams)) {
      store.teams = normalizeTeams(body.teams);
    }

    await saveStoreToDisk();

    broadcast({
      type: 'teams:synced',
      teams: store.teams,
      appsScriptUrl: store.appsScriptUrl || '',
      updatedAt: store.updatedAt,
    });

    res.json({
      ok: true,
      teams: store.teams,
      appsScriptUrl: store.appsScriptUrl || '',
      updatedAt: store.updatedAt,
    });
  });

  // Save shared Google Apps Script URL
  app.put('/api/config', async (req: Request, res: Response) => {
    const body = (req.body ?? {}) as { appsScriptUrl?: unknown };
    store.appsScriptUrl = sanitizeString(body.appsScriptUrl, 500);
    await saveStoreToDisk();

    broadcast({
      type: 'config:updated',
      appsScriptUrl: store.appsScriptUrl,
    });

    res.json({ ok: true, appsScriptUrl: store.appsScriptUrl });
  });

  // =========================================================================
  // Frontend Static / Vite Middleware
  // =========================================================================
  const distPath = path.resolve(__dirname, 'dist');
  const isDevLifecycle = process.env.npm_lifecycle_event === 'dev';
  const isProd =
    !isDevLifecycle &&
    (process.env.NODE_ENV === 'production' || fs.existsSync(distPath));

  if (!isProd) {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(distPath));
    // Fallback for root script.js / style.css if referenced directly
    app.get('/script.js', (_req: Request, res: Response) => {
      const distScript = path.resolve(distPath, 'script.js');
      const rootScript = path.resolve(__dirname, 'script.js');
      res.sendFile(fs.existsSync(distScript) ? distScript : rootScript);
    });
    app.get('/style.css', (_req: Request, res: Response) => {
      const distStyle = path.resolve(distPath, 'style.css');
      const rootStyle = path.resolve(__dirname, 'style.css');
      res.sendFile(fs.existsSync(distStyle) ? distStyle : rootStyle);
    });
    app.get('*', (_req: Request, res: Response) => {
      res.sendFile(path.resolve(distPath, 'index.html'));
    });
  }

  httpServer.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
