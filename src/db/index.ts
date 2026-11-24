import fs from 'node:fs';
import path from 'node:path';
import { drizzle } from 'drizzle-orm/node-postgres';
import { Pool } from 'pg';
import * as schema from './schema.ts';

declare global {
  var _postgresPool: Pool | undefined;
  var _dbHealthCache: { available: boolean; checkedAt: number } | undefined;
  var _dbFallbackLogged: boolean | undefined;
}

const DB_HEALTH_TTL_OK_MS = 30_000;
const DB_HEALTH_TTL_FAIL_MS = 60_000;

export function resolveSqlHost(): string | undefined {
  const configured = (process.env.SQL_HOST || '').trim();
  if (!configured) return undefined;

  if (configured.startsWith('/app/cloudsql/')) {
    if (fs.existsSync(path.join(configured, '.s.PGSQL.5432'))) {
      return configured;
    }
    return undefined;
  }

  return configured;
}

function createPoolForHost(host: string | undefined): Pool {
  const pool = new Pool({
    host,
    user: process.env.SQL_USER,
    password: process.env.SQL_PASSWORD,
    database: process.env.SQL_DB_NAME,
    max: 10,
    connectionTimeoutMillis: 5000,
  });

  pool.on('error', () => {
    global._dbHealthCache = {
      available: false,
      checkedAt: Date.now(),
    };
  });

  return pool;
}

export const createPool = () => {
  if (!global._postgresPool) {
    global._postgresPool = createPoolForHost(resolveSqlHost());
  }
  return global._postgresPool;
};

const pool = createPool();

export const db = drizzle(pool, { schema });

export function markDatabaseUnavailable() {
  global._dbHealthCache = {
    available: false,
    checkedAt: Date.now(),
  };
}

export async function isDatabaseAvailable(): Promise<boolean> {
  const now = Date.now();
  const cached = global._dbHealthCache;
  if (cached) {
    const ttl = cached.available ? DB_HEALTH_TTL_OK_MS : DB_HEALTH_TTL_FAIL_MS;
    if (now - cached.checkedAt < ttl) {
      return cached.available;
    }
  }

  const host = resolveSqlHost();
  if (
    !host ||
    !process.env.SQL_USER ||
    !process.env.SQL_PASSWORD ||
    !process.env.SQL_DB_NAME
  ) {
    global._dbHealthCache = { available: false, checkedAt: now };
    if (!global._dbFallbackLogged) {
      global._dbFallbackLogged = true;
      console.log('[DB] Cloud SQL socket not active; running in local fallback mode.');
    }
    return false;
  }

  try {
    await pool.query('SELECT 1');
    global._dbHealthCache = { available: true, checkedAt: Date.now() };
    return true;
  } catch {
    global._dbHealthCache = { available: false, checkedAt: Date.now() };
    if (!global._dbFallbackLogged) {
      global._dbFallbackLogged = true;
      console.log('[DB] Cloud SQL unreachable; running in local fallback mode.');
    }
    return false;
  }
}
