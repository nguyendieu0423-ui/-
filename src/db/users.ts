import { asc, eq, or, sql } from 'drizzle-orm';
import { db, isDatabaseAvailable, markDatabaseUnavailable } from './index.ts';
import { users } from './schema.ts';

function isConnectivityError(error: unknown): boolean {
  const err = error as { code?: string; message?: string; cause?: { code?: string; message?: string } };
  const code = err?.code || err?.cause?.code || '';
  const message = `${err?.message || ''} ${err?.cause?.message || ''}`;
  return (
    code === 'ENOENT' ||
    code === 'ECONNREFUSED' ||
    code === 'ENOTFOUND' ||
    code === 'ETIMEDOUT' ||
    message.includes('ENOENT') ||
    message.includes('Connection terminated unexpectedly')
  );
}

export async function getOrCreateUser(
  uid: string,
  email: string,
  displayName?: string
) {
  const normalizedEmail = email.trim().toLowerCase();
  if (!(await isDatabaseAvailable())) {
    throw new Error('Cloud SQL unavailable', { cause: 'DB_UNAVAILABLE' });
  }

  try {
    const existing = await db
      .select()
      .from(users)
      .where(
        or(
          eq(users.uid, uid),
          sql`lower(${users.email}) = ${normalizedEmail}`
        )
      )
      .limit(1);

    if (existing[0]) {
      const updated = await db
        .update(users)
        .set({
          uid,
          email: normalizedEmail,
          displayName: displayName || existing[0].displayName || null,
        })
        .where(eq(users.id, existing[0].id))
        .returning();
      return updated[0];
    }

    const result = await db
      .insert(users)
      .values({
        uid,
        email: normalizedEmail,
        displayName: displayName || null,
      })
      .returning();

    return result[0];
  } catch (error) {
    if (isConnectivityError(error)) {
      markDatabaseUnavailable();
    } else {
      console.error('Database query failed in getOrCreateUser:', error);
    }
    throw new Error('Failed to sync user account. Please try again later.', {
      cause: error,
    });
  }
}

export async function getPrimaryCrmUser() {
  if (!(await isDatabaseAvailable())) {
    return null;
  }

  try {
    const rows = await db
      .select()
      .from(users)
      .orderBy(asc(users.id))
      .limit(1);
    return rows[0] || null;
  } catch (error) {
    if (isConnectivityError(error)) {
      markDatabaseUnavailable();
    }
    return null;
  }
}
