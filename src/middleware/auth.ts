import type { NextFunction, Request, Response } from 'express';
import type { DecodedIdToken } from 'firebase-admin/auth';
import { adminAuth } from '../lib/firebase-admin.ts';
import { isDatabaseAvailable } from '../db/index.ts';
import { getOrCreateUser } from '../db/users.ts';

export interface AuthRequest extends Request {
  user?: DecodedIdToken;
  dbAvailable?: boolean;
  dbUser?: {
    id: number;
    uid: string;
    email: string;
    displayName: string | null;
  };
}

export const requireAuth = async (
  req: AuthRequest,
  res: Response,
  next: NextFunction
) => {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    res.status(401).json({ error: 'Unauthorized: Missing token' });
    return;
  }

  const token = authHeader.split('Bearer ')[1];
  let decodedToken: DecodedIdToken;
  try {
    decodedToken = await adminAuth.verifyIdToken(token);
  } catch (error) {
    console.error('Error verifying Firebase ID token:', error);
    res.status(401).json({ error: 'Unauthorized: Invalid token' });
    return;
  }

  req.user = decodedToken;
  const normalizedEmail = (
    decodedToken.email || `${decodedToken.uid}@user.local`
  )
    .trim()
    .toLowerCase();
  const displayName =
    typeof decodedToken.name === 'string' ? decodedToken.name : null;

  const dbReady = await isDatabaseAvailable();
  if (!dbReady) {
    req.dbUser = {
      id: 0,
      uid: decodedToken.uid,
      email: normalizedEmail,
      displayName,
    };
    req.dbAvailable = false;
    next();
    return;
  }

  try {
    const dbUser = await getOrCreateUser(
      decodedToken.uid,
      normalizedEmail,
      displayName || undefined
    );
    req.dbUser = dbUser;
    req.dbAvailable = true;
    next();
  } catch {
    req.dbUser = {
      id: 0,
      uid: decodedToken.uid,
      email: normalizedEmail,
      displayName,
    };
    req.dbAvailable = false;
    next();
  }
};

export const optionalAuth = async (
  req: AuthRequest,
  _res: Response,
  next: NextFunction
) => {
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith('Bearer ')) {
    req.dbAvailable = false;
    next();
    return;
  }

  const token = authHeader.split('Bearer ')[1];
  try {
    const decodedToken = await adminAuth.verifyIdToken(token);
    req.user = decodedToken;
    const normalizedEmail = (
      decodedToken.email || `${decodedToken.uid}@user.local`
    )
      .trim()
      .toLowerCase();
    const displayName =
      typeof decodedToken.name === 'string' ? decodedToken.name : null;

    if (!(await isDatabaseAvailable())) {
      req.dbUser = {
        id: 0,
        uid: decodedToken.uid,
        email: normalizedEmail,
        displayName,
      };
      req.dbAvailable = false;
      next();
      return;
    }

    try {
      const dbUser = await getOrCreateUser(
        decodedToken.uid,
        normalizedEmail,
        displayName || undefined
      );
      req.dbUser = dbUser;
      req.dbAvailable = true;
    } catch {
      req.dbUser = {
        id: 0,
        uid: decodedToken.uid,
        email: normalizedEmail,
        displayName,
      };
      req.dbAvailable = false;
    }
  } catch {
    req.dbAvailable = false;
  }
  next();
};
