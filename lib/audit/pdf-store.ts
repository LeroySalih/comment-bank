import { randomUUID } from 'crypto';

type Entry = { buffer: Buffer; expiresAt: number };

// Module-level Map — persists for the lifetime of the Node.js process.
// Both routes must use `export const runtime = 'nodejs'` to share this instance.
const store = new Map<string, Entry>();

const TTL_MS = 5 * 60 * 1000; // 5 minutes

/** Store a PDF buffer and return a short-lived token to retrieve it. */
export function storePdf(buffer: Buffer): string {
  const token = randomUUID();
  store.set(token, { buffer, expiresAt: Date.now() + TTL_MS });
  // Lazy cleanup: remove expired entries on each store
  for (const [key, entry] of store) {
    if (entry.expiresAt < Date.now()) store.delete(key);
  }
  return token;
}

/** Retrieve and delete a PDF buffer by token. Returns null if expired or not found. */
export function consumePdf(token: string): Buffer | null {
  const entry = store.get(token);
  if (!entry) return null;
  store.delete(token);
  if (entry.expiresAt < Date.now()) return null;
  return entry.buffer;
}
