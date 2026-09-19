// src/db.ts
// ---------------------------------------------------------------------------
// SQLite persistence layer for the Master Management System.
//
// Holds the authoritative `users` table (technicians + administrators) and a
// `revoked_tokens` table used to invalidate sessions on logout. Passwords are
// NEVER stored in clear text: only a bcrypt hash is kept in `password_hash`.
//
// The SQLite file lives in <cwd>/data/database.sqlite (overridable via DATA_DIR /
// DB_PATH env vars). In Docker this maps to /app/data, which is persisted by a
// bind-mounted volume declared in docker-compose.yml so user accounts survive
// container restarts.
// ---------------------------------------------------------------------------
import Database from "better-sqlite3";
import path from "path";
import fs from "fs";

// Resolve the data directory relative to the process working directory. In the
// Docker image WORKDIR is /app and the compose volume is ./data:/app/data, so
// this resolves to the persisted /app/data folder. Env override allows tests /
// alternate layouts.
export const DATA_DIR: string = process.env.DATA_DIR || path.join(process.cwd(), "data");
export const DB_PATH: string = process.env.DB_PATH || path.join(DATA_DIR, "database.sqlite");

// Guarantee the data directory exists before opening the database file.
if (!fs.existsSync(DATA_DIR)) {
  try {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  } catch (err) {
    console.warn("[db] Impossible de créer le répertoire de données:", DATA_DIR, err);
  }
}

// Open (creating the file if missing) and configure for safe concurrent access.
export const db: Database.Database = new Database(DB_PATH);
db.pragma("journal_mode = WAL");
db.pragma("foreign_keys = ON");

// --- Schema -----------------------------------------------------------------
// `users` is the single source of truth for authentication & authorization.
// `role` is constrained to the two allowed values at the database level.
db.exec(`
CREATE TABLE IF NOT EXISTS users (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  matricule     TEXT    NOT NULL UNIQUE,
  nom           TEXT    NOT NULL,
  role          TEXT    NOT NULL CHECK (role IN ('technician', 'admin')),
  password_hash TEXT    NOT NULL,
  created_at    TEXT    NOT NULL,
  updated_at    TEXT    NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_users_role ON users(role);

-- Session tokens are stateless JWTs identified by a unique jti. To be able to
-- invalidate a session on logout we store its id and expiry here.
CREATE TABLE IF NOT EXISTS revoked_tokens (
  jti  TEXT    PRIMARY KEY,
  exp  INTEGER NOT NULL
);
`);

export type UserRole = "technician" | "admin";

export interface UserRow {
  id: number;
  matricule: string;
  nom: string;
  role: UserRole;
  password_hash: string;
  created_at: string;
  updated_at: string;
}

const nowIso = (): string => new Date().toISOString();

// --- User queries ----------------------------------------------------------
export function countUsers(): number {
  return (db.prepare("SELECT COUNT(*) AS n FROM users").get() as { n: number }).n;
}

export function getUserByMatricule(matricule: string): UserRow | null {
  const row = db.prepare("SELECT * FROM users WHERE matricule = ?").get(matricule.trim().toUpperCase()) as UserRow | undefined;
  return row ?? null;
}

export function createUser(input: { matricule: string; nom: string; role: UserRole; passwordHash: string }): UserRow {
  const now = nowIso();
  db.prepare(
    "INSERT INTO users (matricule, nom, role, password_hash, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?)"
  ).run(input.matricule.trim().toUpperCase(), input.nom.trim(), input.role, input.passwordHash, now, now);
  return getUserByMatricule(input.matricule)!;
}

export function listTechnicians(): { matricule: string; name: string }[] {
  const rows = db.prepare("SELECT matricule, nom FROM users WHERE role = 'technician' ORDER BY nom ASC").all() as { matricule: string; nom: string }[];
  return rows.map((r) => ({ matricule: r.matricule, name: r.nom }));
}

export function deleteTechnician(matricule: string): boolean {
  const info = db.prepare("DELETE FROM users WHERE matricule = ? AND role = 'technician'").run(matricule.trim().toUpperCase());
  return info.changes > 0;
}

// --- Token revocation ------------------------------------------------------
export function revokeToken(jti: string, exp: number): void {
  db.prepare("INSERT OR IGNORE INTO revoked_tokens (jti, exp) VALUES (?, ?)").run(jti, Math.floor(exp));
}

export function isTokenRevoked(jti: string): boolean {
  const row = db.prepare("SELECT jti FROM revoked_tokens WHERE jti = ?").get(jti) as { jti: string } | undefined;
  return !!row;
}

export function purgeExpiredRevocations(): void {
  db.prepare("DELETE FROM revoked_tokens WHERE exp < ?").run(Math.floor(Date.now()));
}
