// src/auth.ts
// ---------------------------------------------------------------------------
// Authentication & authorization helpers.
//
// Responsibilities:
//   - Password hashing / verification (bcryptjs, a pure-JS bcrypt implementation
//     so no native build toolchain is required on Alpine / Windows).
//   - Stateless session JWTs (HS256) carrying { role, matricule, name, jti, exp }.
//   - A blacklist backed by SQLite (revoked_tokens) so a logout actually
//     invalidates the token before its natural expiry.
//   - Express middleware: requireAuth (Bearer token gate) and requireRole.
//
// Tokens are opaque to the client: the frontend stores the signed JWT string
// only (never the parsed identity) and lets the server validate it on each call.
// ---------------------------------------------------------------------------
import crypto from "crypto";
import bcrypt from "bcryptjs";
import { isTokenRevoked, UserRole } from "./db";

// Secret used to sign JWTs. In production always set TOKEN_SECRET to a stable,
// high-entropy value (e.g. `openssl rand -hex 32`). If unset, a random secret is
// generated per process start which invalidates sessions on every restart.
const TOKEN_SECRET: string = process.env.TOKEN_SECRET || crypto.randomBytes(32).toString("hex");

// Session lifetime (default 8h). Overridable via SESSION_TTL_MS for testing.
const TOKEN_TTL_MS: number = Number(process.env.SESSION_TTL_MS || 8 * 60 * 60 * 1000);

// bcrypt cost factor. 10 is a reasonable default; raise for more security.
const BCRYPT_ROUNDS: number = Number(process.env.BCRYPT_ROUNDS || 10);

// --- Password hashing ------------------------------------------------------
export function hashPassword(password: string): string {
  return bcrypt.hashSync(password, BCRYPT_ROUNDS);
}

export function comparePassword(password: string, hash: string): boolean {
  try {
    return bcrypt.compareSync(password, hash);
  } catch {
    // Malformed hash -> treat as "no match" (fail closed, never throw on auth path).
    return false;
  }
}

// --- Session JWTs ----------------------------------------------------------
export interface SessionClaims {
  role: UserRole;
  matricule: string;
  name: string;
  jti: string;
  iat: number;
  exp: number;
}

export function signToken(payload: { role: UserRole; matricule: string; name: string }): string {
  const jti = crypto.randomUUID();
  const iat = Date.now();
  const exp = iat + TOKEN_TTL_MS;
  const body = JSON.stringify({ ...payload, jti, iat, exp });
  const header = Buffer.from(JSON.stringify({ alg: "HS256", typ: "JWT" })).toString("base64url");
  const bodyB64 = Buffer.from(body).toString("base64url");
  const sig = crypto.createHmac("sha256", TOKEN_SECRET).update(`${header}.${bodyB64}`).digest("base64url");
  return `${header}.${bodyB64}.${sig}`;
}

export function verifyToken(token: string): SessionClaims | null {
  const parts = token.split(".");
  if (parts.length !== 3) return null;
  const [header, body, sig] = parts;

  // Recompute signature and compare in constant time to thwart timing attacks.
  const expected = crypto.createHmac("sha256", TOKEN_SECRET).update(`${header}.${body}`).digest("base64url");
  const a = Buffer.from(sig);
  const b = Buffer.from(expected);
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;

  let claims: any;
  try {
    claims = JSON.parse(Buffer.from(body, "base64url").toString());
  } catch {
    return null;
  }

  // Enforce expiry.
  if (typeof claims.exp !== "number" || Date.now() > claims.exp) return null;

  // Reject explicitly revoked (logged-out) tokens.
  if (typeof claims.jti === "string" && isTokenRevoked(claims.jti)) return null;

  return claims as SessionClaims;
}

// --- Middleware ------------------------------------------------------------
// requireAuth is mounted under "/api/" (app.use("/api/", requireAuth)), so Express
// strips the "/api/" prefix from req.path. The public login route must be reachable
// without a token in every environment, therefore we match it against BOTH the
// mount-stripped path ("/auth/login") and the full original URL.
export function isPublicRoute(req: any): boolean {
  // requireAuth is mounted under "/api/" (app.use("/api/", requireAuth)), so Express
  // strips the "/api/" prefix from req.path (e.g. "/auth/login"). We match public routes
  // against both the mount-stripped path and the full original URL so the exemption works
  // regardless of how the middleware is mounted.
  const stripped = (req.path || "").replace(/\/+$/, "");
  const full = (req.originalUrl || req.url || "").split("?")[0].replace(/\/+$/, "");
  const isLogin =
    stripped === "/auth/login" ||
    stripped === "/api/auth/login" ||
    full === "/api/auth/login" ||
    full.endsWith("/api/auth/login");
  // Technician list is intentionally public (matricule + name only, never the hash) so the
  // login screen can render the account selector before any authentication.
  const isTechList =
    req.method === "GET" &&
    (stripped === "/technicians" ||
      stripped === "/api/technicians" ||
      full === "/api/technicians" ||
      full.endsWith("/api/technicians"));
  return isLogin || isTechList;
}

export function requireAuth(req: any, res: any, next: any): void {
  if (req.method === "OPTIONS") {
    next();
    return;
  }
  if (isPublicRoute(req)) {
    next();
    return;
  }

  // Local development relaxation (no token required when NODE_ENV is unset/dev).
  const nodeEnv = (process.env.NODE_ENV || "").trim();
  if (nodeEnv === "development" || nodeEnv === "dev" || nodeEnv === "") {
    next();
    return;
  }

  const auth = req.headers["authorization"] || "";
  const m = (auth as string).match(/^Bearer\s+(.+)$/i);
  if (!m) {
    res.status(401).json({ error: "Non autorisé: jeton manquant." });
    return;
  }

  // Static API token (backward compatible) OR a valid session JWT.
  const token = m[1];
  const API_TOKEN = process.env.API_TOKEN;
  if (API_TOKEN && token === API_TOKEN) {
    next();
    return;
  }
  const claims = verifyToken(token);
  if (claims && (claims.role === "admin" || claims.role === "technician")) {
    req.auth = claims;
    next();
    return;
  }
  res.status(401).json({ error: "Non autorisé: jeton invalide ou expiré." });
}

export function requireRole(role: "admin" | "technician") {
  return (req: any, res: any, next: any): void => {
    const auth = req.headers["authorization"] || "";
    const m = (auth as string).match(/^Bearer\s+(.+)$/i);
    if (!m) {
      res.status(401).json({ error: "Non autorisé: jeton manquant." });
      return;
    }
    const claims = verifyToken(m[1]);
    if (!claims || claims.role !== role) {
      res.status(403).json({ error: "Accès refusé: privilège insuffisant." });
      return;
    }
    req.auth = claims;
    next();
  };
}
