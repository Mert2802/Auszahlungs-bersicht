import express from "express";
import cors from "cors";
import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";
import sqlite3 from "sqlite3";
import { open } from "sqlite";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = process.env.PORT || 3001;
const JWT_SECRET = process.env.JWT_SECRET || "change-me";
const ADMIN_USER = process.env.ADMIN_USER || "admin";
const ADMIN_PASS = process.env.ADMIN_PASS || "admin123";

const app = express();
app.use(cors({ origin: true, credentials: true }));
app.use(express.json({ limit: "5mb" }));
app.use(express.static(path.join(__dirname, "..")));

const dbPromise = open({
  filename: "./data.db",
  driver: sqlite3.Database
});

async function initDb() {
  const db = await dbPromise;
  await db.exec(`
    CREATE TABLE IF NOT EXISTS users (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      username TEXT UNIQUE NOT NULL,
      password_hash TEXT NOT NULL,
      is_admin INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL
    );
  `);
  await db.exec(`
    CREATE TABLE IF NOT EXISTS state_data (
      user_id INTEGER NOT NULL,
      system TEXT NOT NULL,
      payload TEXT NOT NULL,
      updated_at TEXT NOT NULL,
      PRIMARY KEY (user_id, system),
      FOREIGN KEY(user_id) REFERENCES users(id)
    );
  `);
  const existing = await db.get("SELECT id FROM users WHERE username = ?", [
    ADMIN_USER
  ]);
  if (!existing) {
    const hash = await bcrypt.hash(ADMIN_PASS, 10);
    await db.run(
      "INSERT INTO users (username, password_hash, is_admin, created_at) VALUES (?, ?, 1, ?)",
      [ADMIN_USER, hash, new Date().toISOString()]
    );
  }
}

function signToken(user) {
  return jwt.sign(
    { id: user.id, username: user.username, is_admin: !!user.is_admin },
    JWT_SECRET,
    { expiresIn: "12h" }
  );
}

function authMiddleware(requireAdmin = false) {
  return (req, res, next) => {
    const header = req.headers.authorization || "";
    const token = header.startsWith("Bearer ") ? header.slice(7) : null;
    if (!token) return res.status(401).json({ error: "unauthorized" });
    try {
      const payload = jwt.verify(token, JWT_SECRET);
      if (requireAdmin && !payload.is_admin) {
        return res.status(403).json({ error: "forbidden" });
      }
      req.user = payload;
      return next();
    } catch (error) {
      return res.status(401).json({ error: "unauthorized" });
    }
  };
}

app.get("/api/ping", (req, res) => {
  res.json({ ok: true });
});

app.post("/api/login", async (req, res) => {
  const { username, password } = req.body || {};
  if (!username || !password) {
    return res.status(400).json({ error: "missing_credentials" });
  }
  const db = await dbPromise;
  const user = await db.get("SELECT * FROM users WHERE username = ?", [
    username
  ]);
  if (!user) {
    return res.status(401).json({ error: "invalid_credentials" });
  }
  const ok = await bcrypt.compare(password, user.password_hash);
  if (!ok) {
    return res.status(401).json({ error: "invalid_credentials" });
  }
  return res.json({
    token: signToken(user),
    user: { id: user.id, username: user.username, is_admin: !!user.is_admin }
  });
});

app.get("/api/me", authMiddleware(), async (req, res) => {
  res.json({ user: req.user });
});

app.get("/api/admin/users", authMiddleware(true), async (req, res) => {
  const db = await dbPromise;
  const users = await db.all(
    "SELECT id, username, is_admin, created_at FROM users ORDER BY created_at DESC"
  );
  res.json({ users });
});

app.post("/api/admin/users", authMiddleware(true), async (req, res) => {
  const { username, password, is_admin } = req.body || {};
  if (!username || !password) {
    return res.status(400).json({ error: "missing_fields" });
  }
  const db = await dbPromise;
  const hash = await bcrypt.hash(password, 10);
  try {
    const result = await db.run(
      "INSERT INTO users (username, password_hash, is_admin, created_at) VALUES (?, ?, ?, ?)",
      [username, hash, is_admin ? 1 : 0, new Date().toISOString()]
    );
    return res.json({ id: result.lastID });
  } catch (error) {
    return res.status(400).json({ error: "user_exists" });
  }
});

app.delete("/api/admin/users/:id", authMiddleware(true), async (req, res) => {
  const id = Number(req.params.id);
  if (!id) return res.status(400).json({ error: "invalid_id" });
  const db = await dbPromise;
  await db.run("DELETE FROM users WHERE id = ?", [id]);
  await db.run("DELETE FROM state_data WHERE user_id = ?", [id]);
  res.json({ ok: true });
});

async function getState(system, userId) {
  const db = await dbPromise;
  return db.get(
    "SELECT payload FROM state_data WHERE user_id = ? AND system = ?",
    [userId, system]
  );
}

async function setState(system, userId, payload) {
  const db = await dbPromise;
  const existing = await db.get(
    "SELECT user_id FROM state_data WHERE user_id = ? AND system = ?",
    [userId, system]
  );
  const now = new Date().toISOString();
  if (!existing) {
    await db.run(
      "INSERT INTO state_data (user_id, system, payload, updated_at) VALUES (?, ?, ?, ?)",
      [userId, system, JSON.stringify(payload), now]
    );
  } else {
    await db.run(
      "UPDATE state_data SET payload = ?, updated_at = ? WHERE user_id = ? AND system = ?",
      [JSON.stringify(payload), now, userId, system]
    );
  }
}

app.get("/api/state/:system", authMiddleware(), async (req, res) => {
  const system = req.params.system;
  const row = await getState(system, req.user.id);
  if (!row) return res.json({ payload: null });
  return res.json({ payload: JSON.parse(row.payload) });
});

app.put("/api/state/:system", authMiddleware(), async (req, res) => {
  const system = req.params.system;
  const payload = req.body || {};
  await setState(system, req.user.id, payload);
  res.json({ ok: true });
});

app.get("/api/miete", authMiddleware(), async (req, res) => {
  const db = await dbPromise;
  const row = await getState("miete", req.user.id);
  if (!row) return res.json({ payload: null });
  return res.json({ payload: JSON.parse(row.payload) });
});

app.put("/api/miete", authMiddleware(), async (req, res) => {
  const payload = req.body || {};
  await setState("miete", req.user.id, payload);
  res.json({ ok: true });
});

initDb().then(() => {
  app.listen(PORT, () => {
    console.log(`Server listening on ${PORT}`);
  });
});
