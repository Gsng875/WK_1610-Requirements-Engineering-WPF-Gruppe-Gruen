-- Wiederholbar: vorhandene Benutzer, Themen und Entscheidungen bleiben erhalten.
PRAGMA foreign_keys = ON;
CREATE TABLE IF NOT EXISTS users (
 id TEXT PRIMARY KEY, name TEXT NOT NULL,
 role TEXT NOT NULL CHECK(role IN ('examiner','student')),
 salt TEXT NOT NULL, hash TEXT NOT NULL,
 iterations INTEGER NOT NULL DEFAULT 100000
);
CREATE TABLE IF NOT EXISTS topics (
 id INTEGER PRIMARY KEY, examiner TEXT NOT NULL REFERENCES users(id),
 payload TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'offen'
 CHECK(status IN ('offen','freigegeben','abgelehnt')),
 review TEXT NOT NULL DEFAULT '{}', decided_at TEXT
);
CREATE TABLE IF NOT EXISTS audit (
 id INTEGER PRIMARY KEY, topic INTEGER NOT NULL REFERENCES topics(id),
 actor TEXT NOT NULL, action TEXT NOT NULL, at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS sessions (
 id TEXT PRIMARY KEY, user_id TEXT NOT NULL REFERENCES users(id),
 csrf TEXT NOT NULL, expires INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS attempts (
 id TEXT PRIMARY KEY, count INTEGER NOT NULL, until INTEGER NOT NULL
);
CREATE INDEX IF NOT EXISTS topics_examiner ON topics(examiner);
CREATE INDEX IF NOT EXISTS audit_topic ON audit(topic);
CREATE INDEX IF NOT EXISTS sessions_expiry ON sessions(expires);
