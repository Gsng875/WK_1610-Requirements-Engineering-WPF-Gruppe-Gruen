// Kleine lokale Demo: HTTP-Server, SQLite und Sicherheit ohne npm-Pakete.
const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const { randomBytes, scryptSync, timingSafeEqual } = require('node:crypto');
const { DatabaseSync } = require('node:sqlite');

const dataDir = process.env.DATA_DIR || path.join(__dirname, 'data');
fs.mkdirSync(dataDir, { recursive: true });
const db = new DatabaseSync(path.join(dataDir, 'thesisflow.sqlite'));
db.exec(`PRAGMA journal_mode=WAL;
 CREATE TABLE IF NOT EXISTS users (id TEXT PRIMARY KEY, name TEXT, role TEXT, salt TEXT, hash TEXT);
 CREATE TABLE IF NOT EXISTS topics (id INTEGER PRIMARY KEY, examiner TEXT, payload TEXT,
 status TEXT DEFAULT 'offen', review TEXT DEFAULT '{}', decided_at TEXT);
 CREATE TABLE IF NOT EXISTS audit (id INTEGER PRIMARY KEY, topic INTEGER, actor TEXT, action TEXT, at TEXT);`);

// Rollen und Zuordnung legt das System fest, niemals das Anmeldeformular.
for (const [id, name, role, password] of [
 ['weber', 'Prof. Dr. Anna Weber', 'examiner', process.env.EXAMINER_PASSWORD || 'PrueferDemo!2026'],
 ['klein', 'Prof. Dr. Jonas Klein', 'examiner', process.env.OTHER_PASSWORD || 'AndereDemo!2026'],
 ['student', 'Moritz Hoffmann', 'student', process.env.STUDENT_PASSWORD || 'StudentDemo!2026']
]) {
 if (!db.prepare('SELECT id FROM users WHERE id=?').get(id)) {
  const salt = randomBytes(16).toString('hex');
  db.prepare('INSERT INTO users VALUES (?,?,?,?,?)').run(id, name, role, salt, scryptSync(password, salt, 64).toString('hex'));
 }
}

const example = {
 student: 'Moritz Hoffmann', matriculation: '1234567', degree: 'B.Sc. Wirtschaftsinformatik',
 department: 'Fachbereich MND', title: 'Konzeption eines digitalen Abschlussarbeiten-Portals',
 question: 'Wie kann ein Webportal den Anmelde- und Freigabeprozess von Abschlussarbeiten im Fachbereich MND vereinfachen?',
 context: 'Die Anmeldung erfolgt bislang über Formulare und E-Mails. Fehlende Angaben und manuelle Freigaben führen zu Rückfragen und schwer nachvollziehbaren Bearbeitungsständen.',
 goal: 'Ein Soll-Prozess mit Rollen, digitaler Themenfreigabe und transparenten Statusanzeigen wird konzipiert und durch einen anklickbaren Prototyp veranschaulicht.',
 method: 'Ist-Prozess und Formulare analysieren, Anforderungen mit User Stories beschreiben, einen Prototyp entwickeln und anhand ausgewählter Nutzungsszenarien prüfen.',
 scope: 'Betrachtet werden Anmeldung und Themenfreigabe. Benotung, Plagiatsprüfung und produktive Hochschulanbindung sind nicht Bestandteil der Arbeit.',
 deliverables: 'Prozessmodell, spezifizierte User Stories, UI-Prototyp und dokumentierte Evaluation.',
 secondExaminer: 'Prof. Dr. Lara Neumann (fiktiv)', plannedStart: '2026-10-15', complete: true, authorized: true
};

for (const [id, examiner, payload] of [
 [1, 'weber', example],
 [2, 'weber', {...example, student: 'Tarek Yilmaz', matriculation: '7654321', degree: 'M.Sc. Wirtschaftsinformatik', title: 'Digitale Freigabeprozesse im Unternehmen', authorized: false}],
 [3, 'klein', {...example, student: 'Lea Sommer', title: 'Visualisierung betrieblicher Kennzahlen'}]
]) db.prepare('INSERT OR IGNORE INTO topics (id,examiner,payload) VALUES (?,?,?)').run(id, examiner, JSON.stringify(payload));

const sessions = new Map(), attempts = new Map();
const token = () => randomBytes(32).toString('hex');
const fail = (status, message) => { throw Object.assign(new Error(message), {status}); };

const getTopic = (id, user) => {
 const row = db.prepare('SELECT * FROM topics WHERE id=? AND examiner=?').get(id, user.id);
 if (!row) fail(404, 'Das Thema ist nicht vorhanden oder Ihnen nicht zugeordnet.');
 return row;
};

const present = row => ({
 id: row.id,
 ...JSON.parse(row.payload),
 status: row.status,
 review: JSON.parse(row.review),
 decidedAt: row.decided_at
});

setInterval(() => {
 const now = Date.now();
 for (const [key, s] of sessions) if (s.expires < now) sessions.delete(key);
 for (const [key, a] of attempts) if (a.until < now) attempts.delete(key);
}, 60000).unref();

async function body(req) {
 if (!req.headers['content-type']?.startsWith('application/json')) fail(415, 'JSON erwartet.');
 let raw = '';
 for await (const chunk of req) {
  raw += chunk;
  if (Buffer.byteLength(raw) > 12000) fail(413, 'Eingabe zu groß.');
 }
 try {
  const b = JSON.parse(raw);
  if (!b || typeof b !== 'object' || Array.isArray(b)) throw Error();
  return b;
 } catch {
  fail(400, 'Ungültige Eingabe.');
 }
}

const server = http.createServer(async (req, res) => {
 res.setHeader('Cache-Control', 'no-store');
 res.setHeader('X-Content-Type-Options', 'nosniff');
 res.setHeader('Referrer-Policy', 'no-referrer');
 res.setHeader('Content-Security-Policy', "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self'; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'");

 const send = (status, value) => {
  res.writeHead(status, {'Content-Type':'application/json; charset=utf-8'});
  res.end(JSON.stringify(value));
 };

 try {
  const url = new URL(req.url, 'http://localhost');

  // Nur festgelegte Dateien ausliefern; kein Zugriff auf Datenbank oder Quellcode.
  const files = {'/':'index.html', '/app.js':'app.js', '/style.css':'style.css'};
  if (req.method === 'GET' && files[url.pathname]) {
   const type = url.pathname.endsWith('.js') ? 'text/javascript' : url.pathname.endsWith('.css') ? 'text/css' : 'text/html';
   res.writeHead(200, {'Content-Type':type+'; charset=utf-8'});
   return res.end(fs.readFileSync(path.join(__dirname, 'public', files[url.pathname])));
  }

  const sessionId = /(?:^|;\s*)session=([a-f0-9]{64})(?:;|$)/.exec(req.headers.cookie || '')?.[1];
  const session = sessions.get(sessionId);
  if (session && session.expires < Date.now()) sessions.delete(sessionId);
  const user = session && session.expires >= Date.now() ? session.user : null;

  // Änderungen benötigen dieselbe Origin und nach Login zusätzlich ein CSRF-Token.
  if (req.method !== 'GET') {
   const expected = process.env.APP_ORIGIN || 'http://localhost:8080';
   if (req.headers.origin !== expected) fail(403, 'Anfrage von einer nicht erlaubten Herkunft.');
   if (user && req.headers['x-csrf-token'] !== session.csrf) fail(403, 'Sicherheitstoken fehlt. Bitte neu anmelden.');
  }

  if (url.pathname === '/api/login' && req.method === 'POST') {
   const b = await body(req);
   if (typeof b.username !== 'string' || typeof b.password !== 'string' || b.username.length > 100 || b.password.length > 200) fail(400, 'Ungültige Anmeldedaten.');

   const key = req.socket.remoteAddress, now = Date.now();
   const a = attempts.get(key) || {count:0, until:now+15*60000};
   if (a.until < now) {a.count=0; a.until=now+15*60000;}
   if (a.count >= 5) fail(429, 'Zu viele Fehlversuche. Bitte nach 15 Minuten erneut versuchen.');

   const u = db.prepare('SELECT * FROM users WHERE id=?').get(b.username.trim().toLowerCase());
   const hash = scryptSync(b.password, u?.salt || 'dummy-salt', 64);

   if (!u || !timingSafeEqual(hash, Buffer.from(u.hash,'hex'))) {
    a.count++;
    attempts.set(key,a);
    fail(401, 'Benutzername oder Passwort ist falsch.');
   }

   if (u.role !== 'examiner') fail(403, 'Zugriff gesperrt: Nur registrierte Erstprüfer:innen dürfen diese Anwendung nutzen.');

   attempts.delete(key);
   if (sessionId) sessions.delete(sessionId);
   const id = token(), csrf = token();
   sessions.set(id, {user:{id:u.id,name:u.name}, csrf, expires:now+30*60000});

   res.setHeader('Set-Cookie', `session=${id}; HttpOnly; SameSite=Strict; Path=/${process.env.COOKIE_SECURE === 'true' ? '; Secure' : ''}`);
   return send(200, {user:{id:u.id,name:u.name},csrf});
  }

  if (!user) fail(401, 'Bitte als Erstprüfer:in anmelden.');
  session.expires = Date.now()+30*60000;

  if (url.pathname === '/api/me' && req.method === 'GET') return send(200, {user,csrf:session.csrf});

  if (url.pathname === '/api/logout' && req.method === 'POST') {
   sessions.delete(sessionId);
   res.setHeader('Set-Cookie','session=; HttpOnly; SameSite=Strict; Path=/; Max-Age=0');
   return send(200,{});
  }

  if (url.pathname === '/api/topics' && req.method === 'GET') {
   return send(200, db.prepare('SELECT * FROM topics WHERE examiner=? ORDER BY id').all(user.id).map(present));
  }

  const match = /^\/api\/topics\/(\d+)(?:\/(review|decision))?$/.exec(url.pathname);
  if (!match) fail(404,'Nicht gefunden.');

  const row = getTopic(Number(match[1]),user);

  if (req.method === 'GET' && !match[2]) {
   return send(200,{
    ...present(row),
    history:db.prepare('SELECT actor,action,at FROM audit WHERE topic=? ORDER BY id DESC').all(row.id)
   });
  }

  if (req.method !== 'POST' || !['review','decision'].includes(match[2])) fail(405,'Aktion nicht erlaubt.');

  const b = await body(req);
  if (!['geeignet','ueberarbeiten','ungeeignet'].includes(b.rating) || typeof b.comment !== 'string' || b.comment.trim().length < 10 || b.comment.length > 2000) fail(400,'Bitte eine Einschätzung und eine Begründung mit 10 bis 2000 Zeichen angeben.');

  const review = JSON.stringify({
   rating:b.rating,
   comment:b.comment.trim(),
   academic:b.academic === true,
   scope:b.scope === true
  });

  let status = 'offen', action = 'Einschätzung gespeichert';

  if (match[2] === 'decision') {
   if (!['freigegeben','abgelehnt'].includes(b.decision) || b.confirmed !== true) fail(400,'Entscheidung muss ausdrücklich bestätigt werden.');

   const topic = JSON.parse(row.payload);
   if (!topic.complete) fail(409,'Die Pflichtangaben zum Thema sind unvollständig.');

   if (b.decision === 'freigegeben' && (b.rating !== 'geeignet' || b.academic !== true || b.scope !== true)) fail(409,'Freigabe benötigt ein geeignetes Thema und beide bestätigten Prüfkriterien.');

   status = b.decision;
   action = status === 'freigegeben' ? 'Thema bestätigt und freigegeben' : 'Thema abgelehnt';
  }

  const at = new Date().toISOString();

  // Jede Entscheidung und ihr Protokolleintrag werden zusammen gespeichert.
  db.exec('BEGIN IMMEDIATE');
  try {
   const result = db.prepare("UPDATE topics SET review=?,status=?,decided_at=? WHERE id=?").run(review,status,status==='offen'?null:at,row.id);
   if (!result.changes) fail(404,'Das Thema wurde nicht gefunden.');

   db.prepare('INSERT INTO audit (topic,actor,action,at) VALUES (?,?,?,?)').run(row.id,user.name,action,at);
   db.exec('COMMIT');
  } catch(e) {
   db.exec('ROLLBACK');
   throw e;
  }

  return send(200,{message:action});
 } catch(e) {
  if (!e.status) console.error(e);
  if (!res.headersSent) send(e.status || 500,{error:e.status?e.message:'Interner Fehler. Bitte erneut versuchen.'});
  else res.end();
 }
});

server.requestTimeout = 15000;
server.listen(Number(process.env.PORT || 8080),'0.0.0.0',()=>console.log('ThesisFlow: http://localhost:'+ (process.env.PORT || 8080)));