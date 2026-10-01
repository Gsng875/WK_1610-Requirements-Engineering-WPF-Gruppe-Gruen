import {token, digest, passwordHash, equal} from './crypto.js';

const fail = (status,message) => {throw Object.assign(new Error(message),{status});};
const json = (value,status=200,headers={}) => Response.json(value,{status,headers});
const present = row => ({id:row.id,...JSON.parse(row.payload),status:row.status,
 review:JSON.parse(row.review),decidedAt:row.decided_at});
const statement = (env,sql,...values) => env.DB.prepare(sql).bind(...values);

// Größe wird während des Lesens begrenzt, nicht nur über den Content-Length-Header.
async function body(request) {
 if (!request.headers.get('Content-Type')?.startsWith('application/json')) fail(415,'JSON erwartet.');
 if (!request.body) fail(400,'Eingabe fehlt.');
 const reader=request.body.getReader(), chunks=[];
 let size=0;
 while(true) {
  const {done,value}=await reader.read(); if(done)break;
  size+=value.length;
  if(size>12000){await reader.cancel();fail(413,'Eingabe zu groß.');}
  chunks.push(value);
 }
 const bytes=new Uint8Array(size);let offset=0;
 for(const chunk of chunks){bytes.set(chunk,offset);offset+=chunk.length;}
 try {const value=JSON.parse(new TextDecoder().decode(bytes));if(!value || typeof value!=='object' || Array.isArray(value))throw Error();return value;}
 catch {fail(400,'Ungültige Eingabe.');}
}

async function handle(request,env) {
 const url=new URL(request.url), now=Date.now();
 // Nur die Oberfläche liegt in public. Themen werden ausschließlich per API geladen.
 if(!url.pathname.startsWith('/api/')) {
  if(!['GET','HEAD'].includes(request.method))fail(405,'Aktion nicht erlaubt.');
  if(!['/','/index.html','/app.js','/style.css','/favicon.ico'].includes(url.pathname))fail(404,'Nicht gefunden.');
  return env.ASSETS.fetch(request);
 }
 if(!env.DB)fail(503,'Die D1-Datenbank ist noch nicht verbunden.');
 if(request.method!=='GET' && request.headers.get('Origin')!==url.origin) fail(403,'Anfrage von einer nicht erlaubten Herkunft.');
 const rawId=/(?:^|;\s*)session=([a-f0-9]{64})(?:;|$)/.exec(request.headers.get('Cookie')||'')?.[1];
 const id=rawId?await digest(rawId):'';
 // Gemeinsame D1-Sitzungen funktionieren auch über mehrere Worker-Instanzen hinweg.
 const session=id?await statement(env,`SELECT s.csrf,s.expires,u.id,u.name,u.role FROM sessions s
  JOIN users u ON u.id=s.user_id WHERE s.id=? AND s.expires>?`,id,now).first():null;
 const user=session?.role==='examiner'?{id:session.id,name:session.name}:null;
 if(user && request.method!=='GET' && !equal(request.headers.get('X-CSRF-Token')||'',session.csrf)) fail(403,'Sicherheitstoken fehlt. Bitte neu anmelden.');
 const secure=url.protocol==='https:'?'; Secure':'';
 const cookie=value=>`session=${value}; HttpOnly; SameSite=Strict; Path=/${secure}`;

 if(url.pathname==='/api/login' && request.method==='POST') {
  const b=await body(request);
  if(typeof b.username!=='string' || typeof b.password!=='string' || b.username.length>100 || b.password.length>200)fail(400,'Ungültige Anmeldedaten.');
  // CF-Connecting-IP wird an der Cloudflare-Kante gesetzt. Kein Klartext-IP-Speicher.
  const key=await digest(request.headers.get('CF-Connecting-IP')||'local');
  const attempt=await statement(env,`INSERT INTO attempts(id,count,until) VALUES(?,1,?)
   ON CONFLICT(id) DO UPDATE SET
   count=CASE WHEN until<=? THEN 1 ELSE count+1 END,
   until=CASE WHEN until<=? THEN excluded.until ELSE until END RETURNING count`,key,now+15*60000,now,now).first();
  if(attempt.count>5)fail(429,'Zu viele Anmeldeversuche. Bitte nach 15 Minuten erneut versuchen.');
  const u=await statement(env,'SELECT * FROM users WHERE id=?',b.username.trim().toLowerCase()).first();
  const hash=await passwordHash(b.password,u?.salt||'dummy-salt',u?.iterations||100000);
  if(!u || !equal(hash,u.hash))fail(401,'Benutzername oder Passwort ist falsch.');
  if(u.role!=='examiner')fail(403,'Zugriff gesperrt: Nur registrierte Erstprüfer:innen dürfen diese Anwendung nutzen.');
  const raw=token(), csrf=token();
  await env.DB.batch([
   statement(env,'DELETE FROM attempts WHERE id=? OR until<=?',key,now),
   statement(env,'DELETE FROM sessions WHERE id=? OR expires<=?',id,now),
   statement(env,'INSERT INTO sessions(id,user_id,csrf,expires) VALUES(?,?,?,?)',await digest(raw),u.id,csrf,now+30*60000)
  ]);
  return json({user:{id:u.id,name:u.name},csrf},200,{'Set-Cookie':cookie(raw)});
 }
 if(!user)fail(401,'Bitte als Erstprüfer:in anmelden.');
 if(url.pathname==='/api/logout' && request.method==='POST') {
  await statement(env,'DELETE FROM sessions WHERE id=?',id).run();
  return json({},200,{'Set-Cookie':cookie('')+'; Max-Age=0'});
 }
 await statement(env,'UPDATE sessions SET expires=? WHERE id=? AND expires>?',now+30*60000,id,now).run();
 if(url.pathname==='/api/me' && request.method==='GET')return json({user,csrf:session.csrf});
 if(url.pathname==='/api/topics' && request.method==='GET') {
  const {results}=await statement(env,'SELECT * FROM topics WHERE examiner=? ORDER BY id',user.id).all();
  return json(results.map(present));
 }
 const match=/^\/api\/topics\/(\d+)(?:\/(review|decision))?$/.exec(url.pathname);
 if(!match)fail(404,'Nicht gefunden.');
 const row=await statement(env,'SELECT * FROM topics WHERE id=? AND examiner=?',Number(match[1]),user.id).first();
 if(!row)fail(404,'Das Thema ist nicht vorhanden oder Ihnen nicht zugeordnet.');
 if(request.method==='GET' && !match[2]) {
  const {results}=await statement(env,'SELECT actor,action,at FROM audit WHERE topic=? ORDER BY id DESC',row.id).all();
  return json({...present(row),history:results});
 }
 if(request.method!=='POST' || !['review','decision'].includes(match[2]))fail(405,'Aktion nicht erlaubt.');
 if(row.status!=='offen')fail(409,'Dieses Thema wurde bereits abschließend entschieden.');
 const b=await body(request);
 if(!['geeignet','ueberarbeiten','ungeeignet'].includes(b.rating) || typeof b.comment!=='string' || b.comment.trim().length<10 || b.comment.length>2000)fail(400,'Bitte eine Einschätzung und eine Begründung mit 10 bis 2000 Zeichen angeben.');
 const review=JSON.stringify({rating:b.rating,comment:b.comment.trim(),academic:b.academic===true,scope:b.scope===true});
 let status='offen',action='Einschätzung gespeichert';
 const topic=JSON.parse(row.payload);
 if(match[2]==='decision') {
  if(!['freigegeben','abgelehnt'].includes(b.decision) || b.confirmed!==true)fail(400,'Entscheidung muss ausdrücklich bestätigt werden.');
  if(!topic.complete)fail(409,'Die Pflichtangaben zum Thema sind unvollständig.');
  if(b.decision==='freigegeben' && (!topic.authorized || b.rating!=='geeignet' || b.academic!==true || b.scope!==true))fail(409,'Freigabe benötigt die vorherige Zulassung, ein geeignetes Thema und beide bestätigten Prüfkriterien.');
  status=b.decision;action=status==='freigegeben'?'Thema bestätigt und freigegeben':'Thema abgelehnt';
 }
 const at=new Date().toISOString();
 // D1-batch ist atomar. changes() protokolliert ausschließlich erfolgreiche Updates.
 const [updated]=await env.DB.batch([
  statement(env,`UPDATE topics SET review=?,status=?,decided_at=? WHERE id=? AND examiner=? AND status='offen'
   AND (?='offen' OR json_extract(payload,'$.complete')=1)
   AND (?!='freigegeben' OR json_extract(payload,'$.authorized')=1)`,review,status,status==='offen'?null:at,row.id,user.id,status,status),
  statement(env,'INSERT INTO audit(topic,actor,action,at) SELECT ?,?,?,? WHERE changes()=1',row.id,user.name,action,at)
 ]);
 if(!updated.meta.changes)fail(409,'Der Vorgang wurde zwischenzeitlich geändert. Bitte neu laden.');
 return json({message:action});
}

export default {
 async fetch(request,env) {
  let response;
  try {response=await handle(request,env);}
  catch(e){if(!e.status)console.error('Worker error:',e.message);response=json({error:e.status?e.message:'Interner Fehler. Bitte Datenbankeinrichtung prüfen.'},e.status||500);}
  const secured=new Response(response.body,response);
  secured.headers.set('Cache-Control','no-store');
  secured.headers.set('X-Content-Type-Options','nosniff');
  secured.headers.set('Referrer-Policy','no-referrer');
  secured.headers.set('Content-Security-Policy',"default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self'; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'");
  if(new URL(request.url).protocol==='https:')secured.headers.set('Strict-Transport-Security','max-age=31536000');
  return secured;
 }
};
