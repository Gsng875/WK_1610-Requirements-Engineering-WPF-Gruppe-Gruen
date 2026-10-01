import {test} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {Miniflare,convertV4MiniflareOptions} from 'miniflare';
import {passwordHash,digest} from '../src/crypto.js';
import {topics} from '../scripts/examples.mjs';

const root=fileURLToPath(new URL('../',import.meta.url));
test('Worker: Zugriff, CSRF, D1-Persistenz und konkurrierende Entscheidungen',async()=>{
 const folder=fs.mkdtempSync(path.join(os.tmpdir(),'thesisflow-d1-'));
 let mf;
 const start=()=>mf=new Miniflare({...convertV4MiniflareOptions({modulesRoot:path.join(root,'src'),
  modules:['worker.js','crypto.js'].map(file=>({type:'ESModule',path:path.join(root,'src',file)})),
  compatibilityDate:'2026-09-30',d1Databases:{DB:'thesisflow-test'},d1Persist:folder,
  serviceBindings:{ASSETS:request=>new Response('Static asset: '+new URL(request.url).pathname)}}),
  resourcePersistencePath:folder,isolatedResourcePersistencePath:folder});
 const request=async(url,{data,cookie,csrf,origin='https://demo.example',ip='192.0.2.1',method}={})=>{
  const response=await mf.dispatchFetch('https://demo.example'+url,{method:method||(data?'POST':'GET'),
   headers:{...(data?{'Content-Type':'application/json',Origin:origin}:{}),...(cookie?{Cookie:cookie}:{}),...(csrf?{'X-CSRF-Token':csrf}:{}),'CF-Connecting-IP':ip},body:data?JSON.stringify(data):undefined});
  const text=await response.text();let value;try{value=JSON.parse(text);}catch{value=text;}
  return {status:response.status,value,cookie:response.headers.get('Set-Cookie')?.split(';')[0],headers:response.headers};
 };
 const login=(username,password,options={})=>request('/api/login',{...options,data:{username,password}});
 try {
  start();let db=await mf.getD1Database('DB');
  const schema=fs.readFileSync(path.join(root,'schema.sql'),'utf8').replace(/--[^\n]*/g,'');
  await db.batch(schema.split(';').filter(sql=>sql.trim()).map(sql=>db.prepare(sql)));
  for(const [id,name,role] of [['weber','Prof. Dr. Anna Weber','examiner'],['klein','Prof. Dr. Jonas Klein','examiner'],['student','Moritz Hoffmann','student']]) {
   await db.prepare('INSERT INTO users(id,name,role,salt,hash) VALUES(?,?,?,?,?)').bind(id,name,role,'test-salt-'+id,await passwordHash('TestPassword!2026','test-salt-'+id)).run();
  }
  for(const [id,examiner,payload] of topics)await db.prepare('INSERT INTO topics(id,examiner,payload) VALUES(?,?,?)').bind(id,examiner,JSON.stringify(payload)).run();
  assert.equal((await request('/')).status,200);
  for(const file of ['/schema.sql','/src/worker.js','/credentials.local.txt','/seed.local.sql'])assert.equal((await request(file)).status,404);
  assert.equal((await request('/api/topics')).status,401);
  assert.equal((await login('student','TestPassword!2026')).status,403);
  assert.equal((await login('weber','TestPassword!2026',{origin:'https://attacker.example'})).status,403);
  const account=await login('weber','TestPassword!2026');assert.equal(account.status,200);
  assert.match(account.headers.get('Set-Cookie'),/HttpOnly/);assert.match(account.headers.get('Set-Cookie'),/Secure/);
  const auth={cookie:account.cookie,csrf:account.value.csrf};
  assert.deepEqual((await request('/api/topics',auth)).value.map(t=>t.id),[1,2]);
  assert.equal((await request('/api/topics/3',auth)).status,404);
  const review={rating:'geeignet',comment:'Das Thema ist fachlich geeignet.',academic:true,scope:true};
  assert.equal((await request('/api/topics/1/review',{cookie:auth.cookie,data:review})).status,403);
  assert.equal((await request('/api/topics/1/review',{...auth,data:review,origin:'https://attacker.example'})).status,403);
  assert.equal((await request('/api/topics/1/review',{...auth,data:{...review,comment:'kurz'}})).status,400);
  assert.equal((await request('/api/topics/1/review',{...auth,data:review})).status,200);
  // Die Sitzung überlebt den Wechsel der Worker-Instanz; keine globale Map.
  await mf.dispose();start();db=await mf.getD1Database('DB');
  assert.equal((await request('/api/me',auth)).value.user.id,'weber');
  const decision={...review,decision:'freigegeben',confirmed:true};
  assert.equal((await request('/api/topics/1/decision',{...auth,data:{...decision,confirmed:false}})).status,400);
  assert.equal((await request('/api/topics/1/decision',{...auth,data:{...decision,academic:false}})).status,409);
  assert.equal((await request('/api/topics/2/decision',{...auth,data:decision})).status,409);
  const attempts=await Promise.all([request('/api/topics/1/decision',{...auth,data:decision}),request('/api/topics/1/decision',{...auth,data:decision})]);
  assert.deepEqual(attempts.map(r=>r.status).sort(),[200,409]);
  let topic=(await request('/api/topics/1',auth)).value;
  assert.equal(topic.status,'freigegeben');assert.equal(topic.history.length,2);assert.ok(topic.decidedAt);
  assert.equal(topic.history[0].actor,'Prof. Dr. Anna Weber');
  assert.equal((await request('/api/topics/1/review',{...auth,data:review})).status,409);
  await db.prepare('UPDATE topics SET payload=? WHERE id=2').bind(JSON.stringify({...topics[1][2],complete:false})).run();
  const reject={...decision,rating:'ungeeignet',decision:'abgelehnt'};
  assert.equal((await request('/api/topics/2/decision',{...auth,data:reject})).status,409);
  await db.prepare('UPDATE topics SET payload=? WHERE id=2').bind(JSON.stringify(topics[1][2])).run();
  assert.equal((await request('/api/topics/2/decision',{...auth,data:reject})).status,200);
  const other=await login('klein','TestPassword!2026');
  assert.deepEqual((await request('/api/topics',{cookie:other.cookie})).value.map(t=>t.id),[3]);
  assert.equal((await request('/api/topics/1',{cookie:other.cookie})).status,404);
  assert.equal((await request('/api/logout',{...auth,data:{}})).status,200);
  assert.equal((await request('/api/me',auth)).status,401);
  const fresh=await login('weber','TestPassword!2026');
  await db.prepare('UPDATE sessions SET expires=? WHERE id=?').bind(Date.now()-1,await digest(fresh.cookie.split('=')[1])).run();
  assert.equal((await request('/api/me',{cookie:fresh.cookie})).status,401);
  for(let i=0;i<5;i++)assert.equal((await login('weber','falsch',{ip:'192.0.2.2'})).status,401);
  await mf.dispose();start();
  assert.equal((await login('weber','TestPassword!2026',{ip:'192.0.2.2'})).status,429);
  assert.equal((await request('/api/topics/1',{cookie:other.cookie})).status,404);
 }finally{await mf?.dispose();fs.rmSync(folder,{recursive:true,force:true});}
});
