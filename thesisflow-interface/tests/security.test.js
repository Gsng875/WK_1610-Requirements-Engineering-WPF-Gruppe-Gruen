const {test} = require('node:test');
const assert = require('node:assert/strict');
const {spawn} = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

test('Rollen, Themenzuordnung, Validierung und dauerhafte Entscheidung', async () => {
 const folder=fs.mkdtempSync(path.join(os.tmpdir(),'thesisflow-test-'));
 let processHandle;

 const start=async()=>{
  processHandle=spawn(process.execPath,[path.join(__dirname,'../server.js')],{env:{...process.env,PORT:'8091',APP_ORIGIN:'http://localhost:8091',DATA_DIR:folder,EXAMINER_PASSWORD:'PrueferDemo!2026',OTHER_PASSWORD:'AndereDemo!2026',STUDENT_PASSWORD:'StudentDemo!2026'},stdio:['ignore','pipe','pipe']});
  await new Promise((resolve,reject)=>{processHandle.stdout.once('data',resolve);processHandle.once('error',reject);processHandle.once('exit',code=>reject(Error('Server beendet: '+code)));});
 };

 const stop=()=>new Promise(resolve=>{processHandle.once('exit',resolve);processHandle.kill();});

 const request=async(url,{data,cookie,csrf,origin='http://localhost:8091'}={})=>{
  const r=await fetch('http://localhost:8091'+url,{method:data?'POST':'GET',headers:{...(data?{'Content-Type':'application/json',Origin:origin}:{}),...(cookie?{Cookie:cookie}:{}),...(csrf?{'X-CSRF-Token':csrf}:{})},body:data?JSON.stringify(data):undefined});
  return {status:r.status,value:await r.json(),cookie:r.headers.get('set-cookie')?.split(';')[0]};
 };

 const login=async(username,password)=>request('/api/login',{data:{username,password}});

 try {
  await start();
  assert.equal((await request('/api/topics')).status,401);

  for(const url of ['/server.js','/data/thesisflow.sqlite','/.env'])assert.equal((await request(url)).status,401);

  assert.equal((await login('student','StudentDemo!2026')).status,403);

  let account=await login('weber','PrueferDemo!2026');assert.equal(account.status,200);
  const auth={cookie:account.cookie,csrf:account.value.csrf};

  const topics=await request('/api/topics',auth);assert.deepEqual(topics.value.map(t=>t.id),[1,2]);
  assert.equal((await request('/api/topics/3',auth)).status,404);

  const review={rating:'geeignet',comment:'Das Thema ist wissenschaftlich geeignet.',academic:true,scope:true};

  assert.equal((await request('/api/topics/1/review',{cookie:auth.cookie,data:review})).status,403);
  assert.equal((await request('/api/topics/1/review',{...auth,data:review,origin:'http://fremd.example'})).status,403);
  assert.equal((await request('/api/topics/1/review',{...auth,data:{...review,comment:'kurz'}})).status,400);

  const decision={...review,decision:'freigegeben',confirmed:true};

  assert.equal((await request('/api/topics/1/decision',{...auth,data:{...decision,confirmed:false}})).status,400);
  assert.equal((await request('/api/topics/1/decision',{...auth,data:{...decision,academic:false}})).status,409);
  assert.equal((await request('/api/topics/2/decision',{...auth,data:decision})).status,200);
  assert.equal((await request('/api/topics/1/review',{...auth,data:review})).status,200);
  assert.equal((await request('/api/topics/1/decision',{...auth,data:decision})).status,200);
  assert.equal((await request('/api/topics/1/decision',{...auth,data:decision})).status,200);

  let topic=(await request('/api/topics/1',auth)).value;
  assert.equal(topic.status,'freigegeben');assert.equal(topic.history.length,3);assert.equal(topic.history[0].actor,'Prof. Dr. Anna Weber');assert.ok(topic.decidedAt);

  assert.equal((await request('/api/topics/2/decision',{...auth,data:{...decision,rating:'ungeeignet',decision:'abgelehnt'}})).status,200);

  const other=await login('klein','AndereDemo!2026');
  assert.equal((await request('/api/topics/1',{cookie:other.cookie})).status,404);
  assert.deepEqual((await request('/api/topics',{cookie:other.cookie})).value.map(t=>t.id),[3]);

  assert.equal((await request('/api/logout',{...auth,data:{}})).status,200);
  assert.equal((await request('/api/topics',auth)).status,401);

  await stop();await start();

  account=await login('weber','PrueferDemo!2026');
  topic=(await request('/api/topics/1',{cookie:account.cookie})).value;
  assert.equal(topic.status,'freigegeben');assert.equal(topic.review.comment,review.comment);
  assert.equal((await request('/api/topics/2',{cookie:account.cookie})).value.status,'abgelehnt');

  for(let i=0;i<5;i++)assert.equal((await login('weber','falsch')).status,401);
  assert.equal((await login('weber','PrueferDemo!2026')).status,429);
 } finally {
  if(processHandle?.exitCode===null)await stop();
  fs.rmSync(folder,{recursive:true,force:true});
 }
});