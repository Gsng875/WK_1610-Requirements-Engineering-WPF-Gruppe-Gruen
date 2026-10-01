import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {spawnSync} from 'node:child_process';
import {passwordHash,token} from '../src/crypto.js';
import {topics} from './examples.mjs';

const root=fileURLToPath(new URL('../',import.meta.url));
process.chdir(root);
const mode=process.argv[2];
if(!['--local','--remote'].includes(mode))throw Error('Bitte --local oder --remote angeben.');
const sqlPath=path.join(root,'seed.local.sql'),credentials=path.join(root,'credentials.local.txt');
const quote=value=>"'"+String(value).replaceAll("'","''")+"'";
// Dieselbe private Seed-Datei lässt sich lokal und online wiederverwenden.
if(!fs.existsSync(sqlPath)) {
 const sql=[],login=['ThesisFlow – private Demo-Zugangsdaten','Nicht in Git hochladen oder öffentlich teilen.',''];
 for(const [id,name,role] of [
  ['weber','Prof. Dr. Anna Weber','examiner'],
  ['klein','Prof. Dr. Jonas Klein','examiner'],
  ['student','Moritz Hoffmann','student']
 ]) {
  const password=token().slice(0,24)+'!A',salt=token();
  const hash=await passwordHash(password,salt);
  sql.push(`INSERT OR IGNORE INTO users(id,name,role,salt,hash,iterations) VALUES(${[id,name,role,salt,hash].map(quote).join(',')},100000);`);
  login.push(`Benutzer: ${id}\nPasswort: ${password}\nRolle: ${role}\n`);
 }
 for(const [id,examiner,payload] of topics)sql.push(`INSERT OR IGNORE INTO topics(id,examiner,payload) VALUES(${id},${quote(examiner)},${quote(JSON.stringify(payload))});`);
 fs.writeFileSync(sqlPath,sql.join('\n')+'\n',{mode:0o600});
 fs.writeFileSync(credentials,login.join('\n'),{mode:0o600});
}
const cli=path.join(root,'node_modules/wrangler/bin/wrangler.js');
const result=spawnSync(process.execPath,[cli,'d1','execute','thesisflow',mode,'--file=seed.local.sql','--yes'],{stdio:'inherit'});
if(result.error)throw result.error;
if(result.status!==0)process.exit(result.status||1);
console.log('Beispieldaten eingerichtet. Zugangsdaten: credentials.local.txt');
console.log('Vorhandene Konten und Entscheidungen wurden nicht überschrieben.');
