import fs from 'node:fs';
const config=JSON.parse(fs.readFileSync(new URL('../wrangler.jsonc',import.meta.url),'utf8'));
const id=config.d1_databases?.[0]?.database_id;
if(!id || id==='00000000-0000-0000-0000-000000000000' || !/^[a-f0-9-]{36}$/i.test(id)) {
 console.error('Bitte zuerst die database_id aus „npx wrangler d1 create thesisflow“ in wrangler.jsonc eintragen.');
 process.exit(1);
}
