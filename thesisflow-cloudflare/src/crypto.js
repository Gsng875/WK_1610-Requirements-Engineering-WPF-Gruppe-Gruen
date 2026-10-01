// Web Crypto läuft sowohl im Cloudflare Worker als auch in Node.js.
const encoder = new TextEncoder();
export const hex = bytes => Array.from(new Uint8Array(bytes), b => b.toString(16).padStart(2,'0')).join('');
export const token = () => hex(crypto.getRandomValues(new Uint8Array(32)));
export async function digest(value) {
 return hex(await crypto.subtle.digest('SHA-256',encoder.encode(value)));
}
export async function passwordHash(password, salt, iterations = 100000) {
 const key = await crypto.subtle.importKey('raw',encoder.encode(password),'PBKDF2',false,['deriveBits']);
 return hex(await crypto.subtle.deriveBits({name:'PBKDF2',salt:encoder.encode(salt),iterations,hash:'SHA-256'},key,256));
}
export function equal(a,b) {
 if (a.length !== b.length) return false;
 let difference = 0;
 for (let i=0;i<a.length;i++) difference |= a.charCodeAt(i)^b.charCodeAt(i);
 return difference === 0;
}
