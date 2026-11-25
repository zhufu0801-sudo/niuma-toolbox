import test from 'node:test';import assert from 'node:assert/strict';import fs from 'node:fs';import path from 'node:path';
import {fileURLToPath} from 'node:url';
// 页面把本机口令存起来反复用时，服务重启后会一直 403。凡是缓存口令的页面，都必须在 403 时重新取口令再试。
test('缓存本机口令的页面都会在 403 时重新取口令',()=>{
  const dir=fileURLToPath(new URL('../public/',import.meta.url)),files=[];
  const walk=d=>{for(const e of fs.readdirSync(d,{withFileTypes:true})){const p=path.join(d,e.name);if(e.isDirectory())walk(p);else if(p.endsWith('.js')&&!/i18n/.test(e.name))files.push(p);}};walk(dir);
  const missing=files.filter(f=>{const s=fs.readFileSync(f,'utf8');return /if\(!token\)/.test(s)&&/\/api\/session/.test(s)&&!/403/.test(s);}).map(f=>path.basename(f));
  assert.deepEqual(missing,[]);
});
