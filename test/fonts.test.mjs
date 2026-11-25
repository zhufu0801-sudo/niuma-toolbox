import test from 'node:test';
import assert from 'node:assert/strict';
import {fontCandidates,findCjkFont} from '../lib/fonts.mjs';
import {ensurePdfEnv} from '../lib/pdf-env.mjs';
test('PDF 字体候选：Windows 雅黑、Mac 苹方优先，可以用 COACH_PDF_FONT 指定 .ttc 里的字体名',()=>{
  assert.deepEqual(fontCandidates({WINDIR:'C:\\Windows'},'win32')[0],{file:'C:\\Windows\\Fonts\\msyh.ttc',family:'MicrosoftYaHei'});
  assert.equal(fontCandidates({},'darwin')[0].family,'PingFangSC-Regular');
  assert.deepEqual(fontCandidates({COACH_PDF_FONT:'/x/a.ttc#Foo'},'linux')[0],{file:'/x/a.ttc',family:'Foo'});
  const f=findCjkFont({COACH_PDF_FONT:'/不存在/x.ttf'},'win32-none');assert.notEqual(f&&f.file,'/不存在/x.ttf');
});
test('读 PDF 前补上 DOMMatrix：简化版能做平移、缩放、求逆',async()=>{
  await ensurePdfEnv();const M=globalThis.DOMMatrix;const m=new M([2,0,0,2,10,20]);const p=m.transformPoint({x:1,y:1});assert.deepEqual([p.x,p.y],[12,22]);
  const back=m.inverse().transformPoint(p);assert.ok(Math.abs(back.x-1)<1e-9&&Math.abs(back.y-1)<1e-9);
});
test('Mac 上 API Key 存进钥匙串：ai.json 只留编号，读取时从钥匙串取回',async()=>{
  const {createSecretBox}=await import('../lib/secret.mjs');const store=new Map(),calls=[];
  const keychain=async args=>{calls.push(args);const a=args[args.indexOf('-a')+1];if(args[0]==='add-generic-password'){store.set(a,args[args.indexOf('-w')+1]);return '';}if(!store.has(a))throw Error('not found');return store.get(a);};
  const box=createSecretBox({platform:'darwin',keychain});const r=await box.protect('sk-mac-secret');assert.equal(r.protection,'keychain');assert.match(r.value,/^kc-/);assert.ok(!r.value.includes('sk-mac'));
  const again=createSecretBox({platform:'darwin',keychain});assert.equal(await again.unprotect(r.value,'keychain'),'sk-mac-secret');
  await assert.rejects(again.unprotect('kc-missing','keychain'),/钥匙串里找不到/);await assert.rejects(again.unprotect('abc','dpapi'),/Windows 上加密/);assert.equal(await again.unprotect('plain','none'),'plain');
});

test('Mac keychain round-trips Unicode across instances without confusing legacy hex-looking keys',async()=>{
 const {createSecretBox}=await import('../lib/secret.mjs');const store=new Map([['kc-legacy','deadbeef']]);
 const keychain=async args=>{const id=args[args.indexOf('-a')+1];if(args[0]==='add-generic-password'){store.set(id,args[args.indexOf('-w')+1]);return '';}const raw=store.get(id);if(raw===undefined)throw Error('missing');return /[^\x20-\x7e]/.test(raw)?Buffer.from(raw,'utf8').toString('hex'):raw;};
 for(const plain of ['sk-测试-123','秘密🔐\n末尾\n','deadbeef']){
  const sealed=await createSecretBox({platform:'darwin',keychain}).protect(plain);
  assert.equal(await createSecretBox({platform:'darwin',keychain}).unprotect(sealed.value,sealed.protection),plain);
 }
 assert.equal(await createSecretBox({platform:'darwin',keychain}).unprotect('kc-legacy','keychain'),'deadbeef');
});
