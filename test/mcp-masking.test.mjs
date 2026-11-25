import test from 'node:test';
import assert from 'node:assert/strict';
import {createMcpMasker} from '../lib/mcp-masking.mjs';
import {createMaskVault,maskEngine} from '../lib/masking.mjs';

test('MCP codes follow UI language and remain restorable across language changes',async()=>{
 const vault=createMaskVault(null);await vault.savePrefs({types:['PHONE']});
 let lang='ja',doc='13812345678',last;
 const masker=createMcpMasker({vault,getLanguage:async()=>lang,execute:async x=>{last=x;return {text:x.action==='insertText'?x.args.text:doc};}});
 const call=(action,args={})=>masker.execute({sessionId:'language-session-0001',action,args});
 try {
  assert.equal((await call('readDocument')).text,'【電話1】');
  lang='zh';doc='13812345678 13912345678';
  assert.equal((await call('readDocument')).text,'【電話1】 【电话2】');
  await call('insertText',{text:'［ 電話1 ］ [电话2]'});
  assert.equal(last.args.text,'13812345678 13912345678');
  assert.equal((await masker.execute({sessionId:'new-zh-session-0001',action:'readDocument'})).text,'【电话1】 【电话2】');
 } finally {masker.close();}
});

test('MCP masks reads, stable isolated sessions, restores every write argument, expiry and switch',async()=>{
 const vault=createMaskVault(null);await vault.saveProfiles({profiles:[{name:'qa',words:'张三\n李四'}]});
 let time=0,doc='张三 13812345678',last;
 const mask=createMcpMasker({vault,now:()=>time,ttl:100,execute:async x=>{last=x;return x.action.startsWith('read')?{text:doc}:{text:x.args.text||'ok'};}});
 const call=(action,args={},sessionId='test-session-0001')=>mask.execute({clientId:'1234567890abcdef',sessionId,action,args});
 try {
  assert.equal((await vault.prefs()).agent,true);
  await assert.rejects(()=>mask.execute({action:'readDocument'}),/MCP/);
  const first=await call('readDocument');assert.equal(first.text,'【名称A】 【电话1】');assert.equal(first.masking.count,2);
  assert.ok(!JSON.stringify(first).includes('张三'));assert.ok(!JSON.stringify(first).includes('13812345678'));
  doc='李四 张三 13812345678';const next=await call('readSelection');assert.equal(next.text,'【名称B】 【名称A】 【电话1】');
  const written=await call('insertText',{text:'[ 名称A ] ［电话1］ 【电话999】'});
  assert.equal(last.args.text,'张三 13812345678 【电话999】');assert.deepEqual(written.masking.unknown,['【电话999】']);assert.ok(!written.text.includes('张三'));
  await call('runTool',{id:'qa',options:{a:'【名称B】',b:'【电话1】',c:'keep'}});assert.deepEqual(last.args.options,{a:'李四',b:'13812345678',c:'keep'});
  doc='李四';assert.equal((await call('readDocument',{},'test-session-0002')).text,'【名称A】');
  time=101;await call('insertText',{text:'【名称A】'});assert.equal(last.args.text,'【名称A】');
  await vault.savePrefs({agent:false});doc='张三 13812345678';assert.equal((await call('readDocument')).text,doc);
  await vault.savePrefs({agent:true,types:['PHONE']});assert.equal((await call('readDocument')).text,'张三 【电话1】');
 } finally {mask.close();}
});

test('code restoration tolerates brackets/spaces, newest wins, does not cascade',async()=>{
 const E=await maskEngine(),pairs=[{masked:'【名称A】',original:'张三'},{masked:'【名称A】',original:'旧名'},{masked:'【电话1】',original:'【名称A】'},{masked:'客户甲',original:'公司'}];
 const r=E.restoreCodes('【 名称 A 】 [名称A] ［名称A］ 【电话1】 客户甲 【名称Z】 [注释]',pairs);
 assert.equal(r.text,'张三 张三 张三 【名称A】 公司 【名称Z】 [注释]');assert.equal(r.done,5);assert.equal(r.unknown.join(','),'【名称Z】');
});
