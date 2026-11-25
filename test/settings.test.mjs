import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {createSettingsStore,validateSettings} from '../lib/settings.mjs';
import {createCoachServer} from '../server.mjs';
test('settings migrate once, persist, reject invalid input and conflicting writes',async()=>{
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'coach-settings-'));try{
 const file=path.join(dir,'settings.json'),store=createSettingsStore(file);
 assert.equal((await store.get()).initialized,false);
 const migrated=await store.migrate({ids:['heading1','insertCard','heading1']});assert.deepEqual(migrated.settings.hiddenIds,['heading1']);
 assert.deepEqual((await store.migrate({ids:['toc']})).settings.hiddenIds,['heading1']);
 const changed=await store.update({text:{font:'微软雅黑',size:14,lineSpacing:2}},migrated.revision);
 await assert.rejects(store.update({hiddenIds:[]},migrated.revision),e=>e.status===409);
 assert.equal((await createSettingsStore(file).get()).settings.text.font,'微软雅黑');
 const before=await fs.readFile(file,'utf8');await assert.rejects(store.update({cards:{limit:1}},changed.revision));assert.equal(await fs.readFile(file,'utf8'),before);
 assert.throws(()=>validateSettings(JSON.parse('{"__proto__":{}}')));
 assert.throws(()=>validateSettings({backups:{directory:'relative'}}));
 await fs.writeFile(file,'broken');await assert.rejects(createSettingsStore(file).get(),/保留原文件/);assert.equal(await fs.readFile(file,'utf8'),'broken');
 }finally{await fs.rm(dir,{recursive:true,force:true});}
});
test('startup, settings, card import/export and backup use authenticated local APIs',async()=>{
 const dir=await fs.mkdtemp(path.join(os.tmpdir(),'coach-settings-api-'));let launched=0;
 const server=createCoachServer({cardsFile:path.join(dir,'cards.json'),settingsFile:path.join(dir,'settings.json'),pathsProvider:async()=>({wpsProgram:'mock-wps'}),launchProgram:async p=>{assert.equal(p,'mock-wps');launched++;}});
 await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));const base='http://127.0.0.1:'+server.address().port;
 try{const token=(await fetch(base+'/api/session').then(r=>r.json())).token;
 const req=(url,method='GET',body)=>fetch(base+url,{method,headers:{'Content-Type':'application/json','X-Coach-Token':token},body:body===undefined?undefined:JSON.stringify(body)});
 assert.equal((await fetch(base+'/api/settings')).status,403);
 const settings=await req('/api/settings').then(r=>r.json());
 const backupDir=path.join(dir,'custom-backups');assert.equal((await req('/api/settings','POST',{patch:{text:{font:'黑体'},backups:{directory:backupDir},cards:{limit:10}},expectedRevision:settings.revision})).status,200);
 assert.equal((await req('/api/settings','POST',{patch:{text:{size:14}},expectedRevision:settings.revision})).status,409);
 const card={id:'import-test',sourceHost:'wps',block:{type:'text',text:'导入内容'}};
 assert.deepEqual(await req('/api/cards/import','POST',{cards:[card]}).then(r=>r.json()),{added:1,skipped:0,total:1});
 assert.equal((await req('/api/cards/import','POST',{cards:[card]}).then(r=>r.json())).skipped,1);
 assert.equal((await req('/api/cards/import','POST',{cards:[{...card,id:'bad',block:{type:'text',text:''}}]})).status,422);
 const exported=await req('/api/cards/export').then(r=>r.json());assert.equal(exported.cards.length,1);
 const startup=await req('/api/startup').then(r=>r.json());assert.equal(startup.cardsCount,1);assert.equal(startup.backupDir,backupDir);assert.equal(startup.wpsFound,true);
 assert.equal((await req('/api/launch-wps','POST',{})).status,200);assert.equal(launched,1);
 const features=await req('/api/features').then(r=>r.json());assert.ok(features.features.find(f=>f.id==='bodyFormat'&&f.hosts[0]==='wps'));assert.deepEqual(features.features.find(f=>f.id==='deleteBlankPages').hosts,['wps']);assert.equal(features.features.some(f=>/^fav/.test(f.id)),false);
 const source=path.join(dir,'example.docx');await fs.writeFile(source,'fixture');const backup=await req('/api/backup','POST',{path:source}).then(r=>r.json());assert.equal(path.dirname(backup.backupPath),backupDir);assert.equal(await fs.readFile(backup.backupPath,'utf8'),'fixture');
 await req('/api/cards','DELETE');assert.equal((await req('/api/cards').then(r=>r.json())).cards.length,0);
 }finally{await new Promise(resolve=>server.close(resolve));await fs.rm(dir,{recursive:true,force:true});}
});
