import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import {mergePublishXml,removePublishEntries,officeManifest,launchAgent,registerWps,registerOffice,uninstall,detect} from '../lib/mac-install.mjs';
const tmp=async t=>{const d=await fs.mkdtemp(path.join(os.tmpdir(),'macinst-'));t.after(()=>fs.rm(d,{recursive:true,force:true}));return d;};
test('publish.xml：换上办公工具箱三条，别的加载项保留；卸载只删自己的',()=>{
  const old='<?xml version="1.0"?>\n<jsplugins>\n  <jspluginonline name="Other" type="wps" url="http://x/"/>\n  <jspluginonline name="WpsCoachWriter" type="wps" url="http://old/"/>\n</jsplugins>\n';
  const next=mergePublishXml(old);assert.match(next,/name="Other"/);assert.equal((next.match(/WpsCoachWriter/g)||[]).length,1);assert.match(next,/url="http:\/\/127\.0\.0\.1:39127\/addon\/et\/"/);assert.ok(!next.includes('http://old/'));
  assert.equal(mergePublishXml(next),next);
  const back=removePublishEntries(next);assert.match(back,/name="Other"/);assert.ok(!/WpsCoach/.test(back));
  assert.match(mergePublishXml(''),/<jsplugins>[\s\S]*WpsCoachSlides[\s\S]*<\/jsplugins>/);assert.match(mergePublishXml('<jsplugins/>'),/WpsCoachSheets/);
  assert.throws(()=>mergePublishXml('<html></html>'),/格式和预期不同/);
});
test('Office 清单和登录启动项',()=>{
  const m=officeManifest('<Id>{{ID}}</Id><V>{{VERSION}}</V><S>{{BASE}}/office/taskpane.html</S><A>{{BASE}}</A>{{REQUIREMENTS}}{{RUNTIMES}}x{{RUNTIMES}}{{EXTENDED}}',{id:'abc',version:'0.14.0.0'});
  assert.equal(m,'<Id>abc</Id><V>0.14.0.0</V><S>https://localhost:39128/office/taskpane.html</S><A>https://localhost:39128</A>x');
  const p=launchAgent({node:'/opt/homebrew/bin/node',appDir:'/Users/a/Library/Application Support/OfficeToolbox/app'});
  assert.match(p,/<string>\/opt\/homebrew\/bin\/node<\/string><string>\/Users\/a\/Library\/Application Support\/OfficeToolbox\/app\/server\.mjs<\/string>/);assert.match(p,/RunAtLoad<\/key><true\/>/);
});
test('在模拟的 Mac 用户目录里登记 WPS、Office，再卸载',async t=>{
  const home=await tmp(t),app=path.join(home,'app');
  await fs.mkdir(path.join(home,'Library','Containers','com.kingsoft.wpsoffice.mac'),{recursive:true});
  await fs.mkdir(path.join(home,'Library','Containers','com.microsoft.Word','Data'),{recursive:true});
  await fs.mkdir(path.join(app,'public','office'),{recursive:true});await fs.writeFile(path.join(app,'package.json'),'{"version":"0.14.0"}');await fs.writeFile(path.join(app,'public','office','manifest-template.xml'),'<Id>{{ID}}</Id>{{BASE}}');
  const pub=path.join(home,'Library','Containers','com.kingsoft.wpsoffice.mac','Data','.kingsoft','wps','jsaddons','publish.xml');await fs.mkdir(path.dirname(pub),{recursive:true});await fs.writeFile(pub,'<jsplugins>\n<jspluginonline name="Other" url="x"/>\n</jsplugins>');
  const files=await registerWps(home);assert.equal(files.length,2);assert.match(await fs.readFile(pub,'utf8'),/Other[\s\S]*WpsCoachWriter/);assert.ok((await fs.readdir(path.dirname(pub))).some(n=>n.startsWith('publish.xml.toolbox-backup-')));
  const calls=[];const exec=async(cmd,args)=>{calls.push({cmd,args});if(cmd==='openssl'&&args[0]==='req'){await fs.writeFile(args[args.indexOf('-keyout')+1],'TEST KEY');await fs.writeFile(args[args.indexOf('-out')+1],'-----BEGIN CERTIFICATE-----\nTEST ONLY');}if(cmd==='openssl'&&args[0]==='x509')return 'sha1 Fingerprint=AA:BB';return '';};
  const written=await registerOffice(app,home,exec);assert.equal(written.length,1);assert.match(written[0].replaceAll(path.sep,'/'),/com\.microsoft\.Word\/Data\/Documents\/wef\/OfficeToolbox-.*\.xml$/);
  assert.match(await fs.readFile(path.join(app,'generated','office-cert.pem'),'utf8'),/BEGIN CERTIFICATE/);
  await uninstall(app,home,exec);assert.ok(calls.some(c=>c.cmd==='security'&&c.args[0]==='delete-certificate'));assert.ok(!/WpsCoach/.test(await fs.readFile(pub,'utf8')));assert.match(await fs.readFile(pub,'utf8'),/Other/);assert.deepEqual(await fs.readdir(path.dirname(written[0])),[]);
  assert.deepEqual(await detect(home),{wps:null,office:[]});
});
