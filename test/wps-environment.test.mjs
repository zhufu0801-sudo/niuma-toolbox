import http from 'node:http';
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import {classifyWps,registrationStatus,jsApiFlag,wpsEnvironment} from '../lib/windows.mjs';
import {createCoachServer,ribbonText} from '../server.mjs';

test('识别 WPS 的地区版本、安装范围和主版本号',()=>{
  const cn=classifyWps({program:'C:\\Users\\a\\AppData\\Local\\Kingsoft\\WPS Office\\12.1.0.18276\\office6\\wps.exe',company:'Zhuhai Kingsoft Office Software Co.,Ltd'});
  assert.equal(cn.version,'12.1.0.18276');assert.equal(cn.major,12);assert.equal(cn.region,'cn');assert.equal(cn.scope,'user');
  const pro=classifyWps({program:'C:\\Program Files (x86)\\Kingsoft\\WPS Office\\11.8.2.12085\\office6\\wps.exe',version:'11.8.2.12085',company:'北京金山办公软件股份有限公司'});
  assert.equal(pro.scope,'machine');assert.equal(pro.major,11);assert.equal(pro.region,'cn');
  assert.equal(classifyWps({program:'C:\\Program Files\\Kingsoft\\WPS Office\\11.2.0.11380\\office6\\wps.exe',company:'Kingsoft Corporation'}).region,'intl');
  assert.equal(classifyWps({program:'C:\\Program Files\\Kingsoft\\WPS Office\\11.3.0.1\\office6\\wps.exe',company:'Kingsoft Japan, Inc.'}).region,'jp');
  assert.equal(classifyWps({program:'',language:'日本語'}).region,'jp');
  // 实测：日本版 WPS Office 2（11.2.0.10707）装在 Program Files (x86)，公司名同中国版，靠注册表的默认语言和更新服务器区分。
  const jp=classifyWps({program:'C:\\Program Files (x86)\\Kingsoft\\WPS Office\\11.2.0.10707\\office6\\wps.exe',version:'11,2,0,10707',company:'Zhuhai Kingsoft Office Software Co.,Ltd',regLanguage:'1041',regUpdate:'https://yun.kingsoft.jp/updateserver/update'});
  assert.equal(jp.region,'jp');assert.equal(jp.version,'11.2.0.10707');assert.equal(jp.major,11);assert.equal(jp.scope,'machine');
  // 实测：WPS 12.1 的文件属性版本号是“12,1,0,29690”，不能把主版本识别成 0。
  const comma=classifyWps({program:'D:\\WPS\\office6\\wps.exe',version:'12,1,0,29690'});assert.equal(comma.version,'12.1.0.29690');assert.equal(comma.major,12);
  assert.equal(classifyWps({program:'C:\\Users\\a\\AppData\\Local\\Kingsoft\\WPS Office\\12.1.0.29690\\office6\\wps.exe',version:'12,1,0,29690'}).version,'12.1.0.29690');
  const old=classifyWps({program:'D:\\WPS\\10.1.0.7698\\office6\\wps.exe'});assert.equal(old.major,10);assert.equal(old.region,'unknown');
});

test('读取 publish.xml 的登记情况，区分文字、表格、演示',()=>{
  const xml='<jsplugins><jspluginonline name="WpsContentWorkbench" url="http://127.0.0.1:39127/addon/wps/"/><jspluginonline url="http://127.0.0.1:39127/addon/et/"/></jsplugins>';
  assert.deepEqual(registrationStatus(xml),{found:true,hosts:{wps:true,et:true,wpp:false}});
  assert.deepEqual(registrationStatus(undefined),{found:false,hosts:{wps:false,et:false,wpp:false}});
  assert.equal(registrationStatus('<x url="http://127.0.0.1:39128/addon/wps/"/>').hosts.wps,false,'端口不同不能算已登记');
});

test('oem.ini 的 JsApiPlugin 开关：支持 UTF-16 文件，没写时返回未知',()=>{
  assert.equal(jsApiFlag('[Setting]\nA=1\n[Support]\nJsApiPlugin=true\n'),true);
  assert.equal(jsApiFlag('[Support]\r\nJsApiShowWebDebugger=true\r\nJsApiPlugin=false\r\n[Server]\r\nJsApiPlugin=true'),false);
  assert.equal(jsApiFlag('[Setting]\nA=1'),null);
  assert.equal(jsApiFlag(Buffer.concat([Buffer.from([0xff,0xfe]),Buffer.from('[Support]\nJsApiPlugin=1\n','utf16le')])),true);
});

test('wpsEnvironment 汇总安装目录、注册表结果与登记文件，按版本从新到旧排列',async()=>{
  const local='C:\\Users\\a\\AppData\\Local',dir=name=>({name,isDirectory:()=>true});
  const files={
    'C:\\Users\\a\\AppData\\Local\\Kingsoft\\WPS Office\\12.1.0.18276\\office6\\cfgs\\oem.ini':'[Support]\nJsApiPlugin=true',
    'C:\\Users\\a\\AppData\\Roaming\\kingsoft\\wps\\jsaddons\\publish.xml':'<x url="http://127.0.0.1:39127/addon/wps/"/><x url="http://127.0.0.1:39127/addon/wpp/"/>'
  };
  const seen=[];
  const env=await wpsEnvironment({cache:false,platform:'win32',env:{LOCALAPPDATA:local,APPDATA:'C:\\Users\\a\\AppData\\Roaming',ProgramFiles:'C:\\Program Files'},
    readdir:async base=>{if(base.startsWith(local))return [dir('12.1.0.18276'),dir('templates')];if(base.startsWith('C:\\Program Files'))return [dir('11.8.2.12085')];throw Error('ENOENT');},
    readFile:async file=>{if(file in files)return files[file];throw Error('ENOENT');},
    run:async(cmd,args,opts)=>{seen.push(JSON.parse(opts.env.WPS_CANDIDATES));return {stdout:JSON.stringify([
      {program:'C:\\Program Files\\Kingsoft\\WPS Office\\11.8.2.12085\\office6\\wps.exe',version:'11.8.2.12085',company:'Zhuhai Kingsoft Office Software Co.,Ltd'},
      {program:'C:\\Users\\a\\AppData\\Local\\Kingsoft\\WPS Office\\12.1.0.18276\\office6\\wps.exe',version:'12.1.0.18276',company:'Zhuhai Kingsoft Office Software Co.,Ltd'}])};}});
  assert.equal(seen[0].length,2,'应把两个安装目录都交给版本查询');
  assert.deepEqual(env.installs.map(x=>x.version),['12.1.0.18276','11.8.2.12085']);
  assert.equal(env.installs[0].jsApiPlugin,true);assert.equal(env.installs[1].jsApiPlugin,null);
  assert.deepEqual(env.registration.hosts,{wps:true,et:false,wpp:true});
  const none=await wpsEnvironment({cache:false,platform:'linux',env:{},readdir:async()=>{throw Error('x');},readFile:async()=>{throw Error('x');},stat:async()=>{throw Error('x');}});
  assert.deepEqual(none.installs,[]);assert.equal(none.registration.found,false);
});

test('/api/environment 需要本机口令；/api/host-info 只接受三种应用并清理字段',async t=>{
  const server=createCoachServer({environmentProvider:async()=>({installs:[{version:'12.1.0.1',major:12,region:'cn',scope:'user'}],registration:{found:true,hosts:{wps:true,et:true,wpp:true},file:'x'}})});
  await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>server.close());
  const base=`http://127.0.0.1:${server.address().port}`,token=(await(await fetch(base+'/api/session')).json()).token,headers={'Content-Type':'application/json','X-Coach-Token':token};
  assert.equal((await fetch(base+'/api/environment')).status,403);
  assert.equal((await fetch(base+'/api/host-info',{method:'POST',headers,body:JSON.stringify({host:'word'})})).status>=400,true);
  const ok=await fetch(base+'/api/host-info',{method:'POST',headers,body:JSON.stringify({host:'et',version:'12.1.0.18276'.repeat(20),language:'1041',caps:{RibbonUI:true,'bad key':true,ShowDialog:'yes'}})});
  assert.equal(ok.status,200);
  const env=await(await fetch(base+'/api/environment',{headers})).json();
  assert.equal(env.installs[0].major,12);assert.equal(env.hosts.et.language,'1041');assert.ok(env.hosts.et.version.length<=80);
  assert.deepEqual(env.hosts.et.caps,{RibbonUI:true,ShowDialog:null});
});

test('功能区文字由 ribbon-text.js 提供，设置页的功能列表仍有名称',async t=>{
  const xml=fs.readFileSync(new URL('../public/ribbon.xml',import.meta.url),'utf8'),text=ribbonText(fs.readFileSync(new URL('../public/ribbon-text.js',import.meta.url),'utf8'));
  assert.ok(!/\blabel="/.test(xml),'ribbon.xml 不应再写死中文名称');
  for(const id of [...xml.matchAll(/\bid="([^"]+)"[^>]*getLabel=/g)].map(m=>m[1]))assert.ok(text[id]&&text[id].label,id+' 缺少名称');
  const server=createCoachServer();await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>server.close());
  const base=`http://127.0.0.1:${server.address().port}`,token=(await(await fetch(base+'/api/session')).json()).token;
  const features=(await(await fetch(base+'/api/features',{headers:{'X-Coach-Token':token}})).json()).features;
  assert.ok(features.length>10);assert.ok(features.every(f=>f.label));
});

function i18nWindow(lang){
  const store={};const window={navigator:{language:lang==='ja'?'ja-JP':'zh-CN',languages:[lang==='ja'?'ja-JP':'zh-CN']},localStorage:{getItem:k=>store[k]??null,setItem:(k,v)=>{store[k]=String(v);},removeItem:k=>{delete store[k];}},alert(){},confirm(){return true;},setTimeout,clearTimeout};
  window.window=window;
  // 页面按顺序加载 i18n-ja.js、i18n-ja-extra.js、i18n.js，测试也照此加载
  vm.runInNewContext(['i18n-ja.js','i18n-ja-extra.js','i18n.js'].map(f=>fs.readFileSync(new URL('../public/'+f,import.meta.url),'utf8')).join('\n'),window);
  return window.OfficeI18n;
}
test('日文：整句、片段和带用户内容的提示都能翻译，用户内容保持原样',()=>{
  const ja=i18nWindow('ja');assert.equal(ja.lang(),'ja');
  assert.equal(ja.t('把当前选区设为模块'),'選択範囲をモジュールにする');
  assert.equal(ja.t('已显示“一级标题”。'),'「一级标题」を表示しました。');
  assert.equal(ja.t('已存成卡片：'+ja.mark('销售目标 文字')),'カードとして保存しました：销售目标 文字');
  assert.equal(ja.t('3 行 × 2 列表格：'+ja.mark('区域、销售额')),'3 行 × 2 列の表：区域、销售额');
  const zh=i18nWindow('zh');assert.equal(zh.lang(),'zh');assert.equal(zh.t('把当前选区设为模块'),'把当前选区设为模块');assert.equal(zh.t('已存成卡片：'+zh.mark('A')),'已存成卡片：A');
});
test('日文字典覆盖界面里的固定文字',()=>{
  const ja=i18nWindow('ja'),cjk=/[\u4e00-\u9fff]/;
  const html=['tools.html','start.html','settings.html','studio.html'].map(f=>fs.readFileSync(new URL('../public/'+f,import.meta.url),'utf8')).join('\n').replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>|<!--[\s\S]*?-->/g,'').replace(/<[^>]*translate="no"[^>]*>[^<]*/g,'');
  const missing=[...new Set([...html.matchAll(/>([^<>]+)</g)].map(m=>m[1].trim()).filter(s=>cjk.test(s)))].filter(s=>ja.t(s)===s&&!['文字'].includes(s));
  assert.deepEqual(missing,[]);
});

test('自动更新留下的旧版本目录排在注册表指向的版本后面，并标成旧目录',async()=>{
  const local='C:\\Users\\a\\AppData\\Local',dir=name=>({name,isDirectory:()=>true});
  const at=v=>`C:\\Users\\a\\AppData\\Local\\Kingsoft\\WPS Office\\${v}\\office6\\wps.exe`;
  const env=await wpsEnvironment({cache:false,platform:'win32',env:{LOCALAPPDATA:local},
    readdir:async base=>base.startsWith(local)?[dir('12.1.0.25225'),dir('12.1.0.29161'),dir('12.1.0.29690')]:[],
    readFile:async()=>{throw Error('ENOENT');},
    run:async()=>({stdout:JSON.stringify([{program:at('12.1.0.29161'),version:'12,1,0,29161',active:true},{program:at('12.1.0.29690'),version:'12,1,0,29690',active:false},{program:at('12.1.0.25225'),version:'12,1,0,25225',active:false}])})});
  assert.deepEqual(env.installs.map(x=>[x.version,x.active]),[['12.1.0.29161',true],['12.1.0.29690',false],['12.1.0.25225',false]]);
});

test('来源为 null（日本版 WPS 侧栏）时：GET 和带口令的请求放行，不带口令的写操作拒绝',async t=>{
  const server=createCoachServer();await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>server.close());
  const port=server.address().port,base=`http://127.0.0.1:${port}`;
  const request=(path,opts={})=>new Promise((resolve,reject)=>{const req=http.request({host:'127.0.0.1',port,path,method:opts.method||'GET',headers:{Origin:'null',...(opts.headers||{})}},res=>{let b='';res.on('data',d=>b+=d);res.on('end',()=>resolve({status:res.statusCode,body:b}));});req.on('error',reject);if(opts.body)req.write(opts.body);req.end();});
  const session=await request('/api/session');assert.equal(session.status,200);const token=JSON.parse(session.body).token;
  assert.equal((await request('/api/settings',{headers:{'X-Coach-Token':token}})).status,200);
  assert.equal((await request('/api/host-info',{method:'POST',headers:{'Content-Type':'text/plain'},body:'{"host":"wps"}'})).status,403);
  assert.equal((await request('/api/host-info',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':token},body:'{"host":"wps"}'})).status,200);
  assert.equal((await request('/api/session',{headers:{Origin:'http://evil.example'}})).status,403);
});

test('导出诊断信息：保存到备份文件夹，不含卡片内容，用户目录替换成 ~',async t=>{
  const os=await import('node:os'),path=await import('node:path'),fsp=await import('node:fs/promises');
  const dir=await fsp.mkdtemp(path.join(os.tmpdir(),'diag-')),backups=path.join(dir,'backups');
  t.after(()=>fsp.rm(dir,{recursive:true,force:true}));
  const server=createCoachServer({cardsFile:path.join(dir,'cards.json'),settingsFile:path.join(dir,'settings.json'),backupDir:backups,
    environmentProvider:async()=>({installs:[{program:os.homedir()+'\\AppData\\Local\\Kingsoft\\WPS Office\\12.1.0.1\\office6\\wps.exe',version:'12.1.0.1',major:12,region:'cn'}],registration:{found:true,hosts:{wps:true,et:true,wpp:true}}})});
  await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>server.close());
  const base=`http://127.0.0.1:${server.address().port}`,token=(await(await fetch(base+'/api/session')).json()).token,headers={'Content-Type':'application/json','X-Coach-Token':token};
  await fetch(base+'/api/cards',{method:'POST',headers,body:JSON.stringify({block:{type:'text',text:'机密的卡片内容'},sourceHost:'wps'})});
  assert.equal((await fetch(base+'/api/diagnostics',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'})).status,403);
  const res=await fetch(base+'/api/diagnostics',{method:'POST',headers,body:'{}'});assert.equal(res.status,200);
  const {path:file}=await res.json();assert.ok(file.startsWith(backups));
  const text=await fsp.readFile(file,'utf8'),data=JSON.parse(text);
  assert.equal(data.cards.count,1);assert.ok(!text.includes('机密的卡片内容'));
  assert.equal(data.wps.installs[0].version,'12.1.0.1');assert.equal(data.wps.installs[0].program,undefined);assert.equal(data.version.split('.').length,3);
});

test('功能区图标按文档写法返回相对 SVG 路径，图标文件都存在，服务能按 PNG 类型返回',async t=>{
  const win={};vm.runInNewContext(fs.readFileSync(new URL('../public/tool-icons.js',import.meta.url),'utf8'),{window:win});
  for(const id of ['settingsOpen','studioOpen','toolsOpen','backupFile','exportPdf','heading1','saveCard']){const src=win.ToolIcons.forAction(id);assert.match(src,/^icons\/[a-z0-9]+\.svg$/);assert.ok(fs.existsSync(new URL('../public/'+src,import.meta.url)),src);}
  const server=createCoachServer();await new Promise(r=>server.listen(0,'127.0.0.1',r));t.after(()=>server.close());
  const port=server.address().port;
  const res=await new Promise((resolve,reject)=>{http.get({host:'127.0.0.1',port,path:'/addon/wpp/icons/large/toolbox.png'},r=>{r.resume();resolve(r);}).on('error',reject);});
  assert.equal(res.statusCode,200);assert.equal(res.headers['content-type'],'image/png');
});

test('系统代理：判断是否放行本机地址',async()=>{
  const {proxyBypassesLocal,proxySettings}=await import('../lib/windows.mjs');
  assert.equal(proxyBypassesLocal({enabled:false}),true);
  assert.equal(proxyBypassesLocal({enabled:true,override:'localhost;<local>'}),false);
  assert.equal(proxyBypassesLocal({enabled:true,override:'localhost;127.*;10.*'}),true);
  const info=await proxySettings({platform:'win32',cache:false,run:async()=>({stdout:'{"enabled":1,"server":"127.0.0.1:7890","override":"<local>","autoConfig":""}'})});
  assert.equal(info.enabled,true);assert.equal(info.server,'127.0.0.1:7890');assert.equal(info.bypassLocal,false);
});

test('功能区 getImage 只返回不带参数的相对路径（WPS 对带 ? 的地址、完整网址、data 地址都不取图）',()=>{
  const win={};vm.runInNewContext(fs.readFileSync(new URL('../public/tool-icons.js',import.meta.url),'utf8'),{window:win});
  const ctx={window:win,location:{origin:'http://127.0.0.1:39127',pathname:'/addon/wpp/index.html'},OfficeTools:{host:()=>'wpp',prefsStamp:()=>0,reportHost(){}},ToolIcons:win.ToolIcons,PaneBridge:{view:v=>v},StudioLaunch:{}};
  vm.runInNewContext(fs.readFileSync(new URL('../public/main.js',import.meta.url),'utf8'),ctx);
  win.OnAddinLoad({Invalidate(){}});
  const ids=[...fs.readFileSync(new URL('../public/ribbon.xml',import.meta.url),'utf8').matchAll(/<button[^>]*\bid="([^"]+)"[^>]*getImage=/g)].map(m=>m[1]);
  assert.ok(ids.length>20);
  for(const id of ids){const src=win.GetImage({Id:id});assert.match(src,/^icons\/[a-z0-9]+\.svg$/,id);}
});
