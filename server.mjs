import {openDocument} from './lib/open-document.mjs';
import {saveNew} from './lib/print-tools.mjs';
import {createTemplateLibrary} from './lib/template-library.mjs';
import {createMcpConfig} from './lib/mcp-config.mjs';
import {createMcpBridge} from './lib/mcp-bridge.mjs';
import {createMcpMasker} from './lib/mcp-masking.mjs';
import {checkRelease} from './lib/release-info.mjs';
import {logSummary,hostSummary,environmentSummary,proxySummary,officeSummary} from './lib/diagnostic-safe.mjs';
import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { randomBytes, timingSafeEqual } from 'node:crypto';
import { MODEL, normalizeRequest, makeMessages, stripThinking } from './lib/coach.mjs';
import { importFile, preflight, exportFile, validateBlocks } from './lib/workspace.mjs';
import { planStudio, exportStudio, openStudioInWps, openStudioFile } from './lib/studio.mjs';
import { createAiStore, createAiClient, outlinePrompt, parseOutline } from './lib/ai.mjs';
import { createPrivacyGuard } from './lib/privacy.mjs';
import { translatePrompt, translateItemsPrompt, parseItems, cleanTranslation, formulaPrompt, parseFormula, writePrompt, cleanWrite, chatPrompt } from './lib/ai-tasks.mjs';
import { createCardStore } from './lib/cards.mjs';
import { createPasteBar, sensitiveReason } from './lib/pastebar.mjs';
import { inspectMerge, runMerge } from './lib/mail-merge.mjs';
import { mergePdfs, pickPages, splitPdf, imagesToPdf, pdfToImages, pdfInfo } from './lib/pdf-tools.mjs';
import { watchClipboard, readClipboardText } from './lib/clipboard-watch.mjs';
import { makeCodes } from './lib/codes.mjs';
import { stitchVertical } from './lib/images.mjs';
import { planRename, applyRename, undoRename } from './lib/rename-files.mjs';
import { convertParagraphs as jaStyleConvert } from './lib/ja-style.mjs';
import { createMaskVault, exportRecordCsv, scanModules, maskModules, restoreModules } from './lib/masking.mjs';
import { createPostalStore } from './lib/postal-jp.mjs';
import {createSettingsStore,outputDirectory} from './lib/settings.mjs';
import {systemPaths,launchWps,wpsEnvironment,proxySettings,officeEnvironment} from './lib/windows.mjs';
import {saveModuleToDesktop} from './lib/module-export.mjs';
import {startOfficeServer} from './lib/office-server.mjs';
import {createDocTools} from './lib/doc-tools.mjs';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
// 功能区文字放在 public/ribbon-text.js（浏览器脚本），这里取出其中的 JSON。
export function ribbonText(source){const start=source.indexOf('{'),end=source.lastIndexOf('}');try{return JSON.parse(source.slice(start,end+1));}catch{return {};}}
const PUBLIC = path.join(ROOT, 'public');
const VERSION = JSON.parse(await fs.readFile(path.join(path.dirname(fileURLToPath(import.meta.url)), 'package.json'), 'utf8')).version;
const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png', '.xml': 'application/xml; charset=utf-8', '.json': 'application/json; charset=utf-8' };
// 诊断信息：只收集版本、WPS 识别结果、加载状态和错误记录，不含卡片内容和文档内容；用户目录替换成 ~。
export async function collectDiagnostics({version=VERSION,settingsState,cardsCount,environment,proxy=null,hosts,office=null,recentFiles=[],logFile=path.join(ROOT,'service-error.log'),home=os.homedir()}={}){
  const hide=value=>JSON.parse(JSON.stringify(value??null).split(JSON.stringify(home).slice(1,-1)).join('~'));
  const errorLog=await logSummary(logFile),serviceLog=await logSummary(path.join(path.dirname(logFile),'service.log'));
  const s=settingsState?.settings||{};
  return hide({product:'办公工具箱',version,generatedAt:new Date().toISOString(),
    system:{platform:process.platform,release:os.release(),arch:process.arch,node:process.version,locale:Intl.DateTimeFormat().resolvedOptions().locale,timeZone:Intl.DateTimeFormat().resolvedOptions().timeZone},
    settings:{language:s.ui?.language,hiddenCount:(s.hiddenIds||[]).length,cardLimit:s.cards?.limit,backupDirectoryConfigured:!!s.backups?.directory,autoOpenWps:!!s.launch?.autoOpenWps},
    cards:{count:cardsCount},wps:environmentSummary(environment),office:officeSummary(office),proxy:proxySummary(proxy),addins:hostSummary(hosts),recentFiles:recentFiles.map(x=>({at:x.at,status:x.status})),serviceErrors:errorLog,serviceLog});
}
export function createCoachServer({ ollamaUrl = 'http://127.0.0.1:11434', model = MODEL, fetchImpl = fetch, aiEnabled = false, studioDir=null, openPresentation=openStudioFile, cardsFile=null, pasteBarFile=null, clipboardWatcher=watchClipboard, clipboardReader=readClipboardText, backupDir=path.join(os.homedir(),'Documents','办公工具箱备份'),outputHome=os.homedir(),settingsFile=null,pathsProvider=systemPaths,launchProgram=launchWps,environmentProvider=wpsEnvironment,proxyProvider=proxySettings,officeProvider=officeEnvironment,aiFile=path.join(ROOT,'generated','ai.json'),aiFetch=fetch,aiSecrets=null,privacyPluginDir=path.join(ROOT,'plugins'),docToolsOptions={},maskFile=path.join(ROOT,'generated','masking.json'),maskSecrets=null,postalFile=path.join(ROOT,'generated','postal-jp.json.gz'),postalOptions={} } = {}) {
  // 客户自己的 AI 服务（OpenAI 兼容接口）；密钥只在本机文件里，接口只返回打码后的提示。
  // 数据脱敏的词表和对照表（2026-11-25）。词表设了“发给 AI 前也隐藏”时，AI 的隐私保护一起用。
  const maskVault=createMaskVault(maskFile,maskSecrets?{secrets:maskSecrets}:{});
  // 日本邮编（2026-11-25）：安装包里带一份；“更新数据”下载的存在 generated/，用较新的那份。
  const postal=createPostalStore({userFile:postalFile,...postalOptions});
  const aiStore=createAiStore(aiFile,aiSecrets?{secrets:aiSecrets}:{}),guard=createPrivacyGuard({getMode:()=>aiStore.mode(),getWords:async()=>{const p=await maskVault.prefs();return p.ai?await maskVault.currentWords():'';},pluginDir:privacyPluginDir}),ai=createAiClient({store:aiStore,fetchImpl:aiFetch,guard});let aiBusy=false;
  // 各 WPS 窗口里的加载项启动时上报版本和可用接口，启动页据此显示当前 WPS 是否兼容。
  const hostInfo={};
  const settings=createSettingsStore(settingsFile,{backupDir});
  // 文件与图片工具（图片转文字、PDF 转 Word、提取 / 压缩图片、历史版本）；上传的文件没有位置时，结果放到备份文件夹。
  const docTools=createDocTools({fallbackDir:async()=>outputDirectory((await settings.get()).settings,{home:outputHome}),...docToolsOptions});let docBusy=false;const autoSeen=new Map();
  const cards = createCardStore(cardsFile,{getLimit:async()=>(await settings.get()).settings.cards.limit});
  // 粘贴栏：自动记录复制内容默认关闭，只在设置里打开后才启动剪贴板监听；关掉或服务退出时停止。
  const pasteBar=createPasteBar(pasteBarFile);let clipWatch=null,clipSkipped=0;
  function syncClipboardWatch(s){const want=!!(s&&s.pasteBar&&s.pasteBar.watch);if(want&&!clipWatch){clipWatch=clipboardWatcher?clipboardWatcher(text=>{if(sensitiveReason(text)){clipSkipped++;return;}pasteBar.add(text,'clipboard','clipboard').then(item=>autoMask('paste',item)).catch(()=>{});}):null;}else if(!want&&clipWatch){clipWatch.stop();clipWatch=null;}}
  settings.get().then(s=>syncClipboardWatch(s.settings),()=>{});
  const token = randomBytes(24).toString('hex');
  const mcp=createMcpBridge(),mcpConfig=createMcpConfig({root:ROOT});
  const mcpMasker=createMcpMasker({vault:maskVault,execute:input=>mcp.execute(input),getLanguage:async()=>{
    const lang=(await settings.get()).settings.ui.language;
    return lang==='auto'?(/^ja/i.test(Intl.DateTimeFormat().resolvedOptions().locale)?'ja':'zh'):lang;
  }});
  const outputDir=async()=>outputDirectory((await settings.get()).settings,{home:outputHome});
  // 卡片 / 粘贴栏“存入时自动脱敏”：开着时，新存进去的内容马上按当前词表和方式脱敏，原文进对照表。
  async function autoMask(kind,item){if(!item)return item;try{const p=await maskVault.prefs();if(!p.autoModules)return item;const lang=(await settings.get()).settings.ui.language==='ja'?'ja':'zh';
    const r=await maskModules({cards,pasteBar,vault:maskVault,items:[{kind,id:item.id}],mode:p.mode,lang,types:p.types,words:await maskVault.currentWords(),auto:true});if(!r.changed)return item;
    const now=(await (kind==='card'?cards:pasteBar).list()).find(x=>x.id===item.id);return now?{...now,masked:r.total}:item;}catch{return item;}}
  const templates=createTemplateLibrary({root:ROOT,outputDir,openFile:openDocument});
  const sourceSeen=new Map();
  const visibleCard=card=>({...card,sourceCheckedAt:card.source?(sourceSeen.get(card.id)||0):undefined});
  let busy = false;
  let transfer = null;
  const json = (res, code, body) => { res.writeHead(code, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' }); res.end(JSON.stringify(body)); };
  // 最近的静态文件请求（只记路径和结果，不记内容），写进诊断信息，用来排查功能区图标显示成问号这类问题。
  const recentFiles=[];
  let startupRelease=Promise.resolve({status:'unset'});
  const server=http.createServer(async (req, res) => {
    const expectedHost = `127.0.0.1:${req.socket.localPort}`;
    if(!String(req.url||'').startsWith('/api/')){const entry={at:Date.now(),method:req.method,path:String(req.url||'').split('?')[0].slice(0,120),query:String(req.url||'').includes('?'),host:String(req.headers.host||'').slice(0,40),status:0};recentFiles.push(entry);if(recentFiles.length>40)recentFiles.shift();res.on('finish',()=>{entry.status=res.statusCode;});}
    if (req.headers.host !== expectedHost) return json(res, 403, { error: '只接受本机连接。' });
    const origin = req.headers.origin;
    // 实测日本版 WPS（11.2，WPS 2019 内核）的侧栏发请求时 Origin 是 "null"。只放行两种情况：读取（GET，响应不带跨域许可，别的网页读不到内容）
    // 和带本机口令头的请求（自定义请求头会触发跨域预检，本服务不放行预检，别的网页发不出来）。
    const nullOk = origin === 'null' && (req.method === 'GET' || Boolean(req.headers['x-coach-token']));
    if (origin && origin !== `http://${expectedHost}` && !nullOk) return json(res, 403, { error: '请求来源不匹配。' });
    const url = new URL(req.url, `http://${expectedHost}`);
    try {
      if(url.pathname==='/api/library'){
        const supplied=Buffer.from(String(req.headers['x-coach-token']||'')),expected=Buffer.from(token);if(supplied.length!==expected.length||!timingSafeEqual(supplied,expected))return json(res,403,{error:'本机连接校验失败。'});
        if(req.method!=='POST')return json(res,405,{error:'只接受 POST。'});let size=0,chunks=[];for await(const c of req){size+=c.length;if(size>100000)return json(res,413,{error:'内容过大。'});chunks.push(c);}try{const input=JSON.parse(Buffer.concat(chunks).toString('utf8'));if(input.action==='list')return json(res,200,{items:await templates.list()});if(docBusy)return json(res,409,{error:'上一个文件还在处理，请稍等。'});docBusy=true;try{return json(res,200,await templates.run(input));}finally{docBusy=false;}}catch(e){return json(res,422,{error:e.message});}
      }
      if(url.pathname.startsWith('/api/mcp/')){
        const supplied=Buffer.from(String(req.headers['x-coach-token']||'')),expected=Buffer.from(token);
        if(supplied.length!==expected.length||!timingSafeEqual(supplied,expected))return json(res,403,{error:'本机连接校验失败。'});
        if(req.method!=='POST')return json(res,405,{error:'只接受 POST。'});
        let size=0,chunks=[];for await(const chunk of req){size+=chunk.length;if(size>300000)return json(res,413,{error:'请求过大。'});chunks.push(chunk);}
        try{const input=JSON.parse(Buffer.concat(chunks).toString('utf8')||'{}'),action=url.pathname.slice('/api/mcp/'.length);
          if(action==='config-preview')return json(res,200,await mcpConfig.preview(input.client));
          if(action==='config-apply')return json(res,200,await mcpConfig.apply(input.previewId));
          if(action==='status')return json(res,200,mcp.status());
          if(action==='register')return json(res,200,mcp.register(input));
          if(action==='poll')return json(res,200,mcp.poll(input));
          if(action==='result')return json(res,200,mcp.finish(input));
          if(action==='execute')return json(res,200,await mcpMasker.execute(input));
          return json(res,404,{error:'没有这个操作。'});
        }catch(e){return json(res,e.status||422,{error:e.message});}
      }
      const linked=url.pathname.match(/^\/api\/modules\/([A-Za-z0-9_-]{1,80})\/(refresh|links)(?:\/([a-f0-9]{20}))?$/);
      if(linked){
        const supplied=Buffer.from(String(req.headers['x-coach-token']||'')),expected=Buffer.from(token);if(supplied.length!==expected.length||!timingSafeEqual(supplied,expected))return json(res,403,{error:'本机连接已过期，请重新打开工具箱。'});
        try{
          if(req.method==='DELETE'&&linked[2]==='links'&&linked[3]){await cards.detach(linked[1],linked[3]);return json(res,200,{ok:true});}
          if(req.method!=='POST')return json(res,405,{error:'不支持此操作。'});
          if(!/^application\/json\b/i.test(req.headers['content-type']||''))return json(res,415,{error:'需要 JSON 请求。'});
          const parts=[];let size=0;for await(const part of req){size+=part.length;if(size>800000)return json(res,413,{error:'联动内容过长，请缩小模块。'});parts.push(part);}let body;try{body=JSON.parse(Buffer.concat(parts).toString('utf8'));}catch{return json(res,400,{error:'联动请求格式不正确。'});}if(!body||typeof body!=='object'||Array.isArray(body))return json(res,400,{error:'联动请求格式不正确。'});
          if(linked[2]==='refresh'){const card=await cards.refresh(linked[1],body);sourceSeen.set(card.id,Date.now());return json(res,200,{card:visibleCard(card)});}
          const link=linked[3]?await cards.acknowledge(linked[1],linked[3],body):await cards.attach(linked[1],body);return json(res,200,{link});
        }catch(e){return json(res,e.status||422,{error:e.message});}
      }
      const desktopModule=url.pathname.match(/^\/api\/modules\/([A-Za-z0-9_-]{1,80})\/desktop$/);
      if(desktopModule){
        const supplied=Buffer.from(String(req.headers['x-coach-token']||'')),expected=Buffer.from(token);if(supplied.length!==expected.length||!timingSafeEqual(supplied,expected))return json(res,403,{error:'本机连接已过期，请重新打开页面。'});
        if(req.method!=='POST')return json(res,405,{error:'请点击“保存到桌面”。'});
        try{const card=(await cards.list()).find(c=>c.id===desktopModule[1]);if(!card)return json(res,404,{error:'这个模块已不存在，请刷新模块栏。'});const preferences=(await settings.get()).settings,paths=await pathsProvider();return json(res,200,await saveModuleToDesktop(card,preferences.export[card.sourceHost]||'docx',paths.desktop));}catch(e){return json(res,422,{error:e.message});}
      }
      if(['/api/settings','/api/settings/migrate','/api/startup','/api/features','/api/launch-wps','/api/cards/export','/api/cards/import','/api/environment','/api/host-info','/api/diagnostics','/api/updates'].includes(url.pathname)){
        const supplied=Buffer.from(String(req.headers['x-coach-token']||'')),expected=Buffer.from(token);if(supplied.length!==expected.length||!timingSafeEqual(supplied,expected))return json(res,403,{error:'本机连接已过期，请重新打开页面。'});
        if(req.method==='GET'){
          if(url.pathname==='/api/settings')return json(res,200,await settings.get());
          if(url.pathname==='/api/environment'){let env={installs:[],registration:{found:false,hosts:{wps:false,et:false,wpp:false},file:''}};try{env=await environmentProvider();}catch{}let proxy=null,office=null;try{proxy=await proxyProvider();}catch{}try{office=await officeProvider();}catch{}return json(res,200,{...env,proxy,office,hosts:hostInfo,office:await officeProvider().catch(()=>null),recentFiles});}
          if(url.pathname==='/api/cards/export')return json(res,200,{version:1,cards:(await cards.list()).map(({source,targets,revision,...portable})=>portable)});
          if(url.pathname==='/api/startup'){const state=await settings.get(),paths=await pathsProvider();return json(res,200,{version:VERSION,wpsFound:!!paths.wpsProgram,cardsCount:(await cards.list()).length,backupDir:state.settings.backups.directory});}
          if(url.pathname==='/api/features'){const xml=await fs.readFile(path.join(PUBLIC,'ribbon.xml'),'utf8'),text=ribbonText(await fs.readFile(path.join(PUBLIC,'ribbon-text.js'),'utf8')),features=[];for(const group of xml.matchAll(/<group\b([^>]*)>([\s\S]*?)<\/group>/g)){const groupId=group[1].match(/\bid="([^"]+)"/)?.[1];if(groupId==='favGroup')continue;const hosts=(groupId==='wordGroup'||groupId==='docGroup')?['wps']:(groupId==='sheetGroup'||groupId==='sheetPlusGroup')?['et']:groupId==='slideGroup'?['wpp']:['wps','et','wpp'];for(const match of group[2].matchAll(/<button\b([^>]*)\/?\s*>/g)){const attrs=match[1];if(!attrs.includes('getVisible="GetVisible"'))continue;const id=attrs.match(/\bid="([^"]+)"/)?.[1],label=attrs.match(/\blabel="([^"]+)"/)?.[1]||text[id]?.label;if(/^c[wep]_/.test(id||''))continue;if(id&&label)features.push({id,label,hosts});}}return json(res,200,{features});}
          return json(res,405,{error:'不支持此操作。'});
        }
        if(req.method!=='POST')return json(res,405,{error:'不支持此操作。'});
        if(url.pathname==='/api/updates'){const state=await settings.get();return json(res,200,await checkRelease(state.settings.updates.url,VERSION));}
        if(url.pathname==='/api/diagnostics'){
          let environment=null;try{environment=await environmentProvider();}catch(e){environment={error:e.message};}
          const state=await settings.get(),proxy=await proxyProvider().catch(()=>null),office=await officeProvider().catch(()=>null),data=await collectDiagnostics({settingsState:state,cardsCount:(await cards.list()).length,environment,proxy,office,hosts:hostInfo,recentFiles});
          const dir=state.settings.backups.directory,stamp=new Date(),pad=n=>String(n).padStart(2,'0');
          const file=path.join(dir,'诊断信息-'+stamp.getFullYear()+pad(stamp.getMonth()+1)+pad(stamp.getDate())+'-'+pad(stamp.getHours())+pad(stamp.getMinutes())+pad(stamp.getSeconds())+'-'+randomBytes(3).toString('hex')+'.json');
          try{await fs.mkdir(dir,{recursive:true});await fs.writeFile(file,JSON.stringify(data,null,2),{encoding:'utf8',flag:'wx'});}catch(e){return json(res,500,{error:'诊断信息没能保存：'+e.message});}
          return json(res,200,{path:file});
        }
        if(req.method!=='POST')return json(res,405,{error:'不支持此操作。'});
        if(!/^application\/json\b/i.test(req.headers['content-type']||''))return json(res,415,{error:'需要 JSON 请求。'});
        const chunks=[];let size=0;for await(const part of req){size+=part.length;if(size>(url.pathname==='/api/cards/import'?8*1024*1024:30000))return json(res,413,{error:'提交内容过长。'});chunks.push(part);}let input;try{input=JSON.parse(Buffer.concat(chunks).toString('utf8'));}catch{return json(res,400,{error:'提交内容格式不正确。'});}
        try{
          if(url.pathname==='/api/settings'){if(input.patch?.cards?.limit<(await cards.list()).length)throw Error('上限不能少于当前已有卡片数量。请先导出或手动删除不需要的卡片。');const saved=await settings.update(input.patch,input.expectedRevision);syncClipboardWatch(saved.settings);return json(res,200,saved);}
          if(url.pathname==='/api/settings/migrate')return json(res,200,await settings.migrate(input));
          if(url.pathname==='/api/cards/import')return json(res,200,await cards.merge(input.cards));
          if(url.pathname==='/api/host-info'){const host=input.host;if(!['wps','et','wpp'].includes(host))throw Error('来源应用无效。');const text=v=>typeof v==='string'?v.slice(0,80):'';const caps={};for(const [k,v] of Object.entries(input.caps||{}).slice(0,20))if(/^[a-zA-Z]{1,30}$/.test(k))caps[k]=v===true?true:v===false?false:null;const page=['main','pane','settings'].includes(input.page)?input.page:'main',prev=hostInfo[host]||{},pages={...(prev.pages||{})};pages[page]={lang:text(input.ui&&input.ui.lang),pref:text(input.ui&&input.ui.pref),probe:typeof (input.ui&&input.ui.probe)==='string'?input.ui.probe.slice(0,200):'',errors:(Array.isArray(input.errors)?input.errors:[]).slice(-5).map(e=>String(e).slice(0,200)),at:Date.now()};hostInfo[host]={version:text(input.version)||prev.version||'',build:text(input.build)||prev.build||'',name:text(input.name)||prev.name||'',language:text(input.language)||prev.language||'',caps:Object.keys(caps).length?caps:prev.caps||{},pages,at:Date.now()};return json(res,200,{ok:true});}
          if(url.pathname==='/api/launch-wps'){const paths=await pathsProvider();await launchProgram(paths.wpsProgram);return json(res,200,{ok:true});}
          return json(res,405,{error:'不支持此操作。'});
        }catch(e){return json(res,e.status||422,{error:e.message});}
      }
      if(req.method==='GET'&&url.pathname==='/api/ai/config'){const v=await aiStore.publicView(),i=await guard.info();return json(res,200,{...v,privacy:{mode:i.mode,apiVersion:i.apiVersion,plugins:i.plugins,failed:i.failed.map(f=>({file:f.file,error:f.error}))}});}
      if(req.method==='POST'&&['/api/ai/config','/api/ai/test','/api/ai/models','/api/ai/ppt-outline','/api/ai/translate','/api/ai/formula','/api/ai/write','/api/ai/chat','/api/ai/plugins/reload'].includes(url.pathname)){
        const supplied=Buffer.from(String(req.headers['x-coach-token']||'')),expected=Buffer.from(token);
        if(supplied.length!==expected.length||!timingSafeEqual(supplied,expected))return json(res,403,{error:'本机连接已过期，请刷新页面后重试。'});
        if(!/^application\/json\b/i.test(req.headers['content-type']||''))return json(res,415,{error:'需要 JSON 请求。'});
        let chunks=[],size=0;const cap=url.pathname==='/api/ai/chat'?160000:80000;for await(const chunk of req){size+=chunk.length;if(size>cap)return json(res,413,{error:'内容过长，请缩短后重试。'});chunks.push(chunk);}
        let input;try{input=JSON.parse(Buffer.concat(chunks).toString('utf8')||'{}');}catch{return json(res,400,{error:'请求格式不对。'});}
        try{
          if(url.pathname==='/api/ai/config')return json(res,200,await aiStore.save(input));
          if(url.pathname==='/api/ai/plugins/reload'){await guard.reload();const i=await guard.info();return json(res,200,{plugins:i.plugins,failed:i.failed});}
          if(aiBusy)return json(res,409,{error:'AI 正在处理上一个请求，请稍等。'});
          aiBusy=true;
          try{
            if(url.pathname==='/api/ai/test')return json(res,200,await ai.test());
            if(url.pathname==='/api/ai/models')return json(res,200,{models:await ai.models()});
            const lang=input.lang==='ja'?'ja':'zh',confirmed=!!input.privacyConfirmed;
            if(url.pathname==='/api/ai/translate'){
              if(Array.isArray(input.items)){
                const items=input.items.map(x=>String(x??''));if(!items.length||items.length>200||items.join('').length>8000||items.some(x=>x.length>2000))throw Object.assign(Error('一次最多处理 200 个单元格、共 8000 字以内。'),{status:400});
                const r=await ai.chat(translateItemsPrompt({items,action:String(input.action||''),lang}),{maxTokens:8000,feature:'translate',confirmed});
                return json(res,200,{items:parseItems(r.answer,items.length),model:r.model,provider:r.provider,privacy:r.privacy});
              }
              const text=String(input.text||'');if(!text.trim()||text.length>8000)throw Object.assign(Error('请先选中要处理的文字（8000 字以内）。'),{status:400});
              const r=await ai.chat(translatePrompt({text,action:String(input.action||''),lang}),{maxTokens:8000,feature:'translate',confirmed});
              return json(res,200,{text:cleanTranslation(r.answer,text),model:r.model,provider:r.provider,privacy:r.privacy});
            }
            if(url.pathname==='/api/ai/write'){
              const text=String(input.text||''),action=String(input.action||'');if(!text.trim()||text.length>8000)throw Object.assign(Error('请先选中要处理的文字（8000 字以内）。'),{status:400});
              const r=await ai.chat(writePrompt({text,action,lang,extra:String(input.extra||'')}),{maxTokens:8000,feature:'write',confirmed});
              return json(res,200,{text:cleanWrite(r.answer),model:r.model,provider:r.provider,privacy:r.privacy});
            }
            if(url.pathname==='/api/ai/chat'){
              const r=await ai.chat(chatPrompt({messages:input.messages,lang,context:String(input.context||'')}),{maxTokens:8000,feature:'chat',confirmed});
              return json(res,200,{text:r.answer.trim(),model:r.model,provider:r.provider,privacy:r.privacy});
            }
            if(url.pathname==='/api/ai/formula'){
              const intent=String(input.intent||'').trim();if(!intent||intent.length>500)throw Object.assign(Error('请用一句话写下想算什么（500 字以内）。'),{status:400});
              const r=await ai.chat(formulaPrompt({intent,context:input.context||{},lang}),{maxTokens:8000,feature:'formula',confirmed});
              return json(res,200,{...parseFormula(r.answer),model:r.model,provider:r.provider,privacy:r.privacy});
            }
            const topic=String(input.topic||'').trim();
            if(!topic||topic.length>2000)throw Object.assign(Error('请写下演示主题（2000 字以内）。'),{status:400});
            const r=await ai.chat(outlinePrompt({topic,audience:String(input.audience||'').slice(0,200),pages:input.pages,lang,notes:String(input.notes||'').slice(0,1000)}),{maxTokens:8000,feature:'ppt-outline',confirmed});
            const outline=parseOutline(r.answer,lang);planStudio({...outline,theme:'midnight',lang});
            return json(res,200,{...outline,model:r.model,provider:r.provider,privacy:r.privacy});
          }finally{aiBusy=false;}
        }catch(e){return json(res,e.status||422,{error:e.message,needConfirm:e.needConfirm||undefined,findings:e.findings||undefined});}
      }
      if(req.method==='POST'&&['/api/studio/preview','/api/studio/export','/api/studio/open'].includes(url.pathname)){
        const supplied=Buffer.from(String(req.headers['x-coach-token']||'')),expected=Buffer.from(token);
        if(supplied.length!==expected.length||!timingSafeEqual(supplied,expected))return json(res,403,{error:'本机连接已过期，请重新尝试。'});
        if(!/^application\/json\b/i.test(req.headers['content-type']||''))return json(res,415,{error:'需要 JSON 请求。'});
        let chunks=[],size=0;for await(const chunk of req){size+=chunk.length;if(size>14*1024*1024)return json(res,413,{error:"大纲或图片过大，请压缩图片或分批制作。"});chunks.push(chunk);}
        try{
          const input=JSON.parse(Buffer.concat(chunks).toString('utf8'));
          if(url.pathname.endsWith('/preview'))return json(res,200,planStudio(input));
          const {buffer,plan}=await exportStudio(input);
          if(url.pathname.endsWith('/open')){
            // 用演示标题做文件名，放在“生成文件的文件夹”里（以前放在安装目录 generated\studio、名字是一串数字，客户找不到）。
            const filename=await saveNew(studioDir||await outputDir(),(String(plan.title||'').replace(/[\\/:*?"<>|\x00-\x1f]/g,'').trim().slice(0,40)||'演示文稿')+'.pptx',buffer);
            try{await openPresentation(filename);}catch(e){return json(res,502,{error:e.message+' 文件已保存在：'+filename,filePath:filename});}
            return json(res,200,{opened:true,filePath:filename,pages:plan.pages.length});
          }
          res.writeHead(200,{'Content-Type':'application/vnd.openxmlformats-officedocument.presentationml.presentation','Content-Disposition':`attachment; filename="presentation.pptx"; filename*=UTF-8''${encodeURIComponent(plan.title)}.pptx`,'Cache-Control':'no-store'});return res.end(buffer);
        }catch(e){return json(res,422,{error:'未能制作：'+e.message});}
      }
      // Microsoft Office 版导出 PDF：侧栏用 Office 接口生成 PDF 内容后交给本服务，保存到备份文件夹（新文件名，不覆盖任何文件）。
      if (url.pathname === '/api/save-pdf' && req.method === 'POST') {
        const supplied=Buffer.from(String(req.headers['x-coach-token']||'')),expected=Buffer.from(token);
        if(supplied.length!==expected.length||!timingSafeEqual(supplied,expected))return json(res,403,{error:'本机连接校验失败，请重新打开工具箱。'});
        if(!/^application\/json\b/i.test(req.headers['content-type']||''))return json(res,415,{error:'需要 JSON 请求。'});
        let chunks=[],size=0;for await(const chunk of req){size+=chunk.length;if(size>110*1024*1024)return json(res,413,{error:'PDF 超过 80 MB，请在 Office 里用「另存为 PDF」。'});chunks.push(chunk);}
        try{
          const input=JSON.parse(Buffer.concat(chunks).toString('utf8')),data=Buffer.from(String(input.data||''),'base64');
          if(data.length<5||data.subarray(0,5).toString('latin1')!=='%PDF-')throw new Error('生成的内容不是 PDF，没有保存。');
          const base=String(input.name||'文档').replace(/\.[A-Za-z0-9]{1,5}$/,'').replace(/[\\/:*?"<>|\x00-\x1f]/g,'_').trim().slice(0,80)||'文档';
          const dir=(await settings.get()).settings.backups.directory;await fs.mkdir(dir,{recursive:true});
          const stamp=new Date(),pad=n=>String(n).padStart(2,'0'),name=base+'-'+stamp.getFullYear()+pad(stamp.getMonth()+1)+pad(stamp.getDate())+'-'+pad(stamp.getHours())+pad(stamp.getMinutes())+pad(stamp.getSeconds());
          let target=path.join(dir,name+'.pdf'),n=1;
          for(;;){try{await fs.writeFile(target,data,{flag:'wx'});break;}catch(e){if(e.code!=='EEXIST'||n>50)throw e;target=path.join(dir,name+'-'+(++n)+'.pdf');}}
          return json(res,200,{path:target});
        }catch(e){return json(res,422,{error:e.message});}
      }
      if (url.pathname.startsWith('/api/tools/') && req.method === 'POST') {
        const supplied=Buffer.from(String(req.headers['x-coach-token']||'')),expected=Buffer.from(token);
        if(supplied.length!==expected.length||!timingSafeEqual(supplied,expected))return json(res,403,{error:'本机连接校验失败，请重新打开工具箱。'});
        if(!/^application\/json\b/i.test(req.headers['content-type']||''))return json(res,415,{error:'需要 JSON 请求。'});
        const name=url.pathname.slice('/api/tools/'.length);if(!['ocr','pdf-to-word','extract-images','compress-images','history','reveal','auto-backup','exists','mail-merge','pdf','codes','stitch','rename','ja-style','mask','postal'].includes(name))return json(res,404,{error:'没有这个功能。'});
        let chunks=[],size=0;for await(const chunk of req){size+=chunk.length;if(size>120*1024*1024)return json(res,413,{error:'文件太大了（上传最多约 80 MB）。'});chunks.push(chunk);}
        let input;try{input=JSON.parse(Buffer.concat(chunks).toString('utf8'));}catch{return json(res,400,{error:'提交内容格式不正确。'});}
        const backupDir=(await settings.get()).settings.backups.directory;
        try{
          if(name==='history')return json(res,200,await docTools.history(input,backupDir));
          // 按名单批量生成：preview 只读名单和模板，返回列名对照；正式生成放到“生成文件”文件夹。
          if(name==='mail-merge'){const template=String(input.template||''),list=String(input.list||'');if(!path.isAbsolute(template)||!path.isAbsolute(list))throw Error('请先保存模板文档，并选择名单文件。');
            if(input.preview)return json(res,200,await inspectMerge({template,list}));
            const result=await runMerge({template,list,mode:input.mode==='separate'?'separate':'combined',nameColumn:String(input.nameColumn||''),outputDir:await outputDir()});
            if(input.open&&result.path)await openDocument(result.path,{prefer:'wps'}).catch(()=>{});else if(input.open&&result.folder)await docTools.reveal({path:result.folder}).catch(()=>{});
            return json(res,200,result);}
          // PDF 工具：合并、提取 / 删除页、拆分、图片转 PDF、PDF 转图片。结果放进“生成文件”文件夹，原文件不改；open 时在文件夹里选中结果。
          if(name==='pdf'){const act=String(input.action||''),paths=[].concat(input.paths||[]).map(String).slice(0,300),one=String(input.path||'');let result;
            if(act==='info')return json(res,200,await pdfInfo({path:one}));
            const dir=await outputDir();
            if(act==='merge')result=await mergePdfs({paths,outputDir:dir});else if(act==='pick')result=await pickPages({path:one,pages:String(input.pages||''),mode:input.mode==='remove'?'remove':'keep',outputDir:dir});
            else if(act==='split')result=await splitPdf({path:one,every:input.every,ranges:String(input.ranges||''),outputDir:dir});else if(act==='images-to-pdf')result=await imagesToPdf({paths,size:input.size==='fit'?'fit':'A4',outputDir:dir});
            else if(act==='to-images'){if(docBusy)return json(res,409,{error:'上一个转换还没完成，请稍等。'});docBusy=true;try{result=await pdfToImages({path:one,pages:String(input.pages||''),dpi:input.dpi,outputDir:dir});}finally{docBusy=false;}}
            else return json(res,404,{error:'没有这个功能。'});
            if(input.open)await docTools.reveal({path:result.path||result.folder}).catch(()=>{});return json(res,200,result);}
          // 2026-11-24：批量二维码 / 条形码、PPT 拼长图、批量重命名。结果都放进“生成文件”文件夹，原文件不改（重命名除外，会留记录可改回）。
          if(name==='codes'){const result=await makeCodes({items:input.items,kind:input.kind==='barcode'?'barcode':'qr',outputDir:await outputDir()});if(input.base64){if(result.files.length>200)throw Error('Office 版一次最多 200 个。');result.data=await Promise.all(result.files.map(f=>f?fs.readFile(f).then(b=>b.toString('base64')):''));}return json(res,200,result);}
          if(name==='stitch'){const paths=[].concat(input.paths||[]).map(String),first=paths[0]||'';if(!path.isAbsolute(first))throw Error('图片位置无效。');const out=path.join(path.dirname(first),String(input.name||'长图').replace(/[\\/:*?"<>|\x00-\x1f]/g,'_').slice(0,60)+'.png');const result=await stitchVertical({paths,output:out});if(input.open)await docTools.reveal({path:result.path}).catch(()=>{});return json(res,200,result);}
          if(name==='rename'){const logDir=path.join(await outputDir(),'重命名记录'),act=String(input.action||'preview');
            if(act==='undo')return json(res,200,await undoRename({logDir}));
            if(act==='apply')return json(res,200,await applyRename({paths:input.paths,rule:input.rule,logDir}));
            return json(res,200,await planRename({paths:input.paths,rule:input.rule}));}
          // 2026-11-25：日文文体转换（kuromoji 离线分词，不用 AI）
          if(name==='ja-style')return json(res,200,await jaStyleConvert({paragraphs:input.paragraphs,to:input.to==='polite'?'polite':'plain',plain:input.plain==='da'?'da':'dearu'}));
          // 2026-11-25：数据脱敏。文档里的识别和替换在侧栏里做；这里管词表方案、副本文件名、对照表（加密）、卡片和粘贴栏。
          if(name==='postal'){const act=String(input.action||'status'),items=Array.isArray(input.items)?input.items.slice(0,20001).map(x=>String(x??'').slice(0,200)):[];if(items.length>20000)return json(res,400,{error:'一次最多查 2 万条。'});
            if(act==='status')return json(res,200,await postal.status());
            if(act==='zip'||act==='address'){const results=await (act==='zip'?postal.zip(items):postal.address(items)),st=await postal.status();return json(res,200,{results,date:st.date});}
            if(act==='update')return json(res,200,{info:await postal.update()});
            if(act==='import'){const f=String(input.path||'');if(!path.isAbsolute(f))return json(res,400,{error:'请选择 utf_ken_all.zip 文件。'});return json(res,200,{info:await postal.importFile(f)});}
            return json(res,404,{error:'没有这个功能。'});}
          if(name==='mask'){const act=String(input.action||'state'),lang=input.lang==='ja'?'ja':'zh',id=String(input.id||'');
            if(act==='state')return json(res,200,await maskVault.state());
            if(act==='profiles')return json(res,200,await maskVault.saveProfiles(input));
            if(act==='prefs')return json(res,200,{prefs:await maskVault.savePrefs(input.prefs||{})});
            if(act==='target'){const source=String(input.source||''),fileName=String(input.name||'文档').replace(/[\\/:*?"<>|\x00-\x1f]/g,'_').slice(0,120),m=fileName.match(/^(.*?)(\.[A-Za-z0-9]{1,5})?$/),ext=(m[2]||String(input.ext||'')).toLowerCase();
              const tag=input.suffix==='restore'?(lang==='ja'?'-復元':'-还原'):(lang==='ja'?'-マスキング':'-脱敏');let stem=(m[1]||'文档').replace(/-(脱敏|还原|マスキング|復元)(-\d+)?$/,'');
              const dir=path.isAbsolute(source)&&/[\\/]/.test(source)?path.dirname(source):await outputDir();await fs.mkdir(dir,{recursive:true});
              let target=path.join(dir,stem+tag+ext),n=1;while(await fs.stat(target).then(()=>true,()=>false)){if(++n>99)throw Error('同名文件太多，请清理一下文件夹。');target=path.join(dir,stem+tag+'-'+n+ext);}
              return json(res,200,{path:target,dir});}
            if(act==='save')return json(res,200,{record:await maskVault.addRecord(input.record||{})});
            if(act==='clip-pairs')return json(res,200,{pairs:await maskVault.clipPairs()});
            if(act==='get')return json(res,200,await maskVault.getRecord(id));
            if(act==='find')return json(res,200,{records:await maskVault.findByPath(input.path)});
            if(act==='restored'){await maskVault.markRestored(id,{path:input.path});return json(res,200,{ok:true});}
            if(act==='delete')return json(res,200,{removed:await maskVault.removeRecord(id)});
            if(act==='export'){const p=await exportRecordCsv(await maskVault.getRecord(id),await outputDir(),{lang});if(input.open)await docTools.reveal({path:p}).catch(()=>{});return json(res,200,{path:p});}
            const prefs=await maskVault.prefs(),words=input.words!==undefined?String(input.words):await maskVault.currentWords(),types=Array.isArray(input.types)?input.types:prefs.types;
            if(act==='modules-scan')return json(res,200,{items:await scanModules({cards,pasteBar,types,words})});
            if(act==='modules-apply')return json(res,200,await maskModules({cards,pasteBar,vault:maskVault,items:input.items,mode:['star','block','code'].includes(input.mode)?input.mode:prefs.mode,lang,types,words}));
            if(act==='modules-restore')return json(res,200,await restoreModules({cards,pasteBar,vault:maskVault,id}));
            return json(res,404,{error:'没有这个功能。'});}
          if(name==='reveal')return json(res,200,await docTools.reveal(input));
          // 批量处理前查哪些输出文件名已经被占用，避免覆盖（只查是否存在，不读内容）。
          if(name==='exists'){const list=[].concat(input.paths||[]).slice(0,500).map(String);return json(res,200,{exists:await Promise.all(list.map(p=>path.isAbsolute(p)?fs.stat(p).then(()=>true,()=>false):false))});}
          if(name==='auto-backup'){
            // 运行会改动文档的工具前自动留一份：同一个文件 10 分钟内只留一次，每个文件只保留最近 20 份自动备份。
            const source=String(input.path||'');if(!path.isAbsolute(source))return json(res,200,{skipped:'unsaved'});
            const ext=path.extname(source).toLowerCase();if(!['.doc','.docx','.wps','.rtf','.xls','.xlsx','.et','.csv','.ppt','.pptx','.dps'].includes(ext))return json(res,200,{skipped:'type'});
            const last=autoSeen.get(source)||0;if(Date.now()-last<10*60*1000)return json(res,200,{skipped:'recent'});
            const info=await fs.stat(source).catch(()=>null);if(!info||!info.isFile())return json(res,200,{skipped:'missing'});if(info.size>100*1024*1024)return json(res,200,{skipped:'large'});
            await fs.mkdir(backupDir,{recursive:true});const stamp=new Date(),pad=n=>String(n).padStart(2,'0');
            const target=path.join(backupDir,path.basename(source,path.extname(source))+'-备份-'+stamp.getFullYear()+pad(stamp.getMonth()+1)+pad(stamp.getDate())+'-'+pad(stamp.getHours())+pad(stamp.getMinutes())+pad(stamp.getSeconds())+'-自动'+path.extname(source));
            try{await fs.copyFile(source,target,fs.constants?.COPYFILE_EXCL??1);}catch(e){if(e.code!=='EEXIST')throw e;}
            autoSeen.set(source,Date.now());await docTools.pruneAuto(source,backupDir,20);return json(res,200,{backupPath:target});
          }
          if(docBusy)return json(res,409,{error:'上一个文件还在处理，请稍等。'});docBusy=true;
          try{
            if(name==='ocr')return json(res,200,await docTools.ocr(input));
            if(name==='pdf-to-word')return json(res,200,await docTools.pdfToWord(input));
            if(name==='extract-images')return json(res,200,await docTools.extractImages(input));
            if(name==='compress-images')return json(res,200,await docTools.compressImages(input));
          }finally{docBusy=false;}
        }catch(e){return json(res,422,{error:e.message});}
      }
      if (url.pathname === '/api/backup' && req.method === 'POST') {
        const supplied=Buffer.from(String(req.headers['x-coach-token']||'')),expected=Buffer.from(token);
        if(supplied.length!==expected.length||!timingSafeEqual(supplied,expected))return json(res,403,{error:'本机连接校验失败，请重新打开工具箱。'});
        if(!/^application\/json\b/i.test(req.headers['content-type']||''))return json(res,415,{error:'需要 JSON 请求。'});
        let chunks=[],size=0;for await(const chunk of req){size+=chunk.length;if(size>8000)return json(res,413,{error:'请求过大。'});chunks.push(chunk);}
        try{
          const source=String(JSON.parse(Buffer.concat(chunks).toString('utf8')).path||'');
          if(!path.isAbsolute(source))throw new Error('文件位置无效，请先保存文档。');
          const ext=path.extname(source).toLowerCase();
          if(!['.doc','.docx','.wps','.wpt','.rtf','.xls','.xlsx','.et','.csv','.ppt','.pptx','.dps','.pdf'].includes(ext))throw new Error('只能备份文档、表格、演示或 PDF 文件。');
          const info=await fs.stat(source).catch(()=>null);
          if(!info||!info.isFile())throw new Error('找不到这个文件，请确认文档已经保存到电脑上。');
          if(info.size>200*1024*1024)throw new Error('文件超过 200 MB，请手动复制备份。');
          const actualBackupDir=(await settings.get()).settings.backups.directory;await fs.mkdir(actualBackupDir,{recursive:true});
          const stamp=new Date(),pad=n=>String(n).padStart(2,'0');
          const name=path.basename(source,path.extname(source))+'-备份-'+stamp.getFullYear()+pad(stamp.getMonth()+1)+pad(stamp.getDate())+'-'+pad(stamp.getHours())+pad(stamp.getMinutes())+pad(stamp.getSeconds())+ext;
          let target=path.join(actualBackupDir,name),n=1;
          while(await fs.stat(target).then(()=>true,()=>false))target=path.join(actualBackupDir,name.replace(ext,'-'+(++n)+ext));
          await fs.copyFile(source,target,fs.constants?.COPYFILE_EXCL??1);
          return json(res,200,{backupPath:target});
        }catch(e){return json(res,422,{error:e.message});}
      }
      if (/^\/api\/pastebar(\/(pin|clipboard))?$/.test(url.pathname) && ['GET','POST','DELETE'].includes(req.method)) {
        const supplied=Buffer.from(String(req.headers['x-coach-token']||'')),expected=Buffer.from(token);
        if(supplied.length!==expected.length||!timingSafeEqual(supplied,expected))return json(res,403,{error:'本机连接校验失败，请重新打开工具箱。'});
        const state=async()=>({items:await pasteBar.list(),watching:!!clipWatch,skipped:clipSkipped});
        try{
          if(req.method==='GET')return json(res,200,await state());
          if(req.method==='DELETE'){const id=url.searchParams.get('id');if(id){if(!await pasteBar.remove(id))return json(res,404,{error:'这条内容已经不在粘贴栏里了。'});}else await pasteBar.clear();return json(res,200,await state());}
          if(url.pathname==='/api/pastebar/clipboard'){const text=await clipboardReader();if(!String(text||'').trim())throw Error('剪贴板里没有文字。');const item=await autoMask('paste',await pasteBar.add(text,'clipboard','clipboard'));return json(res,200,{item,...await state()});}
          if(!/^application\/json\b/i.test(req.headers['content-type']||''))return json(res,415,{error:'需要 JSON 请求。'});
          let chunks=[],size=0;for await(const chunk of req){size+=chunk.length;if(size>200000)return json(res,413,{error:'内容过长，请缩小选区。'});chunks.push(chunk);}
          const body=JSON.parse(Buffer.concat(chunks).toString('utf8')||'{}');
          if(url.pathname==='/api/pastebar/pin'){await pasteBar.pin(String(body.id||''),!!body.pinned);return json(res,200,await state());}
          const item=await autoMask('paste',await pasteBar.add(body.block,body.sourceHost,'manual'));return json(res,200,{item,...await state()});
        }catch(e){return json(res,422,{error:e.message});}
      }
      if (url.pathname === '/api/cards' && ['GET','POST','DELETE'].includes(req.method)) {
        const supplied=Buffer.from(String(req.headers['x-coach-token']||'')),expected=Buffer.from(token);
        if(supplied.length!==expected.length||!timingSafeEqual(supplied,expected))return json(res,403,{error:'本机连接校验失败，请重新打开工具箱。'});
        if(req.method==='GET')return json(res,200,{cards:(await cards.list()).map(visibleCard)});
        if(req.method==='DELETE'){const id=url.searchParams.get('id');if(id){if(!await cards.remove(id))return json(res,404,{error:'这张卡片已经不存在了。'});}else await cards.clear();return json(res,200,{ok:true});}
        if(!/^application\/json\b/i.test(req.headers['content-type']||''))return json(res,415,{error:'需要 JSON 请求。'});
        let chunks=[],size=0;for await(const chunk of req){size+=chunk.length;if(size>800000)return json(res,413,{error:'卡片内容过长，请缩小选区。'});chunks.push(chunk);}
        try{const body=JSON.parse(Buffer.concat(chunks).toString('utf8'));const added=await cards.add(body.block,body.sourceHost,body.format,body.source),card=added.source?added:await autoMask('card',added);return json(res,200,{card});}catch(e){return json(res,422,{error:e.message});}
      }
      if (url.pathname === '/api/local-transfer' && ['GET','POST','DELETE'].includes(req.method)) {
        const supplied=Buffer.from(String(req.headers['x-coach-token']||'')),expected=Buffer.from(token);
        if(supplied.length!==expected.length||!timingSafeEqual(supplied,expected))return json(res,403,{error:'本机连接校验失败，请重新打开工具箱。'});
        if(req.method==='DELETE'){transfer=null;return json(res,200,{ok:true});}
        if(req.method==='GET'){if(!transfer||Date.now()-transfer.created>1800000){transfer=null;return json(res,404,{error:'中转区为空或已过期，请先在来源文档中点击“选区放入中转区”。'});}return json(res,200,{block:transfer.block,sourceHost:transfer.sourceHost});}
        let chunks=[],size=0;for await(const chunk of req){size+=chunk.length;if(size>80000)return json(res,413,{error:'中转内容过长，请缩小选区。'});chunks.push(chunk);}
        try{const body=JSON.parse(Buffer.concat(chunks).toString('utf8')),block=validateBlocks([body.block])[0];if(!['wps','et','wpp'].includes(body.sourceHost))throw Error('来源应用无效。');if(block.type==='text'&&block.text.length>10000)throw Error('中转文字最多 1 万字。');if(block.type==='table'&&(block.rows.length>20||block.rows[0].length>10))throw Error('中转表格最多 20 行、10 列。');transfer={block,sourceHost:body.sourceHost,created:Date.now()};try{await cards.add(block,body.sourceHost);}catch(e){console.error('Card save failed:',e.message);}return json(res,200,{ok:true});}catch(e){return json(res,422,{error:e.message});}
      }
      if (req.method === 'POST' && ['/api/workspace/import','/api/workspace/check','/api/workspace/export'].includes(url.pathname)) {
        const supplied = Buffer.from(String(req.headers['x-coach-token'] || ''));
        const expected = Buffer.from(token);
        if (supplied.length !== expected.length || !timingSafeEqual(supplied, expected)) return json(res,403,{error:'连接已过期，请刷新页面。'});
        const chunks=[]; let size=0;
        for await (const chunk of req) {size+=chunk.length; if(size>21*1024*1024) return json(res,413,{error:'文件超过 15 MB 或请求过大，请拆分后再试。'});chunks.push(chunk);}
        try {
          if (!/^application\/json\b/i.test(req.headers['content-type'] || '')) return json(res,415,{error:'需要 JSON 请求。'});
          const body=JSON.parse(Buffer.concat(chunks).toString('utf8'));
          if(url.pathname.endsWith('/import')) return json(res,200,await importFile(body.name,Buffer.from(String(body.data || ''),'base64')));
          if(url.pathname.endsWith('/check')) { const result=preflight(body.blocks,body.target); return json(res,200,{warnings:result.warnings,target:result.target,count:result.blocks.length}); }
          // Validate again at execution: the preview may have become stale.
          const output=await exportFile(body.blocks,body.target);
          res.writeHead(200,{'Content-Type':'application/octet-stream','Content-Disposition':`attachment; filename="workspace.${body.target}"`,'Cache-Control':'no-store'}); return res.end(output);
        } catch(error) { console.error('Workspace conversion failed:',error.message); return json(res,422,{error:'转换未完成：'+(error.message || '文件可能已损坏、被加密或不受支持，请在 WPS 中检查后重试。')}); }
      }
      if(req.method==='GET'&&url.pathname==='/api/updates/startup')return json(res,200,await startupRelease);
      if (req.method === 'GET' && url.pathname === '/api/session') return json(res, 200, { token });
      if (req.method === 'GET' && url.pathname === '/api/health') {
        if(!aiEnabled)return json(res,200,{app:'wps-ai-coach',product:'办公工具箱',version:VERSION,localTools:true,aiEnabled:false});
        let installed = false, engine = false;
        try { const response = await fetchImpl(`${ollamaUrl}/api/tags`, { signal: AbortSignal.timeout(2500) }); const data = await response.json(); engine = response.ok; installed = (data.models || []).some(m => m.name === model || m.model === model); } catch {}
        return json(res, 200, { app: 'wps-ai-coach', version: '0.2.0', model, engine, installed, busy });
      }
      if (req.method === 'POST' && url.pathname === '/api/ask') {
        if(!aiEnabled)return json(res,410,{error:'当前版本为本地办公插件，未启用 AI。'});
        const presented = Buffer.from(String(req.headers['x-coach-token'] || ''));
        const expected = Buffer.from(token);
        if (presented.length !== expected.length || !timingSafeEqual(presented, expected)) return json(res, 403, { error: '页面连接已过期，请刷新侧边栏。' });
        if (!/^application\/json\b/i.test(req.headers['content-type'] || '')) return json(res, 415, { error: '需要 JSON 请求。' });
        let chunks = [], size = 0;
        for await (const chunk of req) { size += chunk.length; if (size > 40000) { json(res, 413, { error: '选中内容过长，请缩小范围。' }); return; } chunks.push(chunk); }
        let request;
        try { request = normalizeRequest(JSON.parse(Buffer.concat(chunks).toString('utf8'))); } catch (error) { return json(res, 400, { error: error.message }); }
        if (busy) return json(res, 409, { error: '正在回答上一个问题，请稍等。' });
        busy = true;
        const controller = new AbortController();
        res.on('close', () => { if (!res.writableEnded) controller.abort(); });
        try {
          const response = await fetchImpl(`${ollamaUrl}/api/chat`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, signal: AbortSignal.any([controller.signal, AbortSignal.timeout(180000)]), body: JSON.stringify({ model, messages: makeMessages(request), stream: false, keep_alive: '10m', options: { num_ctx: 4096, num_predict: 2048, temperature: 0.4 } }) });
          const data = await response.json();
          if (!response.ok) throw new Error(data.error || '本地模型响应失败。');
          const answer = stripThinking(data.message?.content);
          if (!answer) throw new Error('模型没有生成完整回答，请简化问题后重试。');
          return json(res, 200, { answer, model, local: true });
        } catch (error) {
          const message = error.name === 'TimeoutError' ? '本机模型这次思考较久，请缩小问题后重试。' : error.message === 'fetch failed' ? '未连接到 Ollama。请先启动 Ollama，再重新提问。' : error.message;
          if (!res.destroyed) return json(res, 502, { error: message });
        } finally { busy = false; }
        return;
      }
      if (req.method !== 'GET' && req.method !== 'HEAD') return json(res, 405, { error: '不支持此操作。' });
      let relative = decodeURIComponent(url.pathname);
      relative = relative.replace(/^\/addon\/(wps|et|wpp)\//, '/');
      if (relative === '/') relative = '/tools.html';
      if (relative === '/sidebar.html') relative='/tools.html';
      const target = path.resolve(PUBLIC, '.' + relative);
      if (!target.startsWith(PUBLIC + path.sep)) return json(res, 403, { error: '路径无效。' });
      const content = await fs.readFile(target);
      const scriptPolicy = relative === '/publish.html' ? "'self' 'unsafe-inline'" : "'self'";
      res.writeHead(200, { 'Content-Type': MIME[path.extname(target)] || 'application/octet-stream', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff', 'Referrer-Policy': 'no-referrer', 'Content-Security-Policy': `default-src 'self'; script-src ${scriptPolicy}; style-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'self'; object-src 'none'; base-uri 'self'` });
      res.end(req.method === 'HEAD' ? undefined : content);
    } catch (error) { if (!res.headersSent) json(res, error.code === 'ENOENT' ? 404 : 500, { error: '未找到页面或服务暂时不可用。' }); }
  });
  server.once('listening',()=>{startupRelease=settings.get().then(s=>checkRelease(s.settings.updates.url,VERSION)).catch(()=>({status:'unavailable'}));});
  server.once('close',()=>{mcp.close();mcpMasker.close();});
  server.on('close',()=>{if(clipWatch){clipWatch.stop();clipWatch=null;}});
  return server;
}
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const server = createCoachServer({ cardsFile: path.join(ROOT, 'generated', 'cards.json'),pasteBarFile:path.join(ROOT,'generated','pastebar.json'),settingsFile:path.join(ROOT,'generated','settings.json') });
  server.on('error', error => { console.error(error.message); process.exitCode = 1; });
  server.listen(39127, '127.0.0.1', () => console.log('办公工具箱已启动：http://127.0.0.1:39127/'));
  startOfficeServer({ root: ROOT, publicDir: PUBLIC }).catch(error => console.error(error.message));
}
