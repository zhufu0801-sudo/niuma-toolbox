// 办公工具箱 · 给 Microsoft Office 插件用的本机 https 服务：提供侧栏需要的静态文件，并把内容卡片、备份、导出 PDF 这几个接口转给本机工具箱服务。
// Office 网页插件必须从 https 加载；证书由安装脚本在本机生成并经用户确认后信任，保存在 generated\office-cert.pfx。
import https from 'node:https';
import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';
export const OFFICE_PORT=39128;
const MIME={'.json':'application/json; charset=utf-8','.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.svg':'image/svg+xml','.png':'image/png','.xml':'application/xml; charset=utf-8'};
const SHARED=new Set(['/update-banner.js','/print-copy-panel.js','/library-panel.js','/mcp-host.js','/welcome.js','/i18n-ja.js','/i18n-ja-extra.js','/i18n.js','/method-cards.js','/icons/toolbox.svg','/icons/shield.svg','/icons/macrocheck.svg','/icons/postal.svg','/tools.css','/quick-panel.css','/formula-panel.css','/quick-engine.js','/formula-engine.js','/quick-host.js','/quick-catalog.js','/quick-panel.js','/formula-panel.js','/ai-text-panel.js','/locale-engine.js','/ai-formula.js','/doc-outline.js','/sheet-logic.js','/ai-chat-panel.js','/file-tools-panel.js','/proof-engine.js','/id-engine.js','/mask-engine.js','/mask-host.js','/mask-panel.js','/mask-clipboard.js','/vba-lint.js','/vba-panel.js','/postal-panel.js','/postal-format.js']);
export function officeFileAllowed(pathname){return pathname==='/office/shortcuts.json'||/^\/templates\/ppt\/[-a-z0-9]+\.png$/.test(pathname)||/^\/office\/[a-z0-9-]+\.(html|js|css|png)$/.test(pathname)||SHARED.has(pathname);}
// 只转发侧栏用到的接口（卡片、备份、PDF、AI、文件与图片工具）；不带跨域许可头，别的网页读不到内容；写入必须带本机口令头（会触发跨域预检，本服务不放行预检）。
const WRITE={'/api/library':100000,'/api/mcp/register':300000,'/api/mcp/poll':300000,'/api/mcp/result':300000,'/api/cards':800000,'/api/backup':8000,'/api/save-pdf':110*1024*1024,'/api/ai/translate':80000,'/api/ai/formula':80000,'/api/ai/chat':160000,'/api/ai/write':80000,'/api/studio/open':400000,
  '/api/tools/ocr':40*1024*1024,'/api/tools/pdf-to-word':120*1024*1024,'/api/tools/extract-images':8000,'/api/tools/compress-images':8000,'/api/tools/history':8000,'/api/tools/reveal':8000,'/api/tools/exists':200000,'/api/tools/auto-backup':8000,'/api/tools/codes':200000,'/api/tools/ja-style':1600000,'/api/tools/mask':20*1024*1024,'/api/tools/postal':8*1024*1024};
const PROXY=new Set(['/api/updates/startup','/api/session','/api/cards','/api/ai/config',...Object.keys(WRITE)]);
function proxy(req,res,pathname,host,apiPort){
  const origin=req.headers.origin,json=(code,o)=>{res.writeHead(code,{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store'});res.end(JSON.stringify(o));};
  if(origin&&origin!==`https://${host}`)return json(403,{error:'请求来源不匹配。'});
  const write=req.method==='POST'&&WRITE[pathname];
  if(req.method!=='GET'&&!write||req.method==='GET'&&!/^\/api\/(session|cards|ai\/config|updates\/startup)$/.test(pathname))return json(405,{error:'不支持的请求。'});
  if(write&&!req.headers['x-coach-token'])return json(403,{error:'本机连接校验失败，请重新打开工具箱。'});
  const headers={host:`127.0.0.1:${apiPort}`};if(req.headers['x-coach-token'])headers['x-coach-token']=String(req.headers['x-coach-token']);if(write)headers['content-type']=String(req.headers['content-type']||'');
  const up=http.request({host:'127.0.0.1',port:apiPort,method:req.method,path:pathname,headers},r=>{res.writeHead(r.statusCode||502,{'Content-Type':'application/json; charset=utf-8','Cache-Control':'no-store','X-Content-Type-Options':'nosniff'});r.pipe(res);});
  up.on('error',()=>{if(!res.headersSent)json(502,{error:'本地服务未连接，请先双击桌面上的“办公工具箱”。'});});
  if(!write)return up.end();
  let size=0;req.on('data',c=>{size+=c.length;if(size>write){up.destroy();if(!res.headersSent)json(413,{error:'内容过大。'});req.destroy();}else up.write(c);});req.on('end',()=>up.end());
}
export function createOfficeServer({pfx,passphrase,cert,key,publicDir,port=OFFICE_PORT,apiPort=39127}){
  return https.createServer(cert&&key?{cert,key}:{pfx,passphrase},async(req,res)=>{
    const host=String(req.headers.host||'');
    if(host!==`localhost:${port}`&&host!==`127.0.0.1:${port}`){res.writeHead(403);return res.end();}
    const api=String(req.url||'').split('?')[0];if(PROXY.has(api))return proxy(req,res,api,host,apiPort);
    if(req.method!=='GET'&&req.method!=='HEAD'){res.writeHead(405,{Allow:'GET, HEAD'});return res.end();}
    let pathname;try{pathname=decodeURIComponent(new URL(req.url,`https://${host}`).pathname);}catch{res.writeHead(400);return res.end();}
    if(pathname==='/'||pathname==='/office/')pathname='/office/taskpane.html';
    if(!officeFileAllowed(pathname)){res.writeHead(404);return res.end();}
    const target=path.resolve(publicDir,'.'+pathname);
    if(!target.startsWith(path.resolve(publicDir)+path.sep)){res.writeHead(403);return res.end();}
    try{const body=await fs.readFile(target);res.writeHead(200,{'Content-Type':MIME[path.extname(target)]||'application/octet-stream','Cache-Control':'no-store','X-Content-Type-Options':'nosniff','Referrer-Policy':'no-referrer'});res.end(req.method==='HEAD'?undefined:body);}
    catch{res.writeHead(404);res.end();}
  });
}
// 有证书才启动；没装 Office 版时什么也不做。
export async function startOfficeServer({root,publicDir,log=console}){
  // Windows 安装脚本生成 office-cert.pfx；Mac 安装脚本用系统自带的 openssl 生成 office-cert.pem + office-key.pem。
  const dir=path.join(root,'generated');let options=null;
  try{options={cert:await fs.readFile(path.join(dir,'office-cert.pem')),key:await fs.readFile(path.join(dir,'office-key.pem'))};}catch{}
  if(!options)try{options={pfx:await fs.readFile(path.join(dir,'office-cert.pfx')),passphrase:(await fs.readFile(path.join(dir,'office-cert.key'),'utf8')).trim()};}catch{return null;}
  const server=createOfficeServer({...options,publicDir});
  server.on('error',e=>log.error('Office 插件服务未能启动：'+e.message));
  server.listen(OFFICE_PORT,'127.0.0.1',()=>log.log('Office 插件服务已启动：https://localhost:'+OFFICE_PORT+'/office/taskpane.html'));
  return server;
}
