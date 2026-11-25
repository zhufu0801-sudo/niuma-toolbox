// 办公工具箱 · 文件与图片工具（本机服务端）：图片转文字、PDF 转 Word、提取文档里的图片、压缩文档里的图片、历史版本。
// 不加新依赖：文字识别用系统自带的（Windows 的 Windows.Media.Ocr、macOS 的 Vision）；压缩图片用系统自带的（Windows 的 System.Drawing、macOS 的 sips）。
// 原文件一律不改：结果另存为新文件（同一文件夹里加后缀，已存在就再加序号）。
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { execFile } from 'node:child_process';
import JSZip from 'jszip';
import { Document, Packer, Paragraph, HeadingLevel, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, AlignmentType } from 'docx';
import { ensurePdfEnv } from './pdf-env.mjs';
import { layoutOcr } from './ocr-layout.mjs';

const run=(cmd,args,opts={})=>new Promise((resolve,reject)=>{execFile(cmd,args,{windowsHide:true,maxBuffer:64*1024*1024,timeout:opts.timeout||120000,env:{...process.env,...(opts.env||{})},encoding:'utf8'},(err,stdout,stderr)=>{if(err){err.stderr=String(stderr||'');return reject(err);}resolve({stdout:String(stdout||''),stderr:String(stderr||'')});});});
export const OFFICE_ZIP=['.docx','.docm','.xlsx','.xlsm','.pptx','.pptm'];
const IMG_EXT=/\.(png|jpe?g|gif|bmp|tiff?|webp|emf|wmf|svg)$/i;
const CJK='\\u3000-\\u303f\\u3040-\\u30ff\\u3400-\\u9fff\\uf900-\\ufaff\\uff00-\\uffef';
// Windows 识别中文、日文时会在字之间加空格，这里去掉（英文单词之间的空格保留）。
// 版面还原的结果转成纯文字（表格每行用 Tab 分隔），给“图片转文字”用。
export function layoutText(segs){return segs.map(x=>x.table?x.table.map(r=>r.join('\t')).join('\n'):x.para?x.para.text:x.text).join('\n');}
export function tidyOcr(text){return String(text||'').replace(new RegExp('(['+CJK+'])[ \\t]+(?=['+CJK+'])','g'),'$1').replace(/[ \t]+\n/g,'\n').replace(/\n{3,}/g,'\n\n').trim();}
export function imageType(buf){if(!buf||buf.length<8)return '';if(buf[0]===0x89&&buf[1]===0x50&&buf[2]===0x4e&&buf[3]===0x47)return 'png';if(buf[0]===0xff&&buf[1]===0xd8)return 'jpg';if(buf.subarray(0,3).toString('latin1')==='GIF')return 'gif';if(buf[0]===0x42&&buf[1]===0x4d)return 'bmp';if((buf[0]===0x49&&buf[1]===0x49)||(buf[0]===0x4d&&buf[1]===0x4d))return 'tif';if(buf.subarray(0,4).toString('latin1')==='RIFF'&&buf.subarray(8,12).toString('latin1')==='WEBP')return 'webp';return '';}
// OCR 语言：zh 中文简体、ja 日文、en 英文、auto 按系统语言。
const WIN_LANG={zh:'zh-Hans-CN',ja:'ja-JP',en:'en-US'},MAC_LANG={zh:['zh-Hans','en-US'],ja:['ja-JP','en-US'],en:['en-US'],auto:['zh-Hans','ja-JP','en-US']};
const WIN_OCR=String.raw`$ErrorActionPreference='Stop';[Console]::OutputEncoding=[System.Text.UTF8Encoding]::new($false)
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null=[Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime];$null=[Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime];$null=[Windows.Graphics.Imaging.BitmapDecoder,Windows.Graphics,ContentType=WindowsRuntime];$null=[Windows.Globalization.Language,Windows.Globalization,ContentType=WindowsRuntime]
$as=([System.WindowsRuntimeSystemExtensions].GetMethods()|Where-Object{$_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation` + '`' + String.raw`1'})[0]
function Wait($op,[Type]$t){$task=$as.MakeGenericMethod($t).Invoke($null,@($op));$task.Wait(-1)|Out-Null;$task.Result}
$langs=@([Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages|ForEach-Object{$_.LanguageTag})
$engine=$null;$used=''
if($env:OCR_LANG -ne ''){$l=[Windows.Globalization.Language]::new($env:OCR_LANG);if([Windows.Media.Ocr.OcrEngine]::IsLanguageSupported($l)){$engine=[Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($l);$used=$env:OCR_LANG}}
if($engine -eq $null){$engine=[Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages();if($engine -ne $null){$used=$engine.RecognizerLanguage.LanguageTag}}
if($engine -eq $null){Write-Output ('{"error":"nolang","langs":'+(ConvertTo-Json -Compress @($langs))+'}');exit 0}
$out=@()
foreach($p in ($env:OCR_FILES -split '\|')){if($p -eq ''){continue}
 $file=Wait ([Windows.Storage.StorageFile]::GetFileFromPathAsync($p)) ([Windows.Storage.StorageFile])
 $stream=Wait ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
 $dec=Wait ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
 $max=[Windows.Media.Ocr.OcrEngine]::MaxImageDimension
 if($dec.PixelWidth -gt $max -or $dec.PixelHeight -gt $max){$out+=@{file=$p;error='toolarge'};$stream.Dispose();continue}
 $bmp=Wait ($dec.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
 $r=Wait ($engine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])
 $ls=@();foreach($ln in $r.Lines){$ws=@();foreach($w in $ln.Words){$b=$w.BoundingRect;$ws+=@{t=$w.Text;x=[int]$b.X;y=[int]$b.Y;w=[int]$b.Width;h=[int]$b.Height}};$ls+=@{t=$ln.Text;words=@($ws)}}
 $out+=@{file=$p;width=[int]$dec.PixelWidth;lines=@($ls)};$stream.Dispose()}
Write-Output (ConvertTo-Json -Compress -Depth 9 @{lang=$used;langs=@($langs);results=@($out)})`;
const MAC_OCR=`ObjC.import('Vision');ObjC.import('Foundation');
function one(path,langs){var url=$.NSURL.fileURLWithPath(path);var req=$.VNRecognizeTextRequest.alloc.init;req.recognitionLevel=0;req.usesLanguageCorrection=true;
 try{req.recognitionLanguages=$(langs);}catch(e){}
 var h=$.VNImageRequestHandler.alloc.initWithURLOptions(url,$({}));var ok=h.performRequestsError($([req]),null);if(!ok)return {file:path,error:'failed'};
 var res=req.results,out=[];for(var i=0;i<res.count;i++){var o=res.objectAtIndex(i),c=o.topCandidates(1);if(!c.count)continue;var b=o.boundingBox;out.push({t:c.objectAtIndex(0).string.js,x:b.origin.x,y:b.origin.y,h:b.size.height});}
 return {file:path,items:out};}
function run(argv){var langs=argv[0].split(','),files=argv.slice(1),r=[];for(var i=0;i<files.length;i++)r.push(one(files[i],langs));return JSON.stringify({results:r});}`;
// macOS 把 PDF 每一页画成 PNG（PDFKit），给扫描件做文字识别用。
const MAC_PDF=`ObjC.import('PDFKit');ObjC.import('AppKit');ObjC.import('Foundation');
function run(argv){var doc=$.PDFDocument.alloc.initWithURL($.NSURL.fileURLWithPath(argv[0]));if(!doc||doc.isNil())return JSON.stringify({error:'open'});var n=Math.min(doc.pageCount,Number(argv[2])||30),files=[];
 for(var i=0;i<n;i++){var pg=doc.pageAtIndex(i),box=pg.boundsForBox(0),w=box.size.width*2,hh=box.size.height*2;var img=pg.thumbnailOfSizeForBox($.NSMakeSize(w,hh),0);var rep=$.NSBitmapImageRep.imageRepWithData(img.TIFFRepresentation);var png=rep.representationUsingTypeProperties(4,$({}));var f=argv[1]+'/p'+(i+1)+'.png';png.writeToFileAtomically(f,true);files.push(f);}
 return JSON.stringify({files:files,pages:doc.pageCount});}`;
const WIN_SHRINK=String.raw`$ErrorActionPreference='SilentlyContinue';[Console]::OutputEncoding=[System.Text.UTF8Encoding]::new($false)
Add-Type -AssemblyName System.Drawing
$q=[long]$env:SHRINK_Q;$mw=[int]$env:SHRINK_W
$codec=[System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders()|Where-Object{$_.MimeType -eq 'image/jpeg'}
$ep=New-Object System.Drawing.Imaging.EncoderParameters(1);$ep.Param[0]=New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality,$q)
$done=@()
Get-ChildItem -LiteralPath $env:SHRINK_DIR -File|ForEach-Object{
 $f=$_.FullName;$img=$null
 try{$img=[System.Drawing.Image]::FromFile($f)
  $rot=1;if($img.PropertyIdList -contains 274){$rot=[int]$img.GetPropertyItem(274).Value[0]}
  $jpg=$_.Extension -match '^\.jpe?g$';$w=$img.Width;$h=$img.Height;$s=[Math]::Min(1.0,$mw/[double][Math]::Max($w,$h))
  if($rot -ne 1 -or (-not $jpg -and $s -ge 1)){$img.Dispose();return}
  $nw=[int][Math]::Max(1,[Math]::Round($w*$s));$nh=[int][Math]::Max(1,[Math]::Round($h*$s))
  $bmp=New-Object System.Drawing.Bitmap($nw,$nh);$bmp.SetResolution($img.HorizontalResolution,$img.VerticalResolution);$g=[System.Drawing.Graphics]::FromImage($bmp);$g.InterpolationMode=[System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  if(-not $jpg){$g.Clear([System.Drawing.Color]::Transparent)}
  $g.DrawImage($img,0,0,$nw,$nh);$g.Dispose();$img.Dispose();$img=$null
  if($jpg){$bmp.Save($f+'.out',$codec,$ep)}else{$bmp.Save($f+'.out',[System.Drawing.Imaging.ImageFormat]::Png)}
  $bmp.Dispose();$done+=$_.Name}catch{if($img){$img.Dispose()}}}
Write-Output (ConvertTo-Json -Compress @($done))`;

export function createDocTools({platform=process.platform,exec=run,tmpRoot=os.tmpdir(),fallbackDir=async()=>path.join(os.homedir(),'Documents','办公工具箱备份')}={}){
  async function tempDir(){return fs.mkdtemp(path.join(tmpRoot,'officetools-'));}
  async function readSource(input,allowed,maxMb=60){
    if(input.data){const buf=Buffer.from(String(input.data).replace(/^data:[^,]*,/,''),'base64');if(!buf.length)throw Error('文件是空的。');if(buf.length>maxMb*1024*1024)throw Error('文件超过 '+maxMb+' MB。');const name=String(input.name||'文件').replace(/[\\/:*?"<>|\x00-\x1f]/g,'_').slice(0,120);return {buf,name,file:null};}
    const file=String(input.path||'');if(!path.isAbsolute(file))throw Error('请先把文档保存到电脑上，再执行。');
    const ext=path.extname(file).toLowerCase();if(allowed&&!allowed.includes(ext))throw Error('不支持这种文件（'+(ext||'没有扩展名')+'）。支持：'+allowed.join('、')+'。');
    const st=await fs.stat(file).catch(()=>null);if(!st||!st.isFile())throw Error('找不到这个文件，请确认已经保存。');if(st.size>maxMb*1024*1024)throw Error('文件超过 '+maxMb+' MB。');
    return {buf:await fs.readFile(file),name:path.basename(file),file};
  }
  // 新文件放在原文件旁边；上传的文件（没有路径）放到备份文件夹。不覆盖已有文件。
  async function outPath(src,suffix,ext){const dir=src.file?path.dirname(src.file):await fallbackDir();await fs.mkdir(dir,{recursive:true});const base=path.basename(src.name,path.extname(src.name))||'文件';let p=path.join(dir,base+suffix+ext),n=1;while(await fs.stat(p).then(()=>true,()=>false)){if(++n>99)throw Error('同名文件太多了。');p=path.join(dir,base+suffix+'-'+n+ext);}return p;}
  async function writeNew(p,buf){await fs.writeFile(p,buf,{flag:'wx'});return p;}

  // ---- 图片转文字 ----
  async function ocrFiles(files,lang='auto'){
    if(!files.length)return {results:[],lang:''};
    if(platform==='win32'){
      let out;try{out=(await exec('powershell.exe',['-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-Command',WIN_OCR],{env:{OCR_FILES:files.join('|'),OCR_LANG:WIN_LANG[lang]||''},timeout:180000})).stdout;}catch(e){throw Error('Windows 文字识别没能运行：'+(String(e.stderr||e.message).split('\n')[0]).slice(0,200));}
      let d;try{d=JSON.parse(out.trim().split('\n').pop());}catch{throw Error('Windows 文字识别返回的结果看不懂，请重试。');}
      if(d.error==='nolang')throw Error('这台电脑还没有可用的文字识别语言。请在“设置 → 时间和语言 → 语言和区域”里添加中文（简体）或日语，并勾选“光学字符识别”。');
      // 有每个词的位置时按版面重排（同一行的放一起、表格按格子），没有时按识别顺序。
      const results=[].concat(d.results||[]).map(r=>{const lines=[].concat(r.lines||[]),plain=tidyOcr(lines.map(l=>typeof l==='string'?l:l&&l.t||'').join('\n'));let layout=null;try{if(lines.some(l=>l&&l.words))layout=layoutOcr(lines,r.width);}catch{layout=null;}return {file:r.file,error:r.error==='toolarge'?'图片太大（超过 Windows 文字识别的尺寸上限），请缩小后再试。':r.error,text:layout?layoutText(layout):plain,layout};});
      const want=WIN_LANG[lang],note=want&&d.lang&&d.lang!==want?(lang==='zh'?'这台电脑没有中文识别语言，已改用 '+d.lang+'。添加方法：“设置 → 时间和语言 → 语言和区域”里添加中文（简体）。':lang==='ja'?'这台电脑没有日文识别语言，已改用 '+d.lang+'。添加方法：“设置 → 时间和语言 → 语言和区域”里添加日语。':''):'';
      return {results,lang:d.lang,engine:'Windows',note};
    }
    if(platform==='darwin'){
      let out;try{out=(await exec('osascript',['-l','JavaScript','-e',MAC_OCR,(MAC_LANG[lang]||MAC_LANG.auto).join(','),...files],{timeout:180000})).stdout;}catch(e){throw Error('macOS 文字识别没能运行（需要 macOS 10.15 或更新版本；日文需要 macOS 13 或更新版本）：'+String(e.stderr||e.message).split('\n')[0].slice(0,200));}
      const d=JSON.parse(out.trim());
      const results=d.results.map(r=>{if(r.error)return {file:r.file,error:'没能识别这张图片。',text:''};const items=r.items.sort((a,b)=>(b.y-a.y)||(a.x-b.x));const lines=[];for(const it of items){const last=lines[lines.length-1];if(last&&Math.abs(last.y-it.y)<Math.min(last.h,it.h)*0.5)last.t+=' '+it.t;else lines.push({...it});}return {file:r.file,text:tidyOcr(lines.map(l=>l.t).join('\n'))};});
      return {results,lang,engine:'macOS'};
    }
    const map={zh:'chi_sim+eng',ja:'jpn+eng',en:'eng',auto:'chi_sim+jpn+eng'},results=[];
    for(const f of files){try{results.push({file:f,text:tidyOcr((await exec('tesseract',[f,'stdout','-l',map[lang]||map.auto])).stdout)});}catch{throw Error('这台电脑没有可用的文字识别程序。');}}
    return {results,lang,engine:'tesseract'};
  }
  async function ocr(input){
    const lang=['zh','ja','en','auto'].includes(input.lang)?input.lang:'auto';
    const images=[].concat(input.images||(input.image?[input.image]:[])).slice(0,20);
    if(input.path){return ocrDocument(input,lang);}
    if(!images.length)throw Error('请先粘贴或选择一张图片。');
    const dir=await tempDir();try{
      const files=[];for(let i=0;i<images.length;i++){const buf=Buffer.from(String(images[i]).replace(/^data:[^,]*,/,''),'base64'),t=imageType(buf);if(!t)throw Error('第 '+(i+1)+' 张不是能识别的图片（支持 PNG、JPG、GIF、BMP、TIFF）。');if(buf.length>20*1024*1024)throw Error('图片超过 20 MB，请缩小后再试。');const f=path.join(dir,'img'+(i+1)+'.'+t);await fs.writeFile(f,buf);files.push(f);}
      const r=await ocrFiles(files,lang);const text=r.results.map(x=>x.text).filter(Boolean).join('\n\n');
      if(!text){const err=r.results.find(x=>x.error);throw Error(err?err.error:'图片里没有识别出文字。请换一张更清楚的图片，或把图片放大后再试。');}
      return {text,count:r.results.length,engine:r.engine,lang:r.lang,note:r.note||''};
    }finally{await fs.rm(dir,{recursive:true,force:true});}
  }
  // 识别已保存文档（docx/xlsx/pptx）里的全部图片，按图片顺序返回。
  async function ocrDocument(input,lang){
    const src=await readSource(input,OFFICE_ZIP);const dir=await tempDir();
    try{const media=await mediaOf(src.buf);const pics=media.filter(m=>/\.(png|jpe?g|gif|bmp|tiff?)$/i.test(m.name)).slice(0,40);if(!pics.length)throw Error('文档里没有可以识别的图片（PNG、JPG 等）。如果刚插入了图片，请先保存文档。');
      const files=[];for(let i=0;i<pics.length;i++){const f=path.join(dir,(i+1)+'-'+path.basename(pics[i].name));await fs.writeFile(f,await pics[i].file.async('nodebuffer'));files.push(f);}
      const r=await ocrFiles(files,lang);const parts=r.results.map((x,i)=>({image:path.basename(pics[i].name),text:x.text||''}));
      if(!parts.some(p=>p.text))throw Error('文档里的 '+pics.length+' 张图片都没有识别出文字。');
      return {text:parts.filter(p=>p.text).map(p=>p.text).join('\n\n'),parts,count:pics.length,engine:r.engine,lang:r.lang,note:r.note||''};
    }finally{await fs.rm(dir,{recursive:true,force:true});}
  }

  // ---- 文档里的图片 ----
  async function mediaOf(buf){let zip;try{zip=await JSZip.loadAsync(buf);}catch{throw Error('这个文件不是 docx / xlsx / pptx 格式。旧版 doc、xls、ppt 和 .wps 请先另存为新版格式。');}
    const out=[];zip.forEach((name,file)=>{if(!file.dir&&/^(word|xl|ppt)\/media\//.test(name)&&IMG_EXT.test(name))out.push({name,file});});
    const num=s=>Number((s.match(/(\d+)\.[a-z]+$/i)||[0,0])[1]);out.sort((a,b)=>num(a.name)-num(b.name)||a.name.localeCompare(b.name));return out;}
  async function extractImages(input){
    const src=await readSource(input,[...OFFICE_ZIP,'.pdf']);let items=[];
    if(path.extname(src.name).toLowerCase()==='.pdf'||src.buf.subarray(0,5).toString('latin1')==='%PDF-'){items=await pdfImages(src.buf);}
    else{const media=await mediaOf(src.buf);for(const m of media)items.push({name:path.basename(m.name),data:await m.file.async('nodebuffer')});}
    if(!items.length)throw Error('文件里没有找到图片。');
    const base=path.basename(src.name,path.extname(src.name));const parent=src.file?path.dirname(src.file):await fallbackDir();let dir=path.join(parent,base+'-图片'),n=1;while(await fs.stat(dir).then(()=>true,()=>false))dir=path.join(parent,base+'-图片-'+(++n));
    await fs.mkdir(dir,{recursive:true});let i=0;for(const it of items){i++;await fs.writeFile(path.join(dir,String(i).padStart(3,'0')+'-'+it.name.replace(/[\\/:*?"<>|]/g,'_')),it.data,{flag:'wx'});}
    return {dir,count:items.length};
  }
  async function pdfImages(buf){await ensurePdfEnv();const {PDFParse}=await import('pdf-parse');const pdf=new PDFParse({data:new Uint8Array(buf)});
    try{const r=await pdf.getImage({imageBuffer:true,imageThreshold:40});const out=[];for(const p of r.pages||[])for(const [k,img] of (p.images||[]).entries())if(img.data)out.push({name:'第'+p.pageNumber+'页-'+(k+1)+'.png',data:Buffer.from(img.data),width:img.width,height:img.height,page:p.pageNumber});return out;}
    catch(e){if(/canvas|DOMMatrix|ImageData|Path2D/i.test(String(e&&e.message)))throw Error('这台电脑上读不出 PDF 里的图片（缺少图像组件）。可以在 WPS 里打开 PDF 后右键图片另存。');throw e;}
    finally{await pdf.destroy();}}
  // 压缩：JPG 按质量重新保存并限制最长边；PNG 只在比限制大时缩小。没有变小的图片保持原样。
  async function compressImages(input){
    const src=await readSource(input,OFFICE_ZIP,200);const quality=Math.min(95,Math.max(30,Number(input.quality)||75)),maxSide=Math.min(6000,Math.max(400,Number(input.maxSide)||1920));
    let zip;try{zip=await JSZip.loadAsync(src.buf);}catch{throw Error('这个文件不是 docx / xlsx / pptx 格式，请先另存为新版格式。');}
    const media=[];zip.forEach((name,file)=>{if(!file.dir&&/^(word|xl|ppt)\/media\/[^/]+\.(png|jpe?g)$/i.test(name))media.push({name,file});});
    if(!media.length)throw Error('文件里没有 JPG 或 PNG 图片，不需要压缩。');
    const dir=await tempDir();try{
      const map=new Map();for(let i=0;i<media.length;i++){const f='m'+i+path.extname(media[i].name).toLowerCase();await fs.writeFile(path.join(dir,f),await media[i].file.async('nodebuffer'));map.set(f,media[i]);}
      await shrinkDir(dir,quality,maxSide);
      let before=0,after=0,changed=0;
      for(const [f,m] of map){const orig=await fs.readFile(path.join(dir,f));before+=orig.length;let out=null;try{out=await fs.readFile(path.join(dir,f+'.out'));}catch{}if(out&&out.length>100&&out.length<orig.length*0.95&&imageType(out)===imageType(orig)){zip.file(m.name,out);after+=out.length;changed++;}else after+=orig.length;}
      if(!changed)return {changed:0,count:media.length,before,after,path:null};
      const data=await zip.generateAsync({type:'nodebuffer',compression:'DEFLATE',compressionOptions:{level:6}});
      const p=await writeNew(await outPath(src,'-压缩',path.extname(src.name)),data);
      return {changed,count:media.length,before,after,fileBefore:src.buf.length,fileAfter:data.length,path:p};
    }finally{await fs.rm(dir,{recursive:true,force:true});}
  }
  async function shrinkDir(dir,quality,maxSide){
    if(platform==='win32'){await exec('powershell.exe',['-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-Command',WIN_SHRINK],{env:{SHRINK_DIR:dir,SHRINK_Q:String(quality),SHRINK_W:String(maxSide)},timeout:600000});return;}
    if(platform==='darwin'){for(const f of await fs.readdir(dir)){const p=path.join(dir,f),jpg=/\.jpe?g$/i.test(f);let w=0,h=0;try{const o=(await exec('sips',['-g','pixelWidth','-g','pixelHeight',p])).stdout;w=Number((o.match(/pixelWidth:\s*(\d+)/)||[])[1]);h=Number((o.match(/pixelHeight:\s*(\d+)/)||[])[1]);}catch{continue;}
      const big=Math.max(w,h)>maxSide;if(!jpg&&!big)continue;const args=[];if(big)args.push('-Z',String(maxSide));if(jpg)args.push('-s','formatOptions',String(quality));try{await exec('sips',[...args,p,'--out',p+'.out']);}catch{}}return;}
    throw Error('这台电脑上没有可用的图片压缩程序。');
  }

  // ---- PDF 转 Word ----
  // 按每页文字重排成段落：行尾没有句号、下一行也不是编号开头时接成一段；“一、”“第一章”“1.”这样的短行设成标题。扫描件先做文字识别。
  function paragraphs(text){
    const lines=String(text||'').replace(/\r/g,'').split('\n').map(s=>s.replace(/[ \t ]+$/,'').replace(/^[ \t ]+/,''));
    const wid=t=>{let n=0;for(const ch of t)n+=/[\u1100-\uffff]/.test(ch)?2:1;return n;},maxw=Math.max(16,...lines.map(wid)),full=maxw*0.75,out=[];let cur=null;
    const heading=t=>{if(t.length>40||/[。！？；，,!?;]$/.test(t))return 0;if(/^第[一二三四五六七八九十百\d０-９]+[章节部分篇節]/.test(t)||/^[一二三四五六七八九十]+[、．.]\s*\S/.test(t))return 1;if(/^[（(][一二三四五六七八九十]+[)）]\s*\S/.test(t))return 2;if(/^\d{1,2}(\.\d{1,2}){0,2}[、．.\s]\s*\S/.test(t)&&t.length<=30)return /^\d{1,2}\.\d/.test(t)?3:2;return 0;};
    // “收件人：…”“宛先：…”这类字段行，各自成段。
    const starts=t=>/^([一二三四五六七八九十]+[、．.]|[（(][一二三四五六七八九十\d]+[)）]|\d{1,3}[、．.)）]|[•●■◆\-–—*·]\s|第[一二三四五六七八九十百\d]+[章节条])/.test(t)||field(t);
    const field=t=>/^[^：:。，、！？!?]{1,12}[：:]\s*\S/.test(t);
    for(const t of lines){if(!t){cur=null;continue;}const h=heading(t);if(h){out.push({text:t,level:h});cur=null;continue;}
      if(cur&&!starts(t)&&!/[。！？!?：:；;」』）)]$/.test(cur.last)&&wid(cur.last)>=(field(cur.last)?maxw*0.92:full)){const glue=/[A-Za-z0-9,]$/.test(cur.last)&&/^[A-Za-z0-9]/.test(t)?' ':'';cur.text+=glue+t;cur.last=t;continue;}
      cur={text:t,last:t,level:0};out.push(cur);}
    return out.map(p=>({text:p.text,level:p.level}));
  }
  // 表格：PDF 里画了框线的表格按格子取出来；页面文字里属于表格的那几行去掉，在原来的位置换成 Word 表格。
  const squash=t=>String(t||'').replace(/\s+/g,'');
  function withTables(text,list){
    const lines=String(text||'').replace(/\r/g,'').split('\n'),segs=[];if(!list.length)return [{text:lines.join('\n')}];
    const used=new Set(),marks=new Map();
    for(const tb of list){const all=squash(tb.map(r=>r.join('')).join(''));if(all.length<2)continue;
      for(let i=0;i<lines.length;i++){if(used.has(i))continue;let j=i,cov=0;while(j<lines.length&&!used.has(j)){const l=squash(lines[j]);if(!l){j++;continue;}if(!all.includes(l))break;cov+=l.length;j++;}
        if(j>i&&cov>=all.length*0.6){for(let k=i;k<j;k++)used.add(k);marks.set(i,tb);break;}}}
    let buf=[];const flush=()=>{if(buf.length)segs.push({text:buf.join('\n')});buf=[];};
    for(let i=0;i<lines.length;i++){if(marks.has(i)){flush();segs.push({table:marks.get(i)});}if(!used.has(i))buf.push(lines[i]);else if(!marks.has(i)&&buf.length)flush();}
    flush();
    // 没能在文字里对上位置的表格放到这一页最后（不丢内容）。
    for(const tb of list)if(![...marks.values()].includes(tb))segs.push({table:tb});
    return segs;
  }
  function wordTable(rows){const n=Math.max(...rows.map(r=>r.length));
    return new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:rows.map((r,ri)=>new TableRow({tableHeader:ri===0,children:Array.from({length:n},(_,ci)=>new TableCell({children:String(r[ci]==null?'':r[ci]).split('\n').map(t=>new Paragraph({children:[new TextRun({text:t.trim(),bold:ri===0})]}))}))}))});}
  async function pdfToWord(input){
    const src=await readSource(input,['.pdf'],80);if(src.buf.subarray(0,5).toString('latin1')!=='%PDF-')throw Error('这不是有效的 PDF 文件。');
    await ensurePdfEnv();const {PDFParse}=await import('pdf-parse');const pdf=new PDFParse({data:new Uint8Array(src.buf)});
    let pages=[],ocred=0,images=[],notes=[],tables=new Map();
    try{let r;try{r=await pdf.getText();}catch(e){if(/password|encrypt/i.test(String(e&&e.message)))throw Error('这个 PDF 有密码保护，请先在 PDF 阅读器里解除密码再转换。');throw e;}
      pages=(r.pages||[]).map(p=>({num:p.num,text:String(p.text||'')}));
      if(pages.length>500)throw Error('PDF 超过 500 页，请先拆分。');
      const chars=pages.reduce((n,p)=>n+p.text.replace(/\s/g,'').length,0);
      if(chars<Math.max(20,pages.length*15)){const shots=await pdfPageImages(pdf,src,Math.min(pages.length||30,30));if(!shots.length)throw Error('这个 PDF 是扫描件（图片），这台电脑上没法把页面转成图片做文字识别。');
        const dir=await tempDir();try{const files=[];for(let i=0;i<shots.length;i++){const f=path.join(dir,'p'+(i+1)+'.png');await fs.writeFile(f,shots[i]);files.push(f);}const o=await ocrFiles(files,['zh','ja','en'].includes(input.lang)?input.lang:'auto');pages=o.results.map((x,i)=>({num:i+1,text:x.text||'',layout:x.layout||null}));ocred=files.length;if(o.note)notes.push(o.note);}finally{await fs.rm(dir,{recursive:true,force:true});}
        if(pages.length<(r.pages||[]).length)notes.push('扫描件一次最多识别 30 页，后面的页没有转换。');}
      else{try{const t=await pdf.getTable();for(const pg of t.pages||[])tables.set(pg.num,(pg.tables||[]).filter(tb=>Array.isArray(tb)&&tb.length>=2&&tb.some(r=>r.length>=2)));}catch{notes.push('没能识别 PDF 里的表格，表格会变成一行一行的文字。');}}
      if(!ocred&&input.images!==false){try{images=await pdfImages(src.buf);}catch{notes.push('PDF 里的图片没能取出来，只转换了文字。');}images=images.filter(im=>im.width>=60&&im.height>=60).slice(0,80);}
    }finally{await pdf.destroy();}
    const children=[];let count=0,tableCount=0;
    const para=x=>{count++;const align=x.align==='right'?AlignmentType.RIGHT:x.level===1&&x.center!==false&&x.fromScan?AlignmentType.CENTER:undefined;children.push(new Paragraph(x.level?{text:x.text,heading:[HeadingLevel.HEADING_1,HeadingLevel.HEADING_2,HeadingLevel.HEADING_3][x.level-1],alignment:align}:{children:[new TextRun(x.text)],spacing:{after:120},alignment:align}));};
    for(const p of pages){
      // 扫描件：按版面还原的结果写——表格还原成 Word 表格，标题居中、落款靠右、填写栏单独一段。
      if(p.layout){for(const seg of p.layout){if(seg.table){tableCount++;children.push(wordTable(seg.table));children.push(new Paragraph({text:''}));}else if(seg.para)para({...seg.para,fromScan:true});else paragraphs(seg.text).forEach(para);}continue;}
      for(const seg of withTables(p.text,tables.get(p.num)||[])){if(seg.table){tableCount++;children.push(wordTable(seg.table));children.push(new Paragraph({text:''}));}else paragraphs(seg.text).forEach(para);}
      for(const im of images.filter(i=>i.page===p.num)){const w=Math.min(600,im.width),h=Math.round(im.height*w/im.width);children.push(new Paragraph({children:[new ImageRun({type:'png',data:im.data,transformation:{width:w,height:h}})]}));}}
    if(!count&&!tableCount)throw Error('PDF 里没有找到文字。');
    const doc=new Document({styles:{default:{document:{run:{font:input.lang==='ja'?'Yu Gothic':'Microsoft YaHei',size:21}}}},sections:[{children}]});
    const p=await writeNew(await outPath(src,'','.docx'),await Packer.toBuffer(doc));
    return {path:p,pages:pages.length,paragraphs:count,tables:tableCount,images:images.length,ocr:ocred,notes};
  }
  async function pdfPageImages(pdf,src,limit){
    // 2026-11-20 真机：按 2 倍画，Windows 文字识别认不出加粗的表头（如“规格”）；画到约 2400 像素宽（A4 约 4 倍）就能认出，一页约 1 秒。
    try{const s=await pdf.getScreenshot({desiredWidth:2400,imageBuffer:true,last:limit});const out=(s.pages||[]).slice(0,limit).map(p=>Buffer.from(p.data)).filter(b=>b.length);if(out.length)return out;}catch{}
    if(platform==='darwin'){const dir=await tempDir();try{let file=src.file;if(!file){file=path.join(dir,'in.pdf');await fs.writeFile(file,src.buf);}const d=JSON.parse((await exec('osascript',['-l','JavaScript','-e',MAC_PDF,file,dir,String(limit)])).stdout.trim());const out=[];for(const f of d.files||[])out.push(await fs.readFile(f));return out;}catch{return [];}finally{await fs.rm(dir,{recursive:true,force:true});}}
    return [];
  }

  // ---- 历史版本：备份文件夹里“文件名-备份-时间”的副本 ----
  async function history(input,backupDir){
    const file=String(input.path||'');if(!path.isAbsolute(file))throw Error('请先把文档保存到电脑上。');
    const ext=path.extname(file),base=path.basename(file,ext),prefix=base+'-备份-';let names=[];try{names=await fs.readdir(backupDir);}catch{}
    const list=[];for(const n of names){if(!n.startsWith(prefix)||path.extname(n).toLowerCase()!==ext.toLowerCase())continue;const m=n.slice(prefix.length).match(/^(\d{4})(\d{2})(\d{2})-(\d{2})(\d{2})(\d{2})/);if(!m)continue;const st=await fs.stat(path.join(backupDir,n)).catch(()=>null);if(!st)continue;list.push({path:path.join(backupDir,n),name:n,size:st.size,time:m[1]+'-'+m[2]+'-'+m[3]+' '+m[4]+':'+m[5]+':'+m[6],auto:/-自动/.test(n)});}
    list.sort((a,b)=>b.time.localeCompare(a.time));return {items:list.slice(0,100),dir:backupDir};
  }
  // 自动备份只保留每个文件最近 20 份，多的删掉（只删本工具自动建的“-自动”副本）。
  async function pruneAuto(file,backupDir,keep=20){const ext=path.extname(file),prefix=path.basename(file,ext)+'-备份-';let names=[];try{names=await fs.readdir(backupDir);}catch{return 0;}const autos=names.filter(n=>n.startsWith(prefix)&&n.includes('-自动')&&path.extname(n).toLowerCase()===ext.toLowerCase()).sort().reverse();let removed=0;for(const n of autos.slice(keep)){await fs.rm(path.join(backupDir,n),{force:true});removed++;}return removed;}
  async function reveal(input){const p=String(input.path||'');if(!path.isAbsolute(p))throw Error('位置无效。');const st=await fs.stat(p).catch(()=>null);if(!st)throw Error('找不到这个位置。');
    if(platform==='win32')await exec('explorer.exe',st.isDirectory()?[p]:['/select,',p]).catch(()=>{});else if(platform==='darwin')await exec('open',st.isDirectory()?[p]:['-R',p]);else await exec('xdg-open',[st.isDirectory()?p:path.dirname(p)]).catch(()=>{});return {ok:true};}
  return {ocr,ocrFiles,extractImages,compressImages,pdfToWord,paragraphs,withTables,history,pruneAuto,reveal,mediaOf};
}
