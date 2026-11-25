import fs from 'node:fs';
import path from 'node:path';
import {createRequire} from 'node:module';
const fontkit=createRequire(import.meta.url)('fontkit');
export const MISSING_FONT_EXIT=78;
export function fontCandidates(lang,bold=false,{platform=process.platform,env=process.env}={}){
 const dir=env.WINDIR?path.join(env.WINDIR,'Fonts'):'';
 const chinese=platform==='win32'?[[path.join(dir,bold?'msyhbd.ttc':'msyh.ttc'),bold?'MicrosoftYaHei-Bold':'MicrosoftYaHei','Microsoft YaHei']]:[
  ['/System/Library/Fonts/PingFang.ttc',bold?'PingFangSC-Semibold':'PingFangSC-Regular','PingFang SC'],
  ['/System/Library/Fonts/Hiragino Sans GB.ttc',bold?'HiraginoSansGB-W6':'HiraginoSansGB-W3','Hiragino Sans GB'],
  ['/System/Library/Fonts/STHeiti Medium.ttc','STHeitiSC-Medium','Heiti SC']
 ];
 const candidates=platform==='win32'?(lang==='ja'?[[path.join(dir,bold?'YuGothB.ttc':'YuGothM.ttc'),bold?'YuGothic-Bold':'YuGothic-Medium','Yu Gothic'],...chinese]:chinese):platform==='darwin'?(lang==='ja'?[
  [bold?'/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc':'/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc',bold?'HiraginoSans-W6':'HiraginoSans-W3','Hiragino Sans'],...chinese
 ]:chinese):[];
 if(env.MANUAL_FONT)candidates.unshift([env.MANUAL_FONT,undefined,env.MANUAL_FONT_NAME||'Manual CJK']);
 return candidates;
}
export function fontSpec(lang,bold=false,options={}){
 const exists=options.exists||fs.existsSync,open=options.open||fontkit.openSync;
 for(const [file,face,name]of fontCandidates(lang,bold,options)){
  if(!exists(file))continue;
  try{const loaded=open(file,face),font=loaded.fonts?loaded.fonts[0]:loaded;return {file,face:/\.ttc$/i.test(file)?face||font.postscriptName:undefined,name,font};}catch{}
 }
 throw Object.assign(Error('No suitable local CJK font for '+lang+'. Install Microsoft YaHei/Yu Gothic or PingFang/Hiragino, or set MANUAL_FONT and MANUAL_FONT_NAME.'),{code:'MANUAL_FONT_MISSING'});
}
