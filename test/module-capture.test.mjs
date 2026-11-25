import test from 'node:test';import assert from 'node:assert/strict';import vm from 'node:vm';import {toolsCode} from './helpers/tools-code.mjs';
test('Word module captures styles for selected text and body does not inherit heading level',async()=>{let saved;const font={Name:'Arial',NameFarEast:'宋体',Size:12,Bold:0,Italic:0,Color:0},paragraph=(level)=>({Range:{Font:font,ParagraphFormat:{Alignment:0},Paragraphs:{Item:()=>({OutlineLevel:level})}}});const range={Font:font,ParagraphFormat:{Alignment:0},Paragraphs:{Item:i=>paragraph(i===1?1:10)}};const root={Application:{ActiveDocument:{Name:'test.docx'},Selection:{Text:'标题\r正文',Range:range,Tables:{Count:0}}}};vm.runInNewContext(toolsCode(),{window:root,location:{pathname:'/addon/wps/',search:''},URLSearchParams,setTimeout,clearTimeout,fetch:async(url,options)=>{if(url==='/api/session')return Response.json({token:'t'});saved=JSON.parse(options.body);return Response.json({card:{title:'测试模块'}});}});await root.OfficeTools.run('saveCard');assert.equal(saved.block.text,'标题\n正文');assert.equal(saved.format.paragraphs[0].level,1);assert.equal(saved.format.paragraphs[1].level,0);assert.equal(saved.format.base.font,'宋体');});
test('Word mixed-format selection does not turn the undefined sentinel 9999999 into a colour or give every paragraph a heading level',async()=>{let saved;
  // 真机 0.6.2：标题+正文混选时 WPS 对整个选区返回 9999999（未定义），曾被存成灰色 #7f9698，插到 Excel 后正文几乎看不见。
  const plain={Name:'宋体',NameFarEast:'宋体',Size:11,Bold:0,Italic:0,Color:-16777216},heading={Name:'黑体',NameFarEast:'黑体',Size:14,Bold:-1,Italic:0,Color:255};
  const para=(font,level)=>({OutlineLevel:level,Range:{Font:font,ParagraphFormat:{Alignment:0},Paragraphs:{Item:()=>({OutlineLevel:level})}}});
  const range={Font:{Name:'',NameFarEast:'',Size:9999999,Bold:9999999,Italic:0,Color:9999999},ParagraphFormat:{Alignment:0},Paragraphs:{Item:i=>{if(i>3)throw Error('no paragraph');return i===1?para(heading,1):para(plain,10);}}};
  const root={Application:{ActiveDocument:{Name:'test.docx'},Selection:{Text:'标题\r第一段\r第二段\r',Range:range,Tables:{Count:0}}}};
  vm.runInNewContext(toolsCode(),{window:root,location:{pathname:'/addon/wps/',search:''},URLSearchParams,setTimeout,clearTimeout,fetch:async(url,options)=>{if(url==='/api/session')return Response.json({token:'t'});saved=JSON.parse(options.body);return Response.json({card:{title:'测试模块'}});}});
  await root.OfficeTools.run('saveCard');const f=saved.format;
  assert.equal(f.base.color,undefined);assert.equal(f.base.level,undefined);
  assert.equal(f.paragraphs[0].level,1);assert.equal(f.paragraphs[0].color,'#ff0000');assert.equal(f.paragraphs[1].color,undefined);
  // 末尾空段落没有对应段落对象时，不再回退成整个选区的格式。
  assert.equal(JSON.stringify(f.paragraphs[f.paragraphs.length-1]),'{}');});
