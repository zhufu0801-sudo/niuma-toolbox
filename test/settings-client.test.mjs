import test from 'node:test';
import assert from 'node:assert/strict';
import {createClock,createDocument,createSharedBrowser,loadScripts,flush} from './helpers/frontend-fakes.mjs';
import {TOOL_MODULES} from './helpers/tools-code.mjs';
import {defaultSettings} from '../lib/settings.mjs';
test('legacy switches migrate to server once; changing settings reaches ribbon and formatting',async()=>{
 const clock=createClock(),browser=createSharedBrowser(),document=createDocument();let state={settings:defaultSettings(),defaults:defaultSettings(),revision:0,initialized:false},migrations=0;
 browser.localStorage.setItem('officeTools.hidden',JSON.stringify({ids:['heading1'],at:5}));
 const selection={Text:'正文',Range:{},Font:{},ParagraphFormat:{}},a={Selection:selection,ActiveDocument:{Name:'test.docx'}};
 const root=loadScripts({...clock,...browser,document,URLSearchParams,Application:a,location:{pathname:'/',search:'?host=wps'},fetch:async(url,opts={})=>{if(url==='/api/session')return {ok:true,json:async()=>({token:'t'})};const body=opts.body&&JSON.parse(opts.body);if(url==='/api/settings/migrate'){migrations++;state.settings.hiddenIds=body.ids;state.initialized=true;state.revision++;}else if(opts.method==='POST'){Object.assign(state.settings,body.patch);state.revision++;}return {ok:true,json:async()=>structuredClone(state)};}},['settings-client.js',...TOOL_MODULES]);
 await root.OfficeTools.ready();assert.equal(migrations,1);assert.equal(root.OfficeTools.isHidden('heading1'),true);
 await root.OfficeTools.setHidden(['toc']);assert.equal(root.OfficeTools.isHidden('toc'),true);assert.equal(root.OfficeTools.isHidden('heading1'),false);
 state.settings.text={font:'黑体',size:15,lineSpacing:2};state.revision++;await clock.advance(2000);await root.OfficeTools.run('bodyFormat');assert.equal(selection.Font.NameFarEast,'黑体');assert.equal(selection.Font.Size,15);assert.equal(selection.ParagraphFormat.LineSpacing,24);assert.equal(migrations,1);
});
test('start page disables launch on version mismatch and reports service errors',async()=>{
 for(const match of [true,false]){const document=createDocument(),calls=[];loadScripts({document,fetch:async(url,options={})=>{calls.push([url,options.method]);const data=url==='/api/session'?{token:'t'}:url==='/api/health'?{app:'wps-ai-coach',version:'test'}:url==='/version.json'?{version:match?'test':'different'}:url==='/api/startup'?{wpsFound:true,cardsCount:3,backupDir:'C:/backup'}:{ok:true};return {ok:true,json:async()=>data};}},['start.js']);await flush();assert.equal(document.getElementById('startBtn').disabled,!match);assert.equal(document.getElementById('versionBanner').hidden,match);if(match){document.getElementById('startBtn').click();await flush();assert.ok(calls.some(([url,method])=>url==='/api/launch-wps'&&method==='POST'));}}
});
