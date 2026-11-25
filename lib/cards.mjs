// 内容卡片库：把文字、表格存成一张张卡片，在文字、表格、演示之间反复取用。
// 只保存在本机 generated/cards.json，重启电脑后仍在；最多保留 40 张，超出时丢掉最旧的。
// file 为空时只放在内存里（测试用），不会写入真实卡片库。
import fs from 'node:fs/promises';
import path from 'node:path';
import { randomBytes } from 'node:crypto';
import { validateBlocks } from './workspace.mjs';
import {moduleMetadata} from './module-format.mjs';
import {checkLinkAnchor,checkLinkSnapshot,conflict} from './link-schema.mjs';

export const CARD_LIMIT = 40;
const SOURCES = ['wps', 'et', 'wpp', 'drop'];

export function checkCardBlock(input) {
  if (input && input.type === 'text' && !String(input.text ?? '').trim()) throw new Error('卡片内容是空的。');
  const block = validateBlocks([input])[0];
  if (block.type === 'text') {
    if (!block.text.trim()) throw new Error('卡片内容是空的。');
    if (block.text.length > 10000) throw new Error('一张卡片最多 1 万字，请分成几张。');
  } else if (block.rows.length > 20 || block.rows[0].length > 10) throw new Error('卡片里的表格最多 20 行、10 列。');
  return block;
}

export function cardTitle(block) {
  if (block.type === 'table') return `${block.rows.length} 行 × ${block.rows[0].length} 列表格：${block.rows[0].filter(Boolean).join('、').slice(0, 24)}`;
  const line = block.text.trim().split(/\n/)[0];
  return line.length > 30 ? line.slice(0, 30) + '…' : line;
}

export function createCardStore(file,{getLimit=()=>CARD_LIMIT}={}) {
  let cards = null,queue=Promise.resolve();
  const serial=fn=>{const result=queue.catch(()=>{}).then(fn);queue=result;return result;};
  async function load() {
    if (cards) return cards;
    if (!file) return (cards = []);
    try { const data = JSON.parse(await fs.readFile(file, 'utf8')); if(!Array.isArray(data.cards))throw Error('卡片结构不正确');cards=data.cards; }
    catch(e) { if(e.code==='ENOENT')cards=[];else throw Error('卡片库无法读取，已保留原文件，请先备份后检查。'); }
    return cards;
  }
  async function save(next) {
    if (!file) return;
    await fs.mkdir(path.dirname(file), { recursive: true });
    const temp = file + '.' + randomBytes(4).toString('hex') + '.tmp';
    await fs.writeFile(temp, JSON.stringify({ cards:next }, null, 1), 'utf8');
    await fs.rename(temp, file);
  }
  return {
    list() { return serial(async()=>(await load()).slice()); },
    add(input, sourceHost,metadata,source) {return serial(async()=>{
      const block = checkCardBlock(input);
      if (!SOURCES.includes(sourceHost)) throw new Error('来源无效。');
      const list = (await load()).slice();
      const card = { id: Date.now().toString(36) + randomBytes(3).toString('hex'), block, ...moduleMetadata(block,metadata??input.format), sourceHost, title: cardTitle(block), created: Date.now(),revision:1 };
      if(source){card.source=checkLinkAnchor(source,true);if(card.source.host!==sourceHost)throw Error('联动来源类型不一致。');}
      list.unshift(card);
      const limit=await getLimit();while(list.length>limit){let index=list.length-1;while(index>=0&&list[index].source)index--;if(index<=0)throw Error('卡片上限已被联动模块占满，请提高上限或手动删除不需要的模块。');list.splice(index,1);}
      await save(list);cards=list;
      return card;
    });},
    refresh(id,input){return serial(async()=>{
      const list=(await load()).slice(),index=list.findIndex(c=>c.id===id),old=list[index];if(!old||!old.source)throw Error('联动来源已不存在。');if(input.sourceMark!==old.source.mark)throw Error('来源定位不匹配。');
      if(input.expectedRevision!==(old.revision||1))throw conflict('来源已在另一个窗口刷新，请重试。');
      const block=checkCardBlock(input.block),metadata=moduleMetadata(block,input.format);
      if(JSON.stringify({block,format:metadata.format})===JSON.stringify({block:old.block,format:old.format}))return old;
      const next={...old,block,title:cardTitle(block),revision:(old.revision||1)+1,updated:Date.now()};delete next.format;delete next.html;Object.assign(next,metadata);list[index]=next;await save(list);cards=list;return next;
    });},
    attach(id,input){return serial(async()=>{
      const list=(await load()).slice(),index=list.findIndex(c=>c.id===id),card=list[index];if(!card?.source)throw Error('这个模块没有联动来源，请先在来源文档中“设为联动模块”。');
      const anchor=checkLinkAnchor(input.anchor),baseline=checkLinkSnapshot(input.baseline);if(!Number.isSafeInteger(input.moduleRevision)||input.moduleRevision<1||input.moduleRevision>(card.revision||1))throw Error('模块版本无效。');
      if((card.targets||[]).length>=100)throw Error('一个模块最多联动 100 个位置。');if((card.targets||[]).some(link=>link.anchor.mark===anchor.mark))throw conflict('这个位置已经建立联动。');
      const link={id:randomBytes(10).toString('hex'),anchor,baseline,moduleRevision:input.moduleRevision,revision:1,created:Date.now()};list[index]={...card,targets:[...(card.targets||[]),link]};await save(list);cards=list;return link;
    });},
    acknowledge(id,linkId,input){return serial(async()=>{
      const list=(await load()).slice(),index=list.findIndex(c=>c.id===id),card=list[index],targets=card?.targets?.slice(),at=targets?.findIndex(t=>t.id===linkId);if(at===undefined||at<0)throw Error('这个联动位置已不存在。');const old=targets[at];
      if(input.expectedRevision!==old.revision)throw conflict('另一个窗口已经更新过这个位置，请重新检查。');
      if(!Number.isSafeInteger(input.moduleRevision)||input.moduleRevision<old.moduleRevision||input.moduleRevision>(card.revision||1))throw Error('模块版本无效。');
      targets[at]={...old,baseline:checkLinkSnapshot(input.baseline),moduleRevision:input.moduleRevision,revision:old.revision+1,updated:Date.now()};list[index]={...card,targets};await save(list);cards=list;return targets[at];
    });},
    detach(id,linkId){return serial(async()=>{const list=(await load()).slice(),index=list.findIndex(c=>c.id===id);if(index<0)throw Error('模块已不存在。');list[index]={...list[index],targets:(list[index].targets||[]).filter(t=>t.id!==linkId)};await save(list);cards=list;});},
    // 数据脱敏（2026-11-25）：换成脱敏后的内容（或还原回原文）。格式（字体、对齐等）保留，html 快照按新内容重新生成（旧快照里有原文）。
    replaceBlock(id,input){return serial(async()=>{
      const list=(await load()).slice(),index=list.findIndex(c=>c.id===id),old=list[index];if(!old)throw Error('这张卡片已经不存在了。');
      const block=checkCardBlock(input),next={...old,block,title:cardTitle(block),revision:(old.revision||1)+1,updated:Date.now()};delete next.format;delete next.html;let meta={};try{meta=moduleMetadata(block,old.format);}catch{meta={};}Object.assign(next,meta);list[index]=next;await save(list);cards=list;return next;
    });},
    remove(id) {return serial(async()=>{
      const list = (await load()).slice(), index = list.findIndex(c => c.id === id);
      if (index < 0) return false;
      list.splice(index, 1); await save(list);cards=list; return true;
    });},
    clear() {return serial(async()=>{await load();await save([]);cards=[];});},
    merge(input){return serial(async()=>{if(!Array.isArray(input)||input.length>200)throw Error('一次最多导入 200 张卡片。');const list=(await load()).slice(),known=new Set(list.map(c=>c.id));let added=0,skipped=0;const incoming=[];for(const entry of input){if(!entry||!SOURCES.includes(entry.sourceHost))throw Error('导入卡片的来源无法识别。');const block=checkCardBlock(entry.block);const id=typeof entry.id==='string'&&/^[a-zA-Z0-9_-]{1,80}$/.test(entry.id)?entry.id:Date.now().toString(36)+randomBytes(6).toString('hex');if(known.has(id)){skipped++;continue;}known.add(id);incoming.push({id,block,...moduleMetadata(block,entry.format),sourceHost:entry.sourceHost,title:cardTitle(block),created:Number.isFinite(entry.created)?entry.created:Date.now()});added++;}if(list.length+incoming.length>await getLimit())throw Error('导入后会超过卡片上限，请先调整上限或减少导入数量；已有卡片未修改。');const next=[...incoming,...list];await save(next);cards=next;return {added,skipped,total:next.length};});},
  };
}
