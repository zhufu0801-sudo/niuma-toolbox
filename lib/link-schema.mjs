import path from 'node:path';
export function checkLinkAnchor(value,source=false){
 if(!value||!['wps','et','wpp'].includes(value.host)||typeof value.path!=='string'||value.path.length>1024||/[\x00-\x1f]/.test(value.path)||!(path.isAbsolute(value.path)||path.win32.isAbsolute(value.path)))throw Error('请先保存文档，再建立联动。');
 if(typeof value.mark!=='string'||!/^OT[SD]_[A-Za-z0-9]{1,32}$/.test(value.mark))throw Error('联动定位标记无效。');
 const result={host:value.host,path:value.path,mark:value.mark};
 if(value.host==='wpp'){
  const ids=source?[value.slideId]:value.slides;if(!Array.isArray(ids)||!ids.length||ids.length>30||ids.some(id=>!Number.isSafeInteger(id)||id<1)||new Set(ids).size!==ids.length)throw Error('幻灯片定位无效。');
  if(source){if(!Number.isSafeInteger(value.shapeId)||value.shapeId<1)throw Error('形状定位无效。');result.slideId=value.slideId;result.shapeId=value.shapeId;}else result.slides=ids.slice();
 }return result;
}
export function checkLinkSnapshot(value){if(typeof value!=='string'||value.length>350000)throw Error('联动内容过长，请缩小选区。');let data;try{data=JSON.parse(value);}catch{throw Error('联动快照无法识别。');}if(!data||typeof data!=='object'||Array.isArray(data))throw Error('联动快照无法识别。');return value;}
export function conflict(message){const error=Error(message);error.status=409;return error;}
