// 简化的 WPS 表格对象模型，只实现“表格整理”工具用到的部分。单元格按 (行,列) 存在 Map 里。
export function colName(n){let s='';while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;}
function parseA1(a){const m=String(a).replace(/\$/g,'').match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);const c=x=>[...x].reduce((s,ch)=>s*26+ch.charCodeAt(0)-64,0);return {r1:+m[2],c1:c(m[1]),r2:+(m[4]||m[2]),c2:c(m[3]||m[1])};}
export function makeBook(sheets){
  const book={Name:'test.xlsx',Path:'C:\\tmp',FullName:'C:\\tmp\\test.xlsx',ReadOnly:false,Saved:true,list:[]};
  function sheet(name,rows){
    const cells=new Map(),ws={Name:name,ProtectContents:false,cells,merges:[],cf:[],PageSetup:{},charts:[]};
    const key=(r,c)=>r+','+c;
    function cell(r,c){if(!cells.has(key(r,c)))cells.set(key(r,c),{Value2:undefined,HasFormula:false,Formula:'',NumberFormat:'General',Interior:{Color:16777215,ColorIndex:-4142},Font:{},Text:''});return cells.get(key(r,c));}
    ws.cell=cell;
    (rows||[]).forEach((row,i)=>row.forEach((v,j)=>{if(v!==undefined){const x=cell(i+1,j+1);if(typeof v==='string'&&v.startsWith('=')){x.HasFormula=true;x.Formula=v;x.Value2=0;}else x.Value2=v;x.Text=v==null?'':String(v);}}));
    function mergeAt(r,c){return ws.merges.find(m=>r>=m.r1&&r<=m.r2&&c>=m.c1&&c<=m.c2);}
    function range(r1,c1,r2,c2){
      const R={Row:r1,Column:c1,Rows:null,Columns:null,Worksheet:ws};
      R.Rows={Count:r2-r1+1,Item:i=>range(r1+i-1,c1,r1+i-1,c2)};R.Columns={Count:c2-c1+1,Item:j=>range(r1,c1+j-1,r2,c1+j-1)};
      R.Cells={Item:(i,j)=>{const x=cell(r1+i-1,c1+j-1),rr=r1+i-1,cc=c1+j-1;const o={get Value2(){return x.Value2;},set Value2(v){x.Value2=typeof v==='string'&&v.startsWith("'")?v.slice(1):v;x.HasFormula=false;x.Formula='';x.Text=x.Value2==null?'':String(x.Value2);},get HasFormula(){return x.HasFormula;},get Formula(){return x.Formula||x.Value2;},set Formula(f){x.Formula=f;x.HasFormula=true;x.Value2=0;},get NumberFormat(){return x.NumberFormat;},set NumberFormat(v){x.NumberFormat=v;},Interior:x.Interior,Font:x.Font,get Text(){return x.Text;},Row:rr,Column:cc,
        get MergeArea(){const m=mergeAt(rr,cc);return m?range(m.r1,m.c1,m.r2,m.c2):range(rr,cc,rr,cc);},Address:(a,b)=>(a===false?'':'$')+colName(cc)+(a===false?'':'$')+rr,Copy:target=>{for(let k=1;k<=target.Rows.Count;k++){const t=cell(target.Row+k-1,target.Column);t.Formula=x.Formula.replace(/([A-Z]+)(\d+)(?![\d:])/g,(m0,cn,rn,off,str)=>str[off-1]==='$'?m0:cn+(+rn+k-1));t.HasFormula=true;}}};return o;}};
      Object.defineProperty(R,'Value2',{get(){if(r1===r2&&c1===c2)return cell(r1,c1).Value2;const out=[];for(let r=r1;r<=r2;r++){const line=[];for(let c=c1;c<=c2;c++)line.push(cell(r,c).Value2);out.push(line);}return out;},set(v){if(!Array.isArray(v)){R.Cells.Item(1,1).Value2=v;return;}v.forEach((row,i)=>row.forEach((x,j)=>{R.Cells.Item(i+1,j+1).Value2=x;}));}});
      Object.defineProperty(R,'NumberFormat',{get(){return cell(r1,c1).NumberFormat;},set(v){for(let r=r1;r<=r2;r++)for(let c=c1;c<=c2;c++)cell(r,c).NumberFormat=v;}});
      Object.defineProperty(R,'HasFormula',{get(){let any=false,all=true;for(let r=r1;r<=r2;r++)for(let c=c1;c<=c2;c++){const f=cell(r,c).HasFormula;any=any||f;all=all&&f;}return all?true:any?null:false;}});
      Object.defineProperty(R,'MergeCells',{get(){return ws.merges.some(m=>m.r1<=r2&&m.r2>=r1&&m.c1<=c2&&m.c2>=c1)?(r1===r2&&c1===c2?true:null):false;}});
      R.Address=(a,b)=>{const p=a===false?'':'$';const s=p+colName(c1)+p+r1;return r1===r2&&c1===c2?s:s+':'+p+colName(c2)+p+r2;};
      R.Offset=(dr,dc)=>range(r1+dr,c1+dc,r2+dr,c2+dc);R.Resize=(nr,nc)=>range(r1,c1,r1+nr-1,c1+nc-1);
      R.FormatConditions={get Count(){return ws.cf.length;},Item:i=>{const x=ws.cf[i-1];return {Formula1:x.formula,Delete(){ws.cf.splice(ws.cf.indexOf(x),1);}};},Add:(type,op,formula)=>{const x={area:R.Address(),formula,get Formula1(){return this.formula;},Delete(){ws.cf.splice(ws.cf.indexOf(x),1);},Interior:{},Font:{}};ws.cf.push(x);return x;},AddDatabar:()=>{const x={area:R.Address(),databar:true,BarColor:{},MinPoint:{Modify(){}},MaxPoint:{Modify(){}}};ws.cf.push(x);return x;}};
      R.Interior={set Color(v){for(let r=r1;r<=r2;r++)for(let c=c1;c<=c2;c++)cell(r,c).Interior.Color=v;}};R.Font={};R.Borders={};
      R.Delete=()=>{for(let r=r1;r<=r2;r++){}const n=r2-r1+1,moved=new Map();for(const [k,v] of cells){const [r,c]=k.split(',').map(Number);if(r<r1)moved.set(k,v);else if(r>r2)moved.set(key(r-n,c),v);}cells.clear();for(const [k,v] of moved)cells.set(k,v);};
      R.Left=c1*50;R.Top=r1*20;R.Width=(c2-c1+1)*50;
      return R;
    }
    ws.Range=a=>{if(typeof a==='object')return a;const p=parseA1(a);return range(p.r1,p.c1,p.r2,p.c2);};
    ws.Cells={Item:(r,c)=>range(r,c,r,c).Cells.Item(1,1),Clear(){cells.clear();},get FormatConditions(){return {get Count(){return ws.cf.length;},Item:i=>{const x=ws.cf[i-1];return {Formula1:x.formula,Delete(){ws.cf.splice(ws.cf.indexOf(x),1);}};}};}};
    const origRange=ws.Range;ws.Range=(a,b)=>{if(b){return range(a.Row,a.Column,b.Row,b.Column);}return origRange(a);};
    Object.defineProperty(ws,'UsedRange',{get(){let r1=Infinity,c1=Infinity,r2=1,c2=1,any=false;for(const [k,v] of cells){if(v.Value2===undefined||v.Value2===''||v.Value2===null){if(!v.HasFormula)continue;}const [r,c]=k.split(',').map(Number);any=true;r1=Math.min(r1,r);c1=Math.min(c1,c);r2=Math.max(r2,r);c2=Math.max(c2,c);}return any?range(r1,c1,r2,c2):range(1,1,1,1);}});
    ws.Rows={Item:r=>range(r,1,r,30)};ws.Columns={Item:c=>range(1,c,1048576,c),AutoFit(){}};
    ws.Activate=()=>{book.active=ws;};ws.range=range;ws.Shapes={AddChart:(t,l,tp,w,h)=>{const ch={type:t,SetSourceData(r){ch.source=r.Address();},ChartTitle:{}};ws.charts.push(ch);return {Chart:ch};}};
    return ws;
  }
  Object.entries(sheets).forEach(([name,rows])=>book.list.push(sheet(name,rows)));
  book.active=book.list[0];
  book.Worksheets={get Count(){return book.list.length;},Item:k=>typeof k==='number'?book.list[k-1]:book.list.find(s=>s.Name===k),Add:(before,after)=>{const ws=sheet('Sheet'+(book.list.length+1),[]);if(before&&!after)book.list.splice(book.list.indexOf(before),0,ws);else book.list.push(ws);return ws;}};
  book.Activate=()=>{};
  return book;
}
export function makeApp(book,select){
  const app={ActiveWorkbook:book,get ActiveSheet(){return book.active;},Calculation:-4105,selection:null};
  Object.defineProperty(app,'Selection',{get(){return app.selection;}});
  Object.defineProperty(app,'ActiveCell',{get(){return app.active||app.selection.Cells.Item(1,1);}});
  app.select=(a,activeAddr)=>{const ws=book.active;if(Array.isArray(a)){const areas=a.map(x=>ws.Range(x));const first=areas[0];app.selection=Object.assign(Object.create(null),first,{Areas:{Count:areas.length,Item:i=>areas[i-1]},Rows:first.Rows,Columns:first.Columns,Cells:first.Cells,Address:first.Address});}else{const r=ws.Range(a);r.Areas={Count:1,Item:()=>r};app.selection=r;}app.active=activeAddr?ws.Range(activeAddr).Cells.Item(1,1):null;return app;};
  if(select)app.select(select);
  return app;
}
