(function(root){
  'use strict';var F=root.OfficeFormula;
  // WPS removes optional quotes around simple sheet names. Preserve string literals.
  function canonical(formula){return String(formula).replace(/"(?:[^"]|"")*"|'(?:[^']|'')*'!|[^"']+/g,function(token){if(token[0]==='"')return token;if(token[0]==="'"){var name=token.slice(1,-2).replace(/''/g,"'");return /^[A-Za-z_\u0080-\uffff][A-Za-z0-9_\u0080-\uffff]*$/.test(name)?name+'!':token;}return token;});}
  function sameFormula(a,b){return canonical(a)===canonical(b);}
  function bookId(book){return String(book.FullName||book.Name);}
  function sheet(book,name){var s=book.Worksheets.Item(name);if(!s)throw Error('找不到工作表：'+name);return s;}
  function range(book,r){return sheet(book,r.sheet).Range(F.address(r,false));}
  function plainCell(c){if(c.HasFormula!==false&&c.HasFormula!==0)throw Error('结果区域存在公式或无法确认状态。');if(c.Value2!==null&&c.Value2!==undefined&&c.Value2!=='')throw Error('结果区域已有内容，请选择空白位置。');if(c.MergeCells!==false&&c.MergeCells!==0)throw Error('结果区域含合并单元格。');try{if(c.HasArray)throw Error('结果区域包含数组公式。');}catch(e){if(e.message==='结果区域包含数组公式。')throw e;}}
  function validTarget(book,p){var s=sheet(book,p.target.sheet);if(s.ProtectContents)throw Error('结果工作表已保护，请先在 WPS 中解除保护。');var r=range(book,p.target);for(var i=1;i<=p.formulas.length;i++)plainCell(r.Cells.Item(i,1));return r;}
  function capture(app){var b=app.ActiveWorkbook,r=app.Selection;if(!b||!r||!r.Rows||!r.Columns)throw Error('请在 WPS 表格中选择区域。');if(r.Areas&&r.Areas.Count!==1)throw Error('请选择连续区域。');var rows=Number(r.Rows.Count),cols=Number(r.Columns.Count);if(rows*cols>100000)throw Error('区域超过 10 万个单元格。');var result={sheet:String(app.ActiveSheet.Name),r1:Number(r.Row),r2:Number(r.Row)+rows-1,c1:Number(r.Column),c2:Number(r.Column)+cols-1};return {book:bookId(b),reference:F.full(result)};}
  function preview(app,input){var b=app.ActiveWorkbook;if(!b)throw Error('请先打开工作簿。');if(b.ReadOnly)throw Error('工作簿只读。');if(input.book&&input.book!==bookId(b))throw Error('已切换工作簿，请重新选择区域。');var p=F.plan(input);p.dependencies.forEach(function(r){range(b,r);});validTarget(b,p);return {book:bookId(b),plan:p,at:Date.now()};}
  function assertBook(app,prepared){var b=app.ActiveWorkbook;if(!b||bookId(b)!==prepared.book)throw Error('当前工作簿已切换或另存，请重新预览。');if(b.ReadOnly)throw Error('工作簿只读。');return b;}
  function apply(app,prepared){
    var b=assertBook(app,prepared),p=prepared.plan,r=validTarget(b,p),written=[];
    p.dependencies.forEach(function(dep){range(b,dep);});
    try{for(var i=0;i<p.formulas.length;i++){var c=r.Cells.Item(i+1,1);plainCell(c);written.push({cell:c,formula:p.formulas[i]});c.Formula=p.formulas[i];if(!sameFormula(c.Formula,p.formulas[i]))throw Error('宿主未接受第 '+(i+1)+' 行公式。');}}
    catch(e){var failed=0;written.reverse().forEach(function(w){try{if(sameFormula(w.cell.Formula,w.formula))w.cell.ClearContents();else if(w.cell.Value2!==''&&w.cell.Value2!==null&&w.cell.Value2!==undefined)failed++;}catch(ignore){failed++;}});throw Error('写入未完成：'+e.message+(failed?'；部分单元格未能恢复，请检查结果区域。':'；本次已写入公式已移除。'));}
    try{r.Calculate();}catch(ignore){}
    var errors=[];for(var j=1;j<=p.formulas.length;j++){try{var text=String(r.Cells.Item(j,1).Text||'');if(/^#(?:N\/A|NAME\?|VALUE!|REF!|DIV\/0!|NUM!|SPILL!|CALC!)/.test(text))errors.push(j);}catch(ignore){}}
    return {book:prepared.book,plan:p,at:Date.now(),errors:errors};
  }
  function undo(app,record){var b=assertBook(app,record),s=sheet(b,record.plan.target.sheet);if(s.ProtectContents)throw Error('工作表已保护，无法撤回。');var r=range(b,record.plan.target),count=record.plan.formulas.length;
    for(var i=1;i<=count;i++)if(!sameFormula(r.Cells.Item(i,1).Formula,record.plan.formulas[i-1]))throw Error('结果公式已被修改，已停止撤回，保留你的修改。');
    var cleared=[];try{for(var j=1;j<=count;j++){var c=r.Cells.Item(j,1);c.ClearContents();cleared.push(j);}}catch(e){var failed=0;cleared.forEach(function(n){try{r.Cells.Item(n,1).Formula=record.plan.formulas[n-1];}catch(ignore){failed++;}});throw Error('撤回失败：'+e.message+(failed?'；部分公式未恢复，请检查。':'；已恢复已清除的公式。'));}return count;
  }
  root.OfficeFormulaHost={capture:capture,preview:preview,apply:apply,undo:undo};
})(window);
