(function () {
  function safe(fn, fallback) { try { var value=fn(); return value === undefined || value === null ? fallback : value; } catch (ignore) { return fallback; } }
  function app() { return safe(function(){ return window.Application || (window.wps && window.wps.Application) || window.wps; },null); }
  function text(value, limit) { return String(value === undefined || value === null ? '' : value).replace(/\u0007/g,'').slice(0,limit); }
  function read(host) {
    var a=app(), c={ connected:!!a, host:host, document:'', selection:'', location:'', detail:'' };
    if(!a) return c;
    if(host==='wps') {
      c.document=text(safe(function(){return a.ActiveDocument.Name;},''),200);
      c.selection=text(safe(function(){return a.Selection.Text;},''),4000);
      c.location=c.selection.trim() ? '已选中文字' : '文字编辑区';
      c.detail=c.selection.trim() ? '已读取 '+c.selection.length+' 字；最多使用 4000 字。' : '选中一段文字，建议会结合你的操作位置。';
    } else if(host==='et') {
      c.document=text(safe(function(){return a.ActiveWorkbook.Name;},''),200);
      var s=safe(function(){return a.Selection;},null);
      var sheet=text(safe(function(){return a.ActiveSheet.Name;},''),100);
      var address=safe(function(){return typeof s.Address==='function' ? s.Address() : s.Address;},'');
      c.location=sheet+(address ? ' · '+text(address,100) : '');
      var rows=Number(safe(function(){return s.Rows.Count;},0));
      var cols=Number(safe(function(){return s.Columns.Count;},0));
      var lines=[];
      for(var r=1;r<=Math.min(rows,8);r++){
        var cells=[];
        for(var col=1;col<=Math.min(cols,5);col++){
          var cell=safe(function(){return s.Cells.Item(r,col);},null);
          cells.push(text(safe(function(){return cell.Text;},''),100));
        }
        lines.push(cells.join(' | '));
      }
      c.selection=lines.join('\n').slice(0,4000);
      c.detail=rows && cols ? '选区 '+rows+' 行 × '+cols+' 列；预览最多 8 行 × 5 列。' : '先选择一片连续的单元格区域。';
    } else if(host==='wpp') {
      c.document=text(safe(function(){return a.ActivePresentation.Name;},''),200);
      c.selection=text(safe(function(){return a.ActiveWindow.Selection.TextRange.Text;},''),4000);
      var slide=safe(function(){return a.ActiveWindow.View.Slide;},null);
      var index=safe(function(){return slide.SlideIndex;},'');
      c.location=index ? '第 '+index+' 页幻灯片' : '演示编辑区';
      if(!c.selection && slide){
        var parts=[],count=Number(safe(function(){return slide.Shapes.Count;},0));
        for(var i=1;i<=Math.min(count,8);i++){
          var shape=safe(function(){return slide.Shapes.Item(i);},null);
          var t=text(safe(function(){return shape.TextFrame.TextRange.Text;},''),500);
          if(t)parts.push(t);
        }
        c.selection=parts.join('\n').slice(0,4000);
        c.detail='当前页文字预览；最多读取 8 个对象。';
      }else c.detail=c.selection ? '已读取选中文字。' : '先点击一页幻灯片，或选中其中的文字。';
    }
    c.connected=!!c.document;
    return c;
  }
  window.CoachBridge={ read:read };
}());
