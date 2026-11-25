(function(root){
 'use strict';var source='https://www.douyin.com/video/7665725907922996532';
 var items=[
 {id:'selectionStats',label:'复制选区合计',desc:'求和、平均、计数，点一下复制结果',group:'统计',steps:'选中数据 → 读取选区 → 点击需要的统计值',kind:'stats',source:source},
 {id:'filterValues',label:'按关键词多选筛选',desc:'搜到多个项目，一次保留这些行',group:'整理',steps:'连同表头选择数据 → 选列并勾选项目 → 生成筛选副本',kind:'filter',source:source},
 {id:'unmergeFill',label:'拆开合并并补齐',desc:'把合并格拆开，每格填回原内容',group:'整理',steps:'完整选中合并区域 → 查看影响范围 → 在副本中处理',kind:'edit',source:source},
 {id:'mergeSame',label:'连续同名合并',desc:'把一列中相邻的相同文字合成一格',group:'整理',steps:'选择一列分类文字，不含表头 → 预览 → 在副本中合并',kind:'edit',note:'只合并相邻且完全相同的非空内容。筛选前可用“拆开合并并补齐”。',source:source},
 {id:'dropdown',label:'制作下拉选项',desc:'每行填一个选项，自动去掉重复项',group:'录入',steps:'选择录入区域 → 填选项或读取选区内容 → 在副本中设置',kind:'dropdown',source:source},
 {id:'preventDuplicates',label:'不许重复录入',desc:'为一列设置重复输入提醒',group:'录入',steps:'选中一列录入区域，不含表头 → 预览 → 在副本中设置',kind:'edit',note:'限制普通键盘录入；粘贴可能绕过数据验证。只检查所选区域。',source:source},
 {id:'cleanInvisible',label:'清理看不见的字符',desc:'解决文字看着一样却匹配不上的问题',group:'整理',steps:'选中要清理的区域 → 对比前后示例 → 在副本中清理',kind:'edit',note:'移除零宽字符、BOM 和部分控制字符；保留换行和制表符，跳过数值与公式。',source:source},
 {id:'groupSummary',label:'按分类汇总',desc:'搜列名，按部门或品类统计次数和金额',group:'统计',steps:'连同表头选择数据 → 选分类列和金额列 → 生成汇总表',kind:'summary',note:'生成静态汇总，不是数据透视表；文本金额不参与求和。',source:source},
 {id:'formulaAssistant',label:'用中文写公式',desc:'查价格、算合计、判断条件，直接关联数据',group:'公式',kind:'formula',steps:'描述需求 → 确认来源和结果位置 → 预览写入',source:source},
 {id:'clearConstants',label:'清空数据，保留公式',desc:'重复使用模板时，公式和格式留下来',group:'整理',steps:'只选择要清空的数据区，不含表头 → 预览 → 生成模板副本',kind:'edit',source:'https://www.douyin.com/video/7666856044953472296'},
 {id:'uniqueCount',label:'去重并统计次数',desc:'列出每个姓名或项目出现了几次',group:'统计',steps:'连同表头选中数据 → 选统计列 → 生成新汇总表',kind:'summary',source:'https://www.douyin.com/video/7670184411354172708'},
 {id:'textIds',label:'长编号按文本录入',desc:'为后续输入订单号保留完整数字',group:'录入',steps:'选中待录入区域 → 预览 → 在副本中设为文本',kind:'edit',note:'只改变后续录入格式。已经丢失的长数字精度无法恢复；原有公式跳过。',source:'https://www.douyin.com/video/7683383527190170922'},
 {id:'readingGuide',label:'阅读时高亮行列',desc:'移动光标时看清当前行和列',group:'原生技巧',kind:'guide',steps:'打开 WPS「视图」→「阅读模式 / 高亮行列」→ 用旁边的小箭头选颜色。',note:'使用 WPS 原生显示效果，不修改单元格底色。此入口提供操作指引。',source:source},
 {id:'renameGuide',label:'文件不用关闭就改名',desc:'直接从顶部文件标签重命名',group:'原生技巧',kind:'guide',steps:'在顶部文件标签上点右键 →「重命名」（部分版本可点击文件名）→ 输入新名字。',note:'此入口提供操作指引；重命名可能影响已有跨文件链接，请先检查。',source:source},
 {id:'pivotGuide',label:'快速找到透视表字段',desc:'字段多时用搜索，少翻长列表',group:'原生技巧',kind:'guide',steps:'点透视表中的任意格 → 在右侧字段列表搜索列名 → 勾选字段，再拖到行、列或值。',note:'需要已有透视表且 WPS 版本提供字段搜索。简单汇总可用「按分类汇总」。',source:source}
 ];
 // AI：打开侧栏里的面板（AI 对话、AI 翻译润色与写作）。所有程序都有。
 var ai=[
 {id:'ocrImage',label:'图片转文字',desc:'粘贴截图或选图片，识别出文字（本机识别，不上传）',group:'文件与图片',kind:'panel',panel:'fileToolsPanel',action:'ocr',steps:'截图 → 在面板里按 Ctrl+V → 识别文字 → 复制或插入文档'},
 {id:'pdfToWord',label:'PDF 转 Word',desc:'PDF 转成可以编辑的 Word，扫描件先识别文字',group:'文件与图片',kind:'panel',panel:'fileToolsPanel',action:'pdf',steps:'选择 PDF → 自动转换 → 打开生成的 Word'},
 {id:'extractImages',label:'提取文档里的图片',desc:'把文档、表格、演示或 PDF 里的图片全部存到一个文件夹',group:'文件与图片',kind:'panel',panel:'fileToolsPanel',action:'images',steps:'保存文档 → 提取全部图片 → 打开文件夹'},
 {id:'compressImages',label:'压缩文档里的图片',desc:'图片太大、文件发不出去时，压缩后另存一份',group:'文件与图片',kind:'panel',panel:'fileToolsPanel',action:'images',steps:'保存文档 → 选压缩程度 → 压缩图片另存'},
 {id:'docHistory',label:'历史版本',desc:'查看并打开以前备份的版本，原文件不被替换',group:'文件与图片',kind:'panel',panel:'fileToolsPanel',action:'history',steps:'查看历史版本 → 打开某个版本 → 需要时另存为'},
 {id:'aiChat',label:'AI 对话',desc:'问操作方法、写公式、起草文字，回答可以插入文档',group:'AI',kind:'panel',panel:'aiChatPanel',steps:'输入问题（可附带选中的内容）→ 发送 → 复制或插入文档'},
 {id:'aiSummary',label:'AI 总结要点',desc:'长文压缩成一句话概括和几条要点',group:'AI',kind:'panel',panel:'aiTextPanel',action:'summary',steps:'选中文字 → 读取选区 → 让 AI 处理 → 插在原文后'},
 {id:'aiContinue',label:'AI 续写',desc:'接着选中的文字往下写，保持文风',group:'AI',kind:'panel',panel:'aiTextPanel',action:'continue',steps:'选中已写好的内容 → 读取选区 → 让 AI 处理 → 接在原文后'},
 {id:'aiExpand',label:'AI 扩写',desc:'几句话扩成一段完整的文字',group:'AI',kind:'panel',panel:'aiTextPanel',action:'expand',steps:'选中文字 → 读取选区 → 让 AI 处理 → 替换或插在后面'},
 {id:'aiShorten',label:'AI 精简改短',desc:'压缩到一半左右，保留关键信息',group:'AI',kind:'panel',panel:'aiTextPanel',action:'shorten',steps:'选中文字 → 读取选区 → 让 AI 处理 → 替换原文'},
 {id:'aiBullets',label:'AI 整理成条目',desc:'一大段话整理成一条一条的要点',group:'AI',kind:'panel',panel:'aiTextPanel',action:'bullets',steps:'选中文字 → 读取选区 → 让 AI 处理 → 替换或插在后面'},
 {id:'aiFormal',label:'AI 改成正式书面语',desc:'口语改成得体的书面语',group:'AI',kind:'panel',panel:'aiTextPanel',action:'formal',steps:'选中文字 → 读取选区 → 让 AI 处理 → 替换原文'},
 {id:'aiTranslate',label:'AI 翻译与润色',desc:'中⇄日翻译、商务敬语、润色',group:'AI',kind:'panel',panel:'aiTextPanel',action:'auto',steps:'选中文字 → 选处理方式 → 读取选区 → 让 AI 处理 → 写入'}
 ];
 function list(host){var out=(host==='et'?items.slice():[]).concat(root.MethodCards?root.MethodCards.forHost(host):[]);var avail={};OfficeTools.list(host).forEach(function(x){if(x.available!==false&&!OfficeTools.isHidden(x.id))avail[x.id]=x;});out=out.filter(function(c){if(c.run&&!avail[c.run])delete c.run;return true;});var covered={};out.forEach(function(c){if(c.run)covered[c.run]=1;});OfficeTools.list(host).filter(function(x){return x.panel&&!OfficeTools.isHidden(x.id)&&!covered[x.id];}).forEach(function(x){out.push({id:x.id,label:x.label,desc:x.desc,group:x.group==='locale'?'中日本地化':x.group==='proof'?'检查与校对':x.group==='sheetdata'?'数据工具':x.group==='slideout'?'导出与瘦身':x.group==='date'?'日期':x.group==='stamp'?'中日本地化':x.group==='docfix'?'文档整理':x.group==='sheetplus'?'表格整理':'常用操作',kind:'existing',fields:x.fields||null,steps:x.fields&&x.fields.some(function(f){return f.required!==false;})?'填写下面的内容，再点“执行这个工具”。':x.group==='sheetplus'?'先选中要处理的单元格（需要两块区域的按住 Ctrl 选），再点“执行这个工具”。会改原数据的会先确认；表格里的修改不能撤销，重要文件先备份。':x.group==='docfix'?'把光标放到要处理的位置或选中内容（不选则处理全文，会先询问），再点“一键完成”。可以按 Ctrl+Z 撤销。':'先选择要处理的内容，再点击执行。'});});return out.concat(ai.map(function(x){return Object.assign({},x);}));}
 root.QuickCatalog={items:items,ai:ai,list:list};
})(window);
