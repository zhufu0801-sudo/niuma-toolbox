# 办公工具箱功能模块

| 文件 | 内容 |
|---|---|
| `core.js` | 共用的读取与检查函数、功能登记表、执行入口、功能开关。必须第一个加载 |
| `word.js` | WPS 文字：标题、正文排版、目录、文字转表格、清理空段落、查找替换、插入表格、匹配周围格式 |
| `doc-plus.js` | WPS 文字“文档整理”：删除空白页、公文格式、日文商务文书格式、段落乱跳页、选中内容改横向、表格美化、图片不乱跑、加粗行内标题、选项对齐、数字英文不断行、去水印、中英文空格。Microsoft Word 版在 `public/office/office-docplus.js` |
| `sheet-plus.js` | WPS 表格“表格整理”：聚光灯、隔行填色、文本转数字、提取、两表核对、工资条、合并单元格编号、排名、到期提醒、进度条、多列合一格、拆行、堆成一列、打印表头、一页宽、删除空白行、目录、改名、按列拆分、多表合并、另存为文件、图表、互换、转置、姓名对齐、按颜色统计、前后缀、IFERROR |
| `sheet-check.js` | 表格体检：只检查不修改 |
| `sheet.js` | WPS 表格：美化表头、列宽、换行、筛选、清理空格、删除重复行、查找替换 |
| `slides.js` | WPS 演示：按大纲制作、插入表格、统一本页字体、复制当前页、匹配本页格式 |
| `batch.js` | 批量与拆分：批量转 PDF（三个宿主）、批量替换文字、批量加水印、按标题拆分文档（WPS 文字）、发送到另一个打开的文档（三个宿主） |
| `proof.js` | WPS 文字“检查与校对”（0.16）：发送前检查、错别字与标点检查（规则在 `public/proof-engine.js`）、比较两个文档、批注修订汇总导出、图表自动编号、修复多级编号 |
| `sheet-data.js` | WPS 表格“数据工具”（0.16）：万元显示、号码校验（规则在 `public/id-engine.js`）、合并多个工作簿、批量二维码 / 条形码（本机服务 `lib/codes.mjs` 生成 PNG）、考勤统计 |
| `slides-export.js` | WPS 演示“导出与瘦身”（0.16）：每页导出图片 / 拼长图（`lib/images.mjs`）、讲义打印、PPT 瘦身 |
| `files-more.js` | 三个宿主（0.16）：批量重命名（`lib/rename-files.mjs`，留记录可改回）、常用语一键插入 |
| `transfer.js` | 三个宿主通用：内容卡片（存、插入、插入最新）、中转区、导出 PDF、备份当前文件 |

加载顺序固定为 core → word → doc-outline（Word 转 PPT 的大纲整理，在 public 根目录）→ doc-plus → word-plus（方法卡的一键版：上下标、落款靠右、斜线表头、合并文档、自动边框） → sheet → sheet-plus → sheet-check → slides → transfer → links → locale → print → dates → approval → proof-engine → id-engine → proof → sheet-data → slides-export → files-more，见 `public/index.html`、`public/tools.html` 和 `test/helpers/tools-code.mjs`，三处要一致。

## 添加一个功能

```js
OfficeTools.register({
  id: 'myAction',            // 与 ribbon.xml 里按钮的 id 相同
  label: '功能名称',          // 侧栏和“选择要显示的功能”里显示
  desc: '一句话说明',
  icon: 'paragraph',         // public/icons/ 里的图标名
  group: 'format',           // format / table / slide / move / file / create
  panel: true,               // false：不放进侧栏“常用”列表（例如有专门的输入区）
  undo: true,                // 文字里是否合并成一步撤销
  hosts: { wps: function (c) { /* c.a 应用, c.d 文档, c.h 宿主, c.options 选项 */ return '完成提示'; } }
});
```

- 同一个 id 可以在不同模块里分别登记不同宿主，会自动合并。
- 返回的提示以“没有 / 当前选区没有 / 选区内没有”开头时，功能区点击会弹窗告诉用户，避免看起来像没反应。
- 新按钮要在 `ribbon.xml` 里加 `getVisible="GetVisible"`，才能跟随用户的功能开关。

## 更新版本时

三处必须同时改成同一个版本号：`package.json` 的 version、`public/version.json`、`core.js` 里的 `VERSION`。
WPS 不重启会继续使用旧脚本；版本不一致时，功能区按钮会提示用户重新打开 WPS。

`links.js`：A3 内容联动。书签/命名区域/幻灯片标签定位；自动读取已打开来源，手动确认目标更新。目标快照不符时拒绝覆盖；结构变化要求重新插入。

0.16 的新工具返回给用户的提示直接按界面语言写（工具里的 `M(中文, 日文)`），数字、文件名夹在句子中间时比词典逐词翻译通顺；固定的名称、说明、选项仍写中文，日文放在 `i18n-ja-extra.js` 最后一段。
Microsoft Office 版在 `public/office/office-batch2.js`：Office.js 能做到的（号码校验、万元显示、二维码、考勤统计、错别字检查、发送前检查、常用语）直接做，做不到的给操作步骤。
