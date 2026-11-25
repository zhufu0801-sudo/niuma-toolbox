const C=require('./manual-extra.cjs'),{T,add,feature,defaultUndo,inv,tr,coverage}=C;
const {splitSteps}=require('../manual-checks.cjs');
const steps=(z,j,e)=>T(splitSteps(z),splitSteps(j),splitSteps(e));
function f(ch,id,title,purpose,st,opt,lim,u=defaultUndo.copy){add(ch,feature(id,title,purpose,st,opt,lim,u));coverage.push({id,chapter:ch});}
const outLimit=T('只处理支持的已保存格式，生成后核对结果。字体、图片和分页可能受宿主或系统组件影响。','対応する保存済み形式を処理し、出力を確認します。フォント・画像・改ページはホストやシステムの影響を受けます。','Processes supported saved formats. Inspect output; fonts, pictures, and pagination depend on host and system components.');
const files=[
['ocrImage','图片转文字','画像から文字起こし','Image OCR','在本机识别图片中的文字。','画像内の文字をローカルで読み取ります。','Recognize image text locally.','在文件与图片面板粘贴截图或选图片 → 点识别，核对金额、编号和标点 → 复制结果或插入文档','画像を選ぶかスクリーンショットを貼り付けます → 認識して金額・番号・句読点を確認します → コピーまたは文書へ挿入します','Choose an image or paste a screenshot → Recognize and verify amounts, IDs, and punctuation → Copy or insert the text'],
['pdfToWord','PDF 转 Word','PDF を Word に変換','PDF to Word','把 PDF 内容转换为可编辑文档。','PDF の内容を編集できる文書にします。','Convert PDF content into an editable document.','选 PDF → 等待转换，扫描件需要识别 → 打开 Word 结果核对段落和表格','PDF を選びます → 変換を待ち、スキャンは文字認識します → 結果の段落と表を確認します','Select a PDF → Convert, using OCR for scanned content → Open the result and inspect paragraphs and tables'],
['extractImages','提取文档里的图片','文書内の画像を取り出す','Extract document images','把支持文档内的图片另存到文件夹。','対応する文書内の画像をフォルダーへ保存します。','Save embedded images from supported files into a folder.','保存源文件 → 在文件与图片面板选择提取图片 → 打开结果文件夹核对图片','元ファイルを保存します → ファイルと画像で画像抽出を選びます → 出力フォルダーで画像を確認します','Save the source → Choose image extraction in Files and images → Inspect the output folder'],
['compressImages','压缩文档里的图片','文書内の画像を圧縮','Compress document images','压缩图片后另存文档副本。','画像を圧縮して文書のコピーを作ります。','Compress images and save a document copy.','保存文档 → 选压缩程度并生成副本 → 打开副本，放大检查小字与图片','文書を保存します → 圧縮の程度を選びコピーを作ります → 拡大して小さい文字と画像を確認します','Save the document → Choose compression and create a copy → Zoom in and inspect small text and images'],
['docHistory','历史版本','履歴バージョン','Backup history','打开已保存的历史备份。','保存された過去のバックアップを開きます。','Open a previously saved backup.','打开历史版本 → 按时间选择候选版本 → 打开核对，需恢复时另存为新文件','履歴を開きます → 時刻で候補を選びます → 内容を確認し、必要なら別名保存します','Open history → Select a candidate by time → Inspect it and save under a new name if needed']
];
for(const x of files)f('files',x[0],T(x[1],x[2],x[3]),T(x[4],x[5],x[6]),steps(x[7],x[8],x[9]),T('使用面板提供的文件选择、识别语言或质量选项；不支持的选项不会凭空补出。','パネルのファイル・認識言語・品質設定を使います。対応していない設定を追加するものではありません。','Use the panel’s file, OCR-language, or quality controls where offered.'),x[0]==='docHistory'?T('只有实际存在的备份可打开，不能找回从未保存且未备份的内容。','実際に存在するバックアップだけを開けます。保存もバックアップもない内容は復元できません。','Only existing backups can be opened. Unsaved content with no backup cannot be recovered.'):outLimit);
f('files','mail-merge-panel',T('按名单批量生成','名簿から一括作成','Generate documents from a list'),T('用一份 DOCX 模板和一张 XLSX 名单为每行生成文档。','DOCX テンプレートと XLSX 名簿から行ごとに文書を作ります。','Generate one document per list row from a DOCX template and XLSX list.'),steps('在模板中输入 {{列名}}，例如 {{姓名}}，并保存为 DOCX → 选择 XLSX 名单，第一行写列名 → 点检查，核对已用、未用和未知占位符 → 选择分别生成或合成一个文件，点生成','テンプレートに {{列名}} を書き DOCX で保存します → 1 行目が列名の XLSX 名簿を選びます → チェックで使用列と不明な項目を確認します → 個別または一つの文書にまとめて作成します','Save a DOCX template with {{column name}} placeholders → Select an XLSX list whose first row contains column names → Check used, unused, and unknown placeholders → Choose separate files or one combined document and generate'),T('分别生成时可选一列作文件名；不选则顺序编号。','個別作成は列をファイル名に使えます。選ばない場合は連番です。','For separate files, choose a naming column or keep sequential numbering.'),T('名单最多 500 行；列名必须与占位符一致。未知占位符不会自动猜测替换。只读已保存内容。','名簿は最大 500 行です。列名と記号を一致させます。不明な記号を推測して置き換えず、保存済みの内容だけを使います。','Up to 500 rows. Placeholder names must match columns; unknown placeholders are not guessed. Only saved content is read.'));
const prints=[
['fax-cover','传真封面','FAX 送付状','Fax cover sheet','填写收发件人、传真号、主题、页数和正文。','宛先・差出人・FAX 番号・件名・枚数・本文を入力します。','Enter recipient, sender, fax number, subject, page count, and message.'],
['address-labels','地址标签','宛名ラベル','Address labels','每行一条地址，按邮编、地址、姓名填写；可从表格读取。','1 行に郵便番号・住所・氏名を入力します。表からも読み取れます。','Enter one postal code, address, and name per line, or read them from a sheet.'],
['envelope','信封与明信片','封筒とはがき','Envelopes and postcards','选择纸张，填寄件人和收件地址，按需要勾选邮编格。','用紙を選び、差出人と宛先を入力し、必要なら郵便番号枠を選びます。','Choose paper, sender and recipient details, and optional postal-code boxes.'],
['vertical','日文竖排通知与挨拶状','縦書き通知と挨拶状','Japanese vertical notices','选择例文或输入标题、正文，一行一段。','例文を選ぶか、タイトルと本文を 1 行 1 段落で入力します。','Choose a sample or enter title and body, one paragraph per line.'],
['shipping','寄件信息单','発送情報シート','Shipping information sheet','选 A4 或标签纸，填寄件人、物品备注、收件姓名电话地址。','A4 またはラベル用紙を選び、差出人・品物・宛先・電話を入力します。','Choose A4 or label paper and enter sender, item notes, recipient, phone, and address.']
];
for(const x of prints)f('print',x[0],T(x[1],x[2],x[3]),T('生成可核对后打印的文书副本。','確認して印刷できる書類のコピーを作ります。','Create a document copy for review before printing.'),steps('打开“传真、信封、寄件单与模板” → '+x[4]+' → 生成并打开副本，确认纸张与内容后自行打印','書類テンプレートの画面を開きます → '+x[5]+' → コピーを開き、用紙と内容を確認して自分で印刷します','Open the fax/envelope/shipping/template panel → '+x[6]+' → Generate and inspect the copy before printing yourself'),T('文件类型、可见的纸张及文字字段按所选种类变化。','種類に応じて用紙と入力欄が切り替わります。','Paper choices and text fields change with the document type.'),T('不发送传真、不购买运单、不自动出纸；寄件信息单不是物流公司正式面单。地址和页数由你核对。','FAX の送信、送り状の購入、自動印刷はしません。発送情報シートは運送会社の正式な送り状ではありません。住所と枚数を確認してください。','No fax transmission, postage purchase, or automatic printing. The shipping information sheet is not an official carrier label. Verify addresses and page counts.'));
f('print','print-copy-panel',T('省墨与传真优化副本','省インクと FAX 用コピー','Ink-saving and fax copies'),T('从已保存文件生成黑白或传真优化副本。','保存済みファイルから白黒または FAX 向けコピーを作ります。','Create monochrome or fax-optimized copies of saved files.'),steps('保存 DOCX、XLSX、PPTX 或支持的 PDF → 填完整路径，选择黑白、副本传真优化或传真黑白 PDF → 点生成，打开副本检查分页、图片和细线','対応するファイルを保存します → 完全パスと白黒・FAX 最適化・FAX 白黒 PDF を選びます → 作成して改ページ・画像・細い線を確認します','Save a supported source file → Enter its full path and choose monochrome, fax-optimized copy, or fax PDF → Generate and inspect page breaks, images, and thin lines'),T('黑白副本、传真优化副本、传真黑白 PDF；骑缝章另见下项。','白黒コピー、FAX 最適化コピー、FAX 白黒 PDF を選びます。割印は次の項目です。','Monochrome copy, fax-optimized copy, or fax monochrome PDF; see the next entry for seam stamps.'),T('传真优化会调整颜色、字号和线条。PDF 结果为图片型，不保留可搜索文字、链接或数字签名。','FAX 最適化では色・文字サイズ・線を変更します。PDF は画像となり、検索文字・リンク・電子署名は残りません。','Fax optimization can change colours, font sizes, and lines. Raster PDF output does not preserve searchable text, links, or digital signatures.'));
f('print','stamp',T('电子印章图片与 PDF 骑缝章','印影画像と PDF の割印','Stamp images and PDF seam stamps'),T('生成印章图片，或把已有 PNG 印章分段排在 PDF 页边。','印影画像を作るか、既存 PNG を分割して PDF のページ端へ配置します。','Create a stamp image or split an existing PNG across PDF page edges.'),steps('生成图片时选形状并填文字；骑缝章时准备已获准使用的透明 PNG → 在省墨与传真优化选择 PDF 骑缝章，填 PDF 与 PNG 完整路径 → 生成副本，按页序检查印章拼合方向和高度','印影は形と文字を指定し、割印は使用できる透明 PNG を用意します → PDF の割印を選び、PDF と PNG の完全パスを入力します → コピーを作り、ページ順・方向・高さを確認します','Choose shape and text for an image, or prepare an authorized transparent PNG → Select PDF seam stamp and enter the PDF and PNG paths → Generate a copy and check split order, direction, and height'),T('印章形状和文字；骑缝章按页数等分，在右边缘依次放置。','印影の形と文字を選びます。割印はページ数で等分し、右端に順に配置します。','Stamp shape and text. Seam stamps divide by page count and appear in sequence on right edges.'),T('图片不提供身份认证或数字签名能力；PDF 图像加工不会保留原来的签名和可搜索文字。','画像に本人認証や電子署名の機能はありません。画像処理した PDF は元の署名や検索文字を保持しません。','An image provides no identity verification or digital-signature capability. Raster processing does not preserve original PDF signatures or searchable text.'));
f('print','impose',T('PDF 多页合一与小册子','PDF の集約と小冊子','PDF multiple-up and booklet layout'),T('把 PDF 页面重排成便于打印的版面。','PDF のページを印刷向けに並べ直します。','Rearrange PDF pages for printing.'),steps('保存 PDF，在模板面板找到 PDF 排版 → 选多页合一或小册子及每页数量 → 生成结果并先打印少量测试页，核对双面翻转方向','PDF を保存し、テンプレート画面の PDF 配置を開きます → 集約または小冊子と枚数を選びます → 少数ページで試し刷りし、両面の反転方向を確認します','Save the PDF and open its layout controls → Choose multiple-up or booklet and pages per sheet → Generate and test-print a few pages to check duplex orientation'),T('按面板选择排版模式和每页数量；打印机双面由打印对话框决定。','配置方式と枚数を選びます。両面設定はプリンターの画面で行います。','Choose layout and page count; set duplex in the printer dialog.'),T('输出为图片型 PDF；不保留搜索文字、链接或签名，不自动控制纸盒和打印机。','画像 PDF として出力し、検索文字・リンク・署名は残しません。給紙やプリンターを自動操作しません。','Output is raster PDF without searchable text, links, or signatures. Paper trays and printers are not automatically controlled.'));
f('dates','library-panel',T('商务与日常文书模板','業務書類と日常書類のテンプレート','Business and everyday document templates'),T('从模板库创建可编辑的商务文件、日常表单和 PPT 副本。','テンプレートから業務書類・日常フォーム・PPT の編集可能なコピーを作ります。','Create editable business documents, everyday forms, and PPT copies from the library.'),steps('打开模板库，选商务、日常文书或 PPT 分类 → 选择中文或日文和具体模板，查看预览 → 点使用这套打开副本，填写实际内容 → 核对金额、税率、登记号、日期和收件人后保存','分類を業務・日常書類・PPT から選びます → 中国語または日本語とテンプレートを選び、確認します → コピーを開いて実際の内容を入力します → 金額・税率・登録番号・日付・宛先を確認して保存します','Choose business, everyday forms, or PPT → Select Chinese or Japanese and a template, then inspect preview → Open a copy and enter actual content → Verify amounts, rates, registration numbers, dates, and recipients before saving'),T('PPT 场景模板、商务表格、请假单和会议纪要等，具体以当前模板列表为准。消费税模板区分 10% 和 8%，提供尾数处理选项。','PPT、業務表、休暇届、議事録などを現在の一覧から選びます。消費税テンプレートは 10% と 8% を分け、端数処理を選べます。','Current library entries include PPT scenarios, business sheets, leave requests, and meeting minutes. Consumption-tax templates separate 10% and 8% and offer rounding choices.'),T('模板只是填写辅助，不替你核实交易与制度要求。不要把模板中的示例公司、示例金额当作真实数据。','入力を助けるテンプレートであり、取引内容や制度要件を確認するものではありません。例の会社名や金額を実データとして使わないでください。','Templates assist entry but do not validate transactions or requirements. Replace sample companies and amounts with real data.'));
add('cards',T(`### 收集 插入与联动

粘贴栏用于临时收集多段内容，卡片用于命名后反复使用；两者都在本机。选中来源再收集，切换到目标文档后放好光标或活动单元格，再选内容插入。不要在表格里把一个活动格误当作整块目标范围。普通插入产生独立内容，不会因为来源后来变化而自动更新。

只有“联动模块”记录来源与目标位置。更新前打开来源、查看版本和目标快照；目标被手改或结构变化时，停止覆盖是保护行为。需要重新插入时先保留手改版本。删除卡片或粘贴栏记录不等于删除已插入文档的内容。`,
`### 収集 挿入と連動

貼り付けトレイは一時的な複数内容、カードは名前を付けた再利用に向き、どちらもローカルです。元の内容を選んで集め、挿入先でカーソルやアクティブセルを選んでから入れます。アクティブセル一つを出力範囲全体と取り違えないでください。通常の挿入は独立した内容で、元の内容が変わっても自動更新しません。

連動モジュールだけが参照元と挿入位置を記録します。更新前に参照元を開き、版と挿入先の状態を確認します。手動変更や構造変更で上書きを止めるのは保護動作です。再挿入の前に手動変更版を残してください。カードやトレイの削除は、既に文書へ入れた内容の削除ではありません。`,
`### Collect insert and link

The paste bar collects temporary pieces; named cards support repeated reuse. Both are local. Select source content before collecting, then choose the destination cursor or active cell before insertion. Do not mistake one active cell for a whole output grid. Ordinary insertion creates independent content and does not update when its source changes.

Only linked modules record source and destination positions. Open sources and review versions and destination snapshots before updating. Blocking an overwrite after manual or structural changes protects your work. Retain manual changes before reinserting. Deleting a card or paste-bar record does not remove content already inserted into a document.`));
const aiEN={aiChat:['AI chat','Ask a question, optionally include a selected excerpt, send, then inspect and copy or insert the answer.'],aiSummary:['Summarize','Read selected text, request a summary, and inspect the conclusion and key points before inserting.'],aiContinue:['Continue writing','Read the existing passage, request a continuation, and review it before appending.'],aiExpand:['Expand text','Read a short passage, request expansion, and review added claims before replacing or appending.'],aiShorten:['Shorten text','Read a passage, request a shorter version, and check that necessary details remain.'],aiBullets:['Create bullet points','Read the passage, request structured points, and verify their order and meaning.'],aiFormal:['Formal wording','Read informal text, request formal wording, and inspect the tone before replacing.'],aiTranslate:['Translate and polish','Choose language or polishing mode, read the selection, request processing, and verify names and terminology.']};
for(const q of inv.ai.filter(q=>q.id.startsWith('ai'))){const e=aiEN[q.id];f('ai',q.id,T(q.label,tr(q.label,'ja'),e[0]),T(q.desc,tr(q.desc,'ja'),e[0]+'.'),steps('先在设置里配置并测试自己的 AI 服务 → '+q.steps+' → 写入前核对事实、数字和保护后的代号','自分の AI 接続を設定してテストします → '+tr(q.steps,'ja')+' → 書き込み前に事実・数字・記号を確認します','Configure and test your own AI service → '+e[1]+' → Verify facts, numbers, and masking codes before insertion'),T('服务商、地址、模型、Key 和保护方式在 AI 设置中配置；文本处理按面板提供的模式选择。','提供元・接続先・モデル・Key・保護方式は AI 設定で指定し、文章の処理方式はパネルで選びます。','Configure provider, endpoint, model, Key, and protection mode in AI settings; choose text-processing modes in the panel.'),T('会把发送内容交给所选服务商，可能计费。自动脱敏不是绝对匿名；不把 AI 回答当事实校验。未配置时可用本地工具建议。','送信内容は選んだ提供元へ渡り、費用が発生する場合があります。マスキングは完全な匿名化を保証しません。回答を事実確認の代わりにしないでください。未設定ならローカル機能の案内を使えます。','Sent content goes to your chosen provider and may incur charges. Masking is not a guarantee of anonymity, and AI output is not fact verification. Local tool suggestions remain useful without AI.'),T('请求发送后不能靠文档撤销让服务商“没收到”；文档插入可按宿主恢复规则处理。','文書の取り消しで送信済み内容を提供元から取り消すことはできません。文書への挿入はホストの規則で戻します。','Document undo cannot retract a request already received by a provider. Recover inserted content according to host rules.'));}
f('ai','ai-formula',T('AI 公式与大纲辅助','AI による数式とアウトラインの支援','AI formula and outline assistance'),T('让已配置的 AI 起草公式解释或 PPT 大纲，供你审核后使用。','設定した AI に数式や PPT の構成を下書きさせ、確認して使います。','Ask configured AI for formula explanations or a presentation outline for your review.'),steps('先确认 AI 设置与发送内容 → 在 AI 对话问公式，或在版式工作室填写主题、听众和章节数 → 生成后核对引用、函数兼容性和事实 → 手动预览后再写入公式或生成 PPT','AI 設定と送信内容を確認します → 対話で数式を質問するか、スタジオでテーマ・聴衆・章数を指定します → 参照・関数対応・事実を確認します → プレビューしてから書き込みや PPT 生成を行います','Review AI settings and outgoing content → Ask about formulas in chat, or enter topic, audience, and section count in the studio → Check references, function support, and facts → Preview before writing formulas or generating a deck'),T('大纲章节数可选 4、6、8、10；普通版式生成不需要 AI。','アウトラインは 4・6・8・10 章を選べます。通常のレイアウト生成に AI は不要です。','Outline section counts are 4, 6, 8, or 10. Ordinary layout generation does not require AI.'),T('与本地公式助手不同，AI 输出不保证可执行或正确；不要直接运行生成的宏或复杂公式。','ローカル数式アシスタントと異なり、出力の正しさや実行可能性は保証しません。生成したマクロや複雑な式をそのまま実行しないでください。','Unlike the local formula assistant, AI output is not guaranteed correct or executable. Review generated macros and complex formulas before use.'),defaultUndo.copy);
add('office',T(`### 入口和运行方式

Office 版使用 Office.js 侧栏，连接本机 HTTPS 服务；Office 证书和加载项登记必须完成。侧栏关闭后 MCP 不能继续通过该页访问文档。WPS 的原生对象接口和 Office.js 不相同，所以不能按 WPS 的功能数量推断 Office 全部可用。

| 功能类 | Office 版处理原则 |
|---|---|
| 文字 表格及演示读写 | 使用宿主公开的 Office.js 能力；不支持时给说明 |
| 文件 PDF 模板 | 可由本机服务处理支持的已保存文件 |
| 号码校验 单位显示 二维码 考勤等 | 有对应 Office 实现的入口直接执行 |
| 宏读取 原生打印或某些版式操作 | 不能使用 WPS 专用接口；可能只提供操作指引 |
| 脱敏复制与粘贴 | 用 Office API 读选区和写文字；权限及宿主限制仍适用 |

首次使用时用小样本检查当前 Word、Excel 或 PowerPoint 的结果；不能把网页代码测试当作桌面兼容性验收。本轮未安装 Microsoft Office，也没有 Office 真机测试。`,
`### 入口と実行方法

Office 版は Office.js サイドバーとローカル HTTPS を使うため、証明書とアドイン登録が必要です。サイドバーを閉じると、そのページ経由の MCP アクセスも止まります。WPS のネイティブ API とは異なり、WPS の機能数から Office の全機能対応を推定できません。

| 種類 | Office での扱い |
|---|---|
| 文書 表 プレゼンの読み書き | 公開された Office.js 機能を使い、非対応時は案内します |
| ファイル PDF テンプレート | 対応する保存済みファイルをローカルサービスで処理します |
| 番号検証 単位表示 QR 勤怠など | Office 用実装がある入口で実行します |
| マクロ読み取り 印刷 一部の配置 | WPS 専用 API は使えず、手順だけを示す場合があります |
| マスキングコピーと貼り付け | Office API で選択範囲を読み、文字を書きます。権限とホスト制限は残ります |

最初は小さなサンプルで現在の Word・Excel・PowerPoint の動作を確認してください。Web コードのテストはデスクトップ互換性の実機検証ではありません。今回 Microsoft Office はなく、実機検証していません。`,
`### Entry points and runtime

Office uses an Office.js sidebar and local HTTPS, requiring certificate trust and add-in registration. Closing the sidebar stops MCP access through that page. WPS native APIs differ from Office.js; a WPS feature count is not an Office compatibility guarantee.

| Feature class | Office behaviour |
|---|---|
| Document spreadsheet and presentation edits | Uses public Office.js capabilities and provides guidance when unsupported |
| Files PDF and templates | The local service can process supported saved files |
| ID validation units QR and attendance | Entries with Office implementations execute directly |
| Macro reading native printing some layouts | WPS-only APIs are unavailable; guidance may replace automation |
| Masked copy and paste | Office APIs read selections and write text, subject to permissions and host limits |

Test a small sample in your installed Word, Excel, or PowerPoint first. Web-code tests do not establish desktop compatibility. Microsoft Office was unavailable in this round and was not tested on a desktop.`));
add('faq',T(`### 现象 原因和处理

| 现象 | 常见原因 | 处理 |
|---|---|---|
| 找不到工具箱标签 | 未登记或宿主未重开 | 重新运行安装入口，保存后退出全部窗口再打开 |
| 本地服务连接失败 | 后台未启动或端口被占用 | 运行桌面 Office Toolbox；看 service-error.log |
| 更新后提示版本不一致 | 旧功能区仍缓存脚本 | 保存并完全退出 WPS，再打开 |
| 选区为空或范围过大 | 光标位置不对或选中整列 | 重新选择有限范围，不同时切换文档 |
| 右边有内容而拒绝输出 | 保护已有单元格 | 插入空白输出列后重新运行 |
| 号码少了末尾数字 | 之前按数值存储已丢精度 | 从原始来源重新输入为文本，不能靠格式找回 |
| 剪贴板一直无法读取 | 宿主权限限制 | 等待提示后粘入面板文本框，点还原并插入 |
| 代号没有复原 | 记录过期 被删 或不属于当前会话 | 查看未知代号列表，重新读取或使用正确记录 |
| 宏读取失败 | VBA 未安装或未授予项目访问 | 阅读权限提示，使用手动粘贴检查，不自动修改安全设置 |
| OCR 日文不识别 | 系统语言组件缺少 | 检查系统 OCR 语言；先用清晰横排印刷体测试 |
| 邮编更新失败 | 网络或官方文件不可达 | 保留现有数据，可下载官方 ZIP 后导入 |
| PDF 打不开或受保护 | 文件不完整或有密码 | 用原阅读器核对文件，由有权限的人处理保护 |
| 收藏或快捷键没触发 | 焦点不在侧栏或组合被占用 | 先点侧栏非输入区，或用功能区访问键 |

### 重启和诊断

先保存文档，关闭宿主窗口，然后重新运行 Office Toolbox 启动入口。不要随意结束电脑上全部 Node 进程，它们可能属于其他程序。仍失败时查看安装目录的 service.log 和 service-error.log；诊断摘要只附与问题相关的内容，隐藏 Key、客户名和文档路径中的私人信息。

重新安装前保留 generated 和独立文档备份。使用同一安装位置完成登记，重开宿主。只删除程序目录并不能自动撤销所有加载项登记；正常卸载更便于移除自己的登记。

### 通过 GitHub Issues 反馈

只使用项目仓库的 Issues。仓库地址以发布页或 README 为准；本手册不猜测尚未配置的仓库地址。选择 Bug 模板，填写程序与版本、系统、最小复现步骤、预期/实际结果、提示原文以及日志相关几行。使用合成样本；不要把 Key 或真实客户文档公开。可直接复制仓库 .github/ISSUE_TEMPLATE/bug_report.md 的三语模板。`,
`### 症状 原因と対処

| 症状 | よくある原因 | 対処 |
|---|---|---|
| タブがない | 登録前または再起動前 | インストールをやり直し、保存して全ウィンドウを開き直します |
| ローカル接続失敗 | サービス停止またはポート競合 | Office Toolbox を起動し service-error.log を確認します |
| 更新後の版不一致 | 古いスクリプトのキャッシュ | 保存して WPS を完全終了し開き直します |
| 範囲が空または大きすぎる | カーソル位置や列全体の選択 | 有限の範囲を選び直し、処理中は文書を変えません |
| 右側への出力を拒否 | 既存セルの保護 | 空き列を挿入してやり直します |
| 番号の末尾が消えた | 数値化で精度を失った | 原資料から文字列として入力し直します |
| クリップボードを読めない | ホストの権限制限 | 案内後にパネルへ貼り、復元して挿入します |
| 記号が戻らない | 期限切れ 削除 別の会話 | 不明な記号を確認し、読み直すか正しい記録を使います |
| マクロを読めない | VBA または権限がない | 説明を読み、手動貼り付けで調べます |
| OCR の日本語が読めない | システム言語がない | OCR 言語を確認し、鮮明な横書き印刷文字で試します |
| 郵便番号更新失敗 | 通信や公式ファイルに問題 | 既存データを残し、公式 ZIP を読み込めます |
| PDF が開けない | 不完全なファイルまたは保護 | 元の閲覧ソフトで確認し、権限のある人が対処します |
| お気に入りやキーが動かない | フォーカスや競合 | サイドバーの入力欄以外を選ぶかアクセスキーを使います |

### 再起動と診断

文書を保存してホストを閉じ、Office Toolbox の起動入口を実行します。他のアプリの可能性があるため、すべての Node プロセスを一括終了しないでください。解決しない場合はインストール先の service.log と service-error.log を確認します。診断情報は関係する部分だけを付け、Key・顧客名・パス内の個人情報を隠します。

再インストール前に generated と文書をバックアップし、同じ場所で登録してホストを開き直します。プログラムのフォルダー削除だけでは登録をすべて解除できません。通常のアンインストールで専用登録を外してください。

### GitHub Issues で報告

プロジェクトの Issues を使います。URL はリリースまたは README で確認し、未設定の URL をここでは推測しません。Bug テンプレートでアプリと版、OS、最小手順、期待と実際の結果、表示文、関係するログ数行を記入します。合成サンプルを使い、Key や実際の顧客文書を公開しないでください。リポジトリの .github/ISSUE_TEMPLATE/bug_report.md に三言語テンプレートがあります。`,
`### Symptoms causes and fixes

| Symptom | Common cause | Action |
|---|---|---|
| Toolbox tab missing | Registration or restart incomplete | Rerun installation, save, and reopen all host windows |
| Local service unavailable | Service stopped or port occupied | Launch Office Toolbox and inspect service-error.log |
| Version mismatch after update | Cached scripts | Save, fully exit WPS, and reopen |
| Empty or oversized selection | Wrong focus or an entire column selected | Select a bounded range and keep the same document active |
| Right-hand output refused | Existing cells protected | Insert blank output columns and retry |
| ID loses final digits | Previous numeric precision loss | Re-enter from the original source as text |
| Clipboard cannot be read | Host permission limits | Paste into the panel after the prompt, then restore and insert |
| Codes do not restore | Expired deleted or different-session map | Inspect unknown codes and reread or use the correct record |
| Macro reading fails | Missing VBA or project access | Read the guidance and use manual code paste |
| Japanese OCR unavailable | Missing system language component | Check OCR languages and test clear horizontal print |
| Postal update fails | Network or official file unavailable | Retain existing data and import an official ZIP if needed |
| PDF will not open | Incomplete file or protection | Check in its original reader; let an authorized person handle protection |
| Shortcut does not fire | Focus or key conflict | Focus the sidebar outside inputs, or use ribbon access keys |

### Restart and diagnostics

Save documents, close the host, and rerun the Office Toolbox launcher. Do not terminate every Node process: other applications may own them. If needed, inspect service.log and service-error.log in the installation directory. Include only relevant diagnostic lines and remove Keys, client names, and private path details.

Back up generated and independent documents before reinstalling. Register at the same location and reopen the host. Deleting the program folder alone does not remove all add-in registration; use the uninstaller for its own entries.

### Report through GitHub Issues

Use the project repository’s Issues page. Obtain its URL from the release or README; this manual does not guess an unconfigured repository address. Use the bug template with application/version, OS, minimal steps, expected/actual result, exact message, and a few relevant log lines. Provide a synthetic sample, never Keys or real client documents. The trilingual template is .github/ISSUE_TEMPLATE/bug_report.md.`));
add('appendix',T(`### 快捷键

| 按键 | 作用与条件 |
|---|---|
| Ctrl+C / Ctrl+V / Ctrl+S | 复制 粘贴 保存；先核对焦点 |
| Ctrl+Z | 仅按宿主支持恢复，不适用于 WPS 表格工具写入 |
| Alt 然后 OT 然后 K | WPS 功能区技巧导航；再按自己设置的访问字母 |
| 自定义 Ctrl+Alt 组合 | 侧栏非输入框获得焦点时只打开工具，不直接改数据 |
| F9 或右键更新域 | 在宿主中更新目录；选择更新整个目录 |

### 文件存放位置

安装目录因安装方式不同而变化，本机常见为 %LOCALAPPDATA%/Programs/WpsContentWorkbench。应用代码与 runtime 在该目录，generated 保存用户数据和生成文件。常见文件有 generated/settings.json、generated/ai.json、generated/masking.json；卡片和粘贴栏由本机存储管理。备份位置可在设置中更改。日志在 service.log、service-error.log。不要把代码目录当作文档唯一备份。

### 许可和版本

第三方许可见安装目录 THIRD-PARTY-NOTICES.txt 与 WPS-SDK-LICENSE.txt；使用素材还要遵守素材自己的许可。本手册生成使用项目已有 docx 与 pdfkit，不额外安装 npm 依赖。

0.16.0 包含校对、数据工具、邮编、文体、宏检查等；本次续做加入 MCP 自动脱敏、脱敏复制/还原粘贴、三语手册和日文公式界面补全。发布版本号以安装程序显示和仓库发布记录为准。本手册原稿为 docs/manual/zh.md、ja.md、en.md；构建命令为 node scripts/docs/build-manual.mjs，输出 release/docs。`,
`### キーボード

| キー | 動作と条件 |
|---|---|
| Ctrl+C / Ctrl+V / Ctrl+S | コピー 貼り付け 保存です。フォーカスを確認します |
| Ctrl+Z | ホスト対応の編集だけを戻します。WPS 表格のツール書き込みには使えません |
| Alt 次に OT 次に K | WPS の便利機能ナビ、その後に登録した文字を押します |
| 自分で設定した Ctrl+Alt | サイドバーの入力欄以外にフォーカスがあるとき、ツールを開くだけです |
| F9 または右クリックで更新 | 目次全体を更新します |

### 保存先

インストール先は導入方法で異なります。この PC の例は %LOCALAPPDATA%/Programs/WpsContentWorkbench です。コードと runtime、その下の generated に利用者データと生成ファイルがあります。主なファイルは generated/settings.json、generated/ai.json、generated/masking.json です。カードとトレイはローカルストレージで管理します。バックアップ先は設定で変更できます。ログは service.log と service-error.log です。コードの場所を文書の唯一のバックアップにしないでください。

### ライセンスと版

第三者ライセンスは THIRD-PARTY-NOTICES.txt と WPS-SDK-LICENSE.txt を確認します。素材自体の条件も守ってください。この手冊は既存の docx と pdfkit で生成し、新しい npm 依存関係は追加しません。

0.16.0 には校正・データ・郵便番号・文体・マクロチェックなどがあります。今回 MCP 自動マスキング、コピーと復元貼り付け、三言語説明書、数式画面の日本語補完を追加しました。リリース番号はインストーラーとリポジトリを確認してください。原稿は docs/manual/zh.md、ja.md、en.md、ビルドは node scripts/docs/build-manual.mjs、出力は release/docs です。`,
`### Keyboard reference

| Key | Action and condition |
|---|---|
| Ctrl+C / Ctrl+V / Ctrl+S | Copy paste save; check focus first |
| Ctrl+Z | Host-supported undo only, not WPS spreadsheet tool writes |
| Alt then OT then K | WPS navigator, followed by your assigned access letter |
| Custom Ctrl+Alt combination | Opens a tool when the sidebar has focus outside an input; does not edit data |
| F9 or right-click update field | Update the whole contents table in the host |

### File locations

Installation paths vary. A common local example is %LOCALAPPDATA%/Programs/WpsContentWorkbench. Code and runtime live there; generated holds user data and outputs. Common files are generated/settings.json, generated/ai.json, and generated/masking.json. Cards and paste-bar items are managed locally. Backup location is configurable. Logs are service.log and service-error.log. Do not treat the code directory as your only document backup.

### Licences and versions

Read THIRD-PARTY-NOTICES.txt and WPS-SDK-LICENSE.txt in the installation, and respect separate asset licences. This manual uses the project’s existing docx and pdfkit packages without new npm dependencies.

0.16.0 includes proofing, data, postal, prose-style, and macro-checking tools. This continuation adds MCP masking, masked copy/restored paste, trilingual manuals, and Japanese formula-dialog completion. Check the installer and repository releases for the published version number. Sources are docs/manual/zh.md, ja.md, and en.md. Build with node scripts/docs/build-manual.mjs; output goes to release/docs.`));
module.exports=C;
