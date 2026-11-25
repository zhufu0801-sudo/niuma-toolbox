const {T,docs,add,feature,defaultUndo}=require('./manual-core.cjs');
const p=(c,z,j,e)=>add(c,T(z,j,e));
p('intro',`这本手册面向第一次使用办公工具箱的人。按正在使用的 WPS 文字、表格或演示查找功能，也可以直接搜索目录中的按钮名。每个功能依次说明用途、步骤、选项、限制和恢复方法。英文版是阅读材料；产品界面目前提供中文和日文。

普通排版、数据整理、模板、PDF 和脱敏在本机执行，不需要 AI。只有你设置并主动调用 AI 服务后，提交的内容才会发给所选服务商；外部 agent 和网页 AI 的使用也要分别理解数据流，见数据脱敏章。更新邮编、检查版本及初次获取运行组件可能联网，它们不等于上传文档。

开始前先保存一份小样本。文字选区、活动单元格、当前幻灯片决定许多工具的作用位置；“全部”“批量”及新文件操作可能超出当前选区。先读确认信息，再保存结果。工具箱不自动发送传真或邮件，也不自动让打印机出纸。

本手册对应 0.16.0 版。Mac 版和 Microsoft Office 版还没有完整实测，遇到问题请在 GitHub Issues 反馈。`,
`初めてオフィスツールボックスを使う方向けのガイドです。WPS Writer、Spreadsheets、Presentation の章から探すか、目次のボタン名を検索してください。各機能は用途・手順・選択項目・制限・戻し方の順に説明します。英語版は説明書であり、製品の画面言語は現在中国語と日本語です。

通常の書式設定、データ整理、テンプレート、PDF、マスキングはローカルで動き、AI は不要です。AI サービスを設定して自分で実行した場合だけ、送信内容を選んだサービスへ渡します。外部 agent と Web AI の経路はマスキングの章を確認してください。郵便番号データの更新、バージョン確認、初回の実行環境取得では通信する場合がありますが、文書のアップロードとは別です。

最初は小さなサンプルを保存して試してください。文字の選択範囲、アクティブセル、現在のスライドが処理位置を決めます。「すべて」「一括」や新規ファイルの操作は選択範囲だけに限定されない場合があります。確認内容を読んでから実行し、結果を確認して保存します。FAX・メールの送信やプリンターへの出力は自動で行いません。

本ガイドはバージョン 0.16.0 に対応しています。Mac 版と Microsoft Office 版は実機での検証が完了していません。問題がある場合は GitHub Issues へご報告ください。`,
`This manual is for first-time Office Toolbox users. Start with the chapter for WPS Writer, Spreadsheets, or Presentation, or search for a button in the contents. Each feature has five parts: purpose, steps, options, limits, and recovery. The English manual does not imply an English product interface; the current UI languages are Chinese and Japanese.

Ordinary formatting, data cleanup, templates, PDF tools, and masking run locally without AI. Content is sent to an AI provider only when you configure a service and actively use it. External agents and web AI have separate data paths, explained in the masking chapter. Postal-data updates, version checks, and initial runtime downloads can use the network without uploading documents.

Begin with a saved, small sample. Text selections, the active cell, and the current slide determine many operation targets. Whole-document, batch, and file-generation operations may extend beyond the selection. Read confirmations, inspect results, and then save. The toolbox does not automatically send fax or email or start physical printing.

This manual covers version 0.16.0. The Mac and Microsoft Office editions have not completed full testing on real machines. Please report problems through GitHub Issues.`);
p('install',`### Windows

1. 完整解压适合系统的 Windows 64 位安装包，不要在 ZIP 内直接运行。
2. 运行 Start.cmd，选择 WPS、Microsoft Office 或两者。把安装目录放在当前用户可以写入且以后不随意移动的位置。
3. Office 安装会请求信任仅供本机 localhost 使用的证书。核对名称，由你本人决定；企业受管电脑由 IT 处理。
4. 保存文档并完全退出 WPS 或 Office，再打开。WPS 中从“办公工具箱”打开侧栏；Office 中从对应加载项入口打开。
5. 每次开机后先运行桌面 Office Toolbox 启动本机服务。未设置自动开机启动。

升级前备份安装目录中的 generated 和重要文档。更新程序文件后重新执行安装登记，再完全关闭所有 WPS 窗口重开，避免旧侧栏继续使用缓存脚本。

### macOS

1. 解压 Mac 安装包，在访达中右键 install.command，选择打开。
2. 按提示选择 WPS 或 Microsoft Office。缺少 Node.js 时安装脚本会询问是否从 nodejs.org 获取。
3. Office 的 localhost 证书信任由本人或 IT 决定；系统可能请求登录密码。
4. 完全退出并重开宿主。Office 可从“插入 → 加载项 → 我的加载项”进入。

Mac 的 WPS 加载项支持取决于版本。系统 OCR 中文需要 macOS 11 或以后，日文需要 macOS 13 或以后；PDF 图片提取等依赖图像组件的分支可能不可用。本轮没有 Mac 真机验收。

### 企业批量部署

按每个用户部署，不把一个后台服务当作多用户服务器。Initialize.ps1 支持 -Targets wps、office 或 wps,office；-NoLaunch 不自动打开，-NoShortcut 不创建桌面快捷方式。Office 证书确认意味着它不一定能完全静默安装。端口固定为 127.0.0.1:39127，Office 另用 39128，同机多个并发登录用户可能冲突。

preset/settings.json 可提供界面和字体设置；preset/ai.json 可提供 AI 地址与模型；preset/plugins 可提供公司插件。预设只在目标文件不存在时复制，不能当作强制覆盖策略。企业部署细节见安装包企业部署说明；先在测试账号演练。

### 卸载与保留数据

Windows 运行 Uninstall.cmd；脚本会移除本工具箱登记、自己的证书和快捷方式，并删除 generated/ai.json 内保存的 AI 设置与 Key。卡片、普通设置等保留。Mac 运行 uninstall.command，按该脚本的提示处理保留数据。不要把卸载等同于彻底擦除个人文档；确认备份后再决定是否删除保留目录。`,
`### Windows

1. Windows 64 ビット用パッケージをすべて展開します。ZIP の中から直接実行しません。
2. Start.cmd を実行して WPS、Microsoft Office、または両方を選びます。現在のユーザーが書き込める固定の場所を使ってください。
3. Office では localhost 専用証明書の信頼確認が出ます。名前を確かめ、本人が判断します。管理対象の PC は IT 担当者に依頼してください。
4. 文書を保存して WPS または Office を完全に終了し、開き直します。WPS の「オフィスツールボックス」、または Office のアドインから入ります。
5. PC の起動後はデスクトップの Office Toolbox を開き、ローカルサービスを起動します。自動起動は設定されません。

更新前に generated と重要な文書をバックアップします。プログラムを更新して登録をやり直した後、WPS の全ウィンドウを閉じて再起動してください。古いサイドバーにはスクリプトが残ることがあります。

### macOS

1. Mac 用パッケージを展開し、Finder で install.command を右クリックして開きます。
2. WPS または Microsoft Office を選びます。Node.js がなければ、nodejs.org から取得するか確認されます。
3. Office の localhost 証明書を信頼するかは本人または IT 担当者が判断します。ログインパスワードを求められる場合があります。
4. ホストを完全に終了して開き直します。Office は「挿入 → アドイン → 個人用アドイン」から開けます。

Mac WPS のアドイン対応は版によって異なります。システム OCR の中国語は macOS 11 以降、日本語は 13 以降が必要です。PDF 画像抽出など、画像コンポーネントが必要な処理は使えない場合があります。今回 Mac 実機では検証していません。

### 企業への一括導入

ユーザーごとに導入します。バックグラウンドサービスを複数ユーザーの共用サーバーとして扱わないでください。Initialize.ps1 は -Targets wps、office、wps,office に対応します。-NoLaunch は自動で開かず、-NoShortcut はショートカットを作りません。Office は証明書確認があり、完全な無人導入にはならない場合があります。固定ポートは 127.0.0.1:39127、Office は追加で 39128 を使います。同時ログインする複数ユーザーでは競合する場合があります。

preset/settings.json は画面やフォント、preset/ai.json は AI の接続先とモデル、preset/plugins は社内プラグインの初期設定に使います。対象ファイルが存在しない場合だけコピーするため、強制上書きの仕組みではありません。パッケージ内の企業向け導入ガイドを読み、テストアカウントで確認してください。

### アンインストールと残るデータ

Windows では Uninstall.cmd を実行します。ツールボックスの登録、専用証明書、ショートカットを削除し、generated/ai.json の AI 設定と Key も削除します。カードや通常の設定は残ります。Mac は uninstall.command の案内に従います。アンインストールは個人文書の完全消去ではありません。残るフォルダーを削除する前にバックアップを確認してください。`,
`### Windows

1. Extract the complete Windows 64-bit package. Do not run it inside a ZIP.
2. Run Start.cmd and select WPS, Microsoft Office, or both. Keep the installation in a stable directory writable by the current user.
3. Office asks you to trust a certificate used only for localhost. Check its name and decide yourself; ask IT on managed devices.
4. Save documents, fully exit WPS or Office, and reopen it. Open Office Toolbox from the WPS ribbon or the Office add-in entry.
5. After starting the computer, launch the desktop Office Toolbox shortcut to start its local service. Installation does not configure automatic startup.

Before upgrading, back up generated and important documents. Update program files, repeat registration, and fully close all WPS windows before reopening them. An old sidebar can retain cached scripts.

### macOS

1. Extract the Mac package. Right-click install.command in Finder and open it.
2. Select WPS or Microsoft Office. If Node.js is missing, the installer asks whether to obtain it from nodejs.org.
3. You or IT decide whether to trust the Office localhost certificate. macOS may ask for your login password.
4. Fully quit and reopen the host. In Office, use Insert, Add-ins, My Add-ins if necessary.

Mac WPS add-in support varies by version. System OCR needs macOS 11 or later for Chinese and 13 or later for Japanese. Features such as PDF image extraction may be unavailable without their image component. This round did not include a Mac desktop test.

### Enterprise deployment

Deploy per user; the background service is not a shared multi-user server. Initialize.ps1 accepts -Targets wps, office, or wps,office. -NoLaunch suppresses opening the app and -NoShortcut suppresses the desktop shortcut. Office certificate confirmation can prevent a fully unattended deployment. Fixed ports are 127.0.0.1:39127 and, for Office, 39128. Concurrent signed-in users on one computer can conflict.

preset/settings.json supplies initial UI and font settings; preset/ai.json supplies AI connection and model settings; preset/plugins supplies company plugins. Presets copy only when the destination file does not exist, so they do not enforce overwrites. Read the enterprise guide in the package and pilot with a test account.

### Removal and retained data

On Windows, run Uninstall.cmd. It removes the toolbox registrations, its own certificates and shortcuts, and generated/ai.json containing saved AI settings and the Key. Cards and ordinary settings remain. On Mac, follow uninstall.command. Uninstalling does not erase personal documents; verify backups before deleting retained folders.`);
p('interface',`### 从哪里开始

功能区按当前宿主显示相关工具。同一功能可能同时出现在大按钮、菜单和侧栏，这些是同一个操作入口。“常用”页放搜索、技巧与常用工具；“创建”页输入表格或大纲；“卡片”页管理可复用内容；“文件”页集中 PDF、模板、打印、脱敏等面板。屏幕较窄时向下滚动，展开折叠标题。

技巧导航先输入需求，如“邮编”“脱敏复制”“表格太宽”。结果可能是可执行工具，也可能是操作指南：指南只展示步骤，不替你修改文件。使用星标收藏后，可在“我的收藏”和功能区收藏组进入；取消星标移除收藏。“选择要显示的功能”只隐藏入口，不删除文件或卸载功能。

### 设置与第一次试用

1. 打开“设置”，选择中文、日本語或跟随系统。
2. 查看正文默认字体、字号、行距，表头颜色和卡片数量设置；这些只影响使用相应工具时的输出。
3. 查看备份目录与自动备份选项，先手动保存文档，再用“备份当前文件”确认位置。
4. AI 地址、模型和 Key 只有要用 AI 时才填写。普通工具无需开通 AI。
5. 更新后若版本不一致，保存文档、关闭所有 WPS 窗口再开。只关侧栏不一定刷新功能区脚本。

语言切换不会翻译你输入的文档、文件名或自定义词表。示例表名不是必须存在的工作表；读取真实选区可带入自己的表名。诊断摘要用于反馈环境；版本信息 URL 未配置时不代表安装失败，工具箱不自动下载安装更新。`,
`### 最初に開く場所

リボンには現在のホストに合う機能を表示します。大きなボタン、メニュー、サイドバーに同じ機能がある場合も、処理は同じです。「よく使う」には検索と便利機能、「作成」には表やアウトラインの入力、「カード」には再利用する内容、「ファイル」には PDF・テンプレート・印刷・マスキングのパネルがあります。幅が狭い場合は下へスクロールし、折りたたみ見出しを開いてください。

便利機能ナビで「郵便番号」「マスキング」「表が広い」などの目的を検索します。実行できるツールと操作ガイドがあり、ガイドは手順を示すだけでファイルを変更しません。星を付けるとお気に入りとリボンのグループから開けます。星を外すと登録を解除します。「表示する機能を選ぶ」は入口を隠すだけで、ファイルや機能を削除しません。

### 設定と最初の試用

1. 設定を開き、中国語・日本語・システムに従うから選びます。
2. 本文のフォント・サイズ・行間、表見出しの色、カード件数を確認します。対応するツールを使ったときの出力に適用します。
3. バックアップ先と自動バックアップを確認します。文書を保存してから「現在のファイルをバックアップ」で保存先を確かめます。
4. AI を使うときだけ接続先・モデル・Key を設定します。通常のツールには AI は不要です。
5. 更新後に版の不一致が出たら保存し、WPS の全ウィンドウを閉じて開き直します。サイドバーを閉じるだけではリボンのスクリプトが更新されない場合があります。

言語を切り替えても、文書・ファイル名・単語リストの内容は翻訳しません。例のシート名を実際に作る必要はありません。選択範囲を読み取ると自分のシート名を使えます。診断情報は環境の報告に使います。バージョン情報 URL の未設定はインストール失敗ではなく、更新の自動ダウンロードや自動インストールは行いません。`,
`### Where to start

The ribbon shows tools for the current host. A large button, menu item, and sidebar item may lead to the same operation. Common tools contains search and guidance; Create accepts tables and outlines; Cards manages reusable content; Files contains PDF, templates, printing, masking, and related panels. Scroll down and expand collapsed headings in a narrow sidebar.

Search the navigator by a task, such as postal lookup, masking, or a wide table. Results can be executable tools or guides. A guide explains steps without changing your file. Star an item to add it to favourites and the ribbon favourites group; remove the star to remove it. Choosing which functions to display hides entries without deleting files or uninstalling functionality.

### Settings and a first trial

1. Open Settings and choose Chinese, Japanese, or follow system.
2. Review default body font, size, spacing, table-header colours, and card count. These affect the relevant tool output.
3. Check the backup directory and automatic-backup setting. Save the document and run Back up current file to confirm the location.
4. Enter an AI endpoint, model, and Key only if you want AI. Ordinary tools do not require it.
5. If versions disagree after an update, save and close every WPS window before reopening. Closing only the sidebar may leave old ribbon scripts active.

Changing the UI language does not translate document content, filenames, or custom word lists. Example sheet names do not need to exist: capture your real selection to use its actual sheet name. Diagnostic summaries help report the environment. An unset version-information URL does not mean installation failed; the toolbox does not download or install updates automatically.`);
p('word',`### 选区与结构

先用标题样式建立结构，再整理正文，最后更新目录。页面看上去像标题不等于它有标题级别。目录使用一至三级标题；手工输入页码不会自动更新。正文整理会跳过识别出的标题，但复杂模板仍应先用副本验证。

快速文字替换、空段落清理和文体转换针对支持的正文范围；页眉、域、表格、修订状态可能限制操作。只调整图片时，先辨别浮动图与嵌入型图片。打印身份证尺寸要先裁掉照片背景，不能用整张拍摄照片的外框当证件边缘。`,
`### 選択範囲と文書構造

見出しスタイルで構造を作り、本文を整え、最後に目次を更新します。見出しに見える文字でも見出しレベルが付いているとは限りません。目次はレベル 1～3 を使い、手入力のページ番号は更新されません。本文整理では認識した見出しを除外しますが、複雑なテンプレートはコピーで試してください。

簡易置換、空段落削除、文体変換は対応する本文範囲に限ります。ヘッダー・フィールド・表・変更履歴に制限があります。画像は浮動型と埋め込み型を区別します。カードサイズで印刷する写真は先に背景を切り取り、撮影写真全体の外枠をカードの端として扱わないでください。`,
`### Selection and structure

Apply heading styles first, format body text next, and update the contents last. Text that looks like a heading may have no heading level. The contents tool uses levels 1–3; manually typed page numbers do not update. Body formatting skips recognized headings, but complex templates still need a trial copy.

Quick replacement, blank-paragraph cleanup, and prose conversion operate only on supported body ranges. Headers, fields, tables, and tracked changes can restrict them. Distinguish floating from inline pictures. Crop the background before setting card dimensions; the outer edge of a photographed scene is not the edge of the card.`);
p('sheet',`### 先分清数值 文字与显示格式

号码、邮编和身份证不是金额。长编号应保留文本，已经因数值格式丢失的尾数无法靠“文本转数字”找回。“万元显示”只改显示；“金额大写”产生文字；“排名不打乱顺序”写公式。结果在右侧时先留空列，新表结果与原表是否联动要看具体功能说明。

WPS 表格的工具写入不能用 Ctrl+Z 恢复。自动备份只复制磁盘上已保存的文件，不包含还未保存的修改。运行前保存并手动备份一次，特别是删除行、批量替换、改名和清空数据。`,
`### 数値 文字列と表示形式

番号・郵便番号・個人番号は金額ではありません。長い番号は文字列として残します。数値化で失った末尾の桁は変換ツールで戻せません。大きな金額の単位表示は表示だけを変え、金額の大字化は文字列を作り、順位付けは数式を書きます。右側に結果を書く処理では空き列を用意します。結果が元データに連動するかは機能ごとに確認してください。

WPS Spreadsheets のツール書き込みは Ctrl+Z では戻せません。自動バックアップはディスクに保存済みのファイルをコピーするため、未保存の変更は含みません。特に行削除・置換・改名・データ消去の前に保存し、手動バックアップを作ってください。`,
`### Values text and display formats

IDs and postal codes are not amounts. Keep long IDs as text; digits already lost by numeric storage cannot be recovered by conversion. Large-number units change display only, uppercase amounts produce text, and ranking writes formulas. Reserve empty output columns where required. Whether a result tracks source changes depends on the individual tool.

WPS spreadsheet tool writes cannot be undone with Ctrl+Z. Automatic backup copies the saved file on disk and excludes unsaved changes. Save and make a manual backup before deleting rows, replacing text, renaming, or clearing data.`);
p('slides',`### 整份演示与当前页

“统一本页字体”和“复制当前页”依赖当前幻灯片。“全部页统一字体”“删除全部动画”“删除全部备注”处理整份演示；先另存副本。套用模板不等于保证所有文本都能放进新布局，完成后检查标题、表格、图片和溢出的长文本。导出图片和 PDF 适合分享，但不是可编辑的 PPT 原稿。`,
`### 全体と現在のスライド

現在のスライドの書式や複製は選択中のページを使います。全ページのフォント統一、アニメーション削除、ノート削除はプレゼン全体が対象です。先に別名保存してください。テンプレート適用後も文字が必ず収まるわけではないため、タイトル・表・画像・長文のはみ出しを確認します。画像や PDF の出力は共有用で、編集可能な PPT 原稿ではありません。`,
`### Whole deck or current slide

Current-slide formatting and duplication use the active slide. Unifying all fonts, removing animations, and removing notes affect the whole deck: save a copy first. Applying a template does not guarantee every text block will fit. Inspect titles, tables, images, and overflowing text. Image and PDF exports are sharing formats, not editable presentation originals.`);
p('files',`### 文件路径和保存内容

文件工具读取磁盘中的内容。先保存文档；完整路径必须指向实际文件，不能只填文件名。输出通常在“生成文件”目录或你选的文件夹，完成提示提供位置或“打开所在文件夹”。生成副本后先确认文件名再继续编辑，避免把旧原稿误当作处理结果。

PDF 页面工具上限：最多合并 50 个文件，单个 PDF 200 MB，总量 500 MB，最多 2000 页；图片转 PDF 最多 200 张、每张 40 MB；PDF 渲染为图片一次最多 100 页。密码或编辑限制应先由有权限的人在原软件中处理；工具箱不绕过保护。`,
`### パスと保存済みの内容

ファイルツールはディスク上の内容を読みます。先に文書を保存し、ファイル名だけでなく実在する完全パスを指定してください。結果は生成ファイルの場所または選んだフォルダーに保存し、完了メッセージや「保存先フォルダーを開く」で確認できます。処理済みのコピーを開いたことを確かめてから編集を続けます。

PDF ページ処理の上限は 50 ファイル、1 PDF 200 MB、合計 500 MB、2000 ページです。画像から PDF は 200 枚、1 枚 40 MB まで、PDF の画像化は一度に 100 ページまでです。パスワードや編集制限は権限のある人が元のソフトで扱います。保護を回避する機能ではありません。`,
`### Paths and saved content

File tools read content saved on disk. Save first and provide a full path to an existing file, not just its name. Results normally go to Generated files or a selected folder. Use the result path or Open containing folder, and verify that you opened the processed copy before continuing to edit.

PDF page tools allow up to 50 merge inputs, 200 MB per PDF, 500 MB total, and 2,000 pages. Image-to-PDF accepts up to 200 images at 40 MB each. Rendering PDF pages to images is limited to 100 pages per operation. An authorized person must handle password or editing restrictions in the source application; the toolbox does not bypass protection.`);
p('privacy',`### 数据在哪里流动

| 数据或操作 | 保存或发送位置 | 你需要注意的事 |
|---|---|---|
| 普通排版和表格整理 | 当前宿主及本机 | 保存原稿，核对改动范围 |
| 卡片 粘贴栏 设置和生成文件 | 安装目录下 generated 及设置的输出/备份位置 | 本机存储不代表文件本身都被加密 |
| 脱敏对应表和词表 | generated/masking.json | 内容用 AES-256-GCM；钥匙由 Windows DPAPI 或 Mac 钥匙串保管 |
| MCP 对应表 | 服务进程内存 | 会话闲置两小时或服务重启后清掉 |
| 内置 AI 请求 | 你选的模型服务商 | 发送前核对保护模式和实际输入 |
| 网页 AI | 你自己粘贴到的网站 | 工具箱不能拦截网站自己的读取或历史保存 |

DPAPI 使用当前 Windows 用户保护；复制 masking.json 到另一台电脑并不保证可解密。Mac 依赖钥匙串可访问。系统密钥不可用时应保留原文件并排查，不要删除旧对应表来“修复”。导出或分享脱敏文件前检查正文、页眉、页脚、批注以及另附的资料；自动识别不保证找到所有客户信息。

AI 保护可以隐藏已支持的号码，客户名等自定义内容需要词表。公司插件会参与 AI 发送流程；插件加载失败时 AI 停用。不要把“本机工具”理解为已经发送给 AI 的内容仍只在本机。`,
`### データの流れ

| データや操作 | 保存先または送信先 | 確認する点 |
|---|---|---|
| 通常の書式設定と表整理 | ホストとローカル PC | 原稿と処理範囲を確認します |
| カード トレイ 設定 生成ファイル | generated と指定した出力・バックアップ先 | ローカル保存のすべてが暗号化されるわけではありません |
| 対応表と単語リスト | generated/masking.json | AES-256-GCM、鍵は Windows DPAPI または Mac キーチェーンで保管します |
| MCP 対応表 | サービスのメモリー | 2 時間使わない場合や再起動で消えます |
| 内蔵 AI の要求 | 自分で選んだモデル提供元 | 送信前の保護モードと入力を確認します |
| Web AI | 自分で貼り付けた Web サイト | サイト独自の読み取りや履歴保存は制御できません |

DPAPI は現在の Windows ユーザーで保護します。masking.json を別の PC にコピーしても復号できるとは限りません。Mac はキーチェーンへのアクセスが必要です。鍵が使えないときは元ファイルを残して調べ、修復のつもりで古い対応表を削除しないでください。共有前には本文・ヘッダー・フッター・コメントと添付資料を確認します。自動検出ですべての顧客情報を見つける保証はありません。

AI 保護は対応する番号を隠します。顧客名などは単語リストへ登録してください。社内プラグインも送信処理に入り、読み込みに失敗すると AI を停止します。ローカルツールであることは、AI に送った内容がローカルに留まるという意味ではありません。`,
`### Data flow

| Data or action | Location | What to check |
|---|---|---|
| Ordinary formatting and spreadsheet cleanup | Host and local computer | Retain the original and inspect the target scope |
| Cards paste bar settings and outputs | generated and configured output or backup locations | Local storage does not mean every file is encrypted |
| Masking maps and word lists | generated/masking.json | AES-256-GCM with a key protected by Windows DPAPI or Mac Keychain |
| MCP maps | Service memory | Cleared after two idle hours or a service restart |
| Built-in AI requests | Your chosen model provider | Inspect protection mode and actual input before sending |
| Web AI | The website where you paste | The toolbox cannot control the site's own collection or history |

DPAPI protects data for the current Windows user. Copying masking.json to another computer does not guarantee decryption. Mac requires Keychain access. If the system key is unavailable, retain the original files while troubleshooting; do not delete old maps as a repair step. Before sharing, inspect body text, headers, footers, comments, and accompanying materials. Automatic detection cannot guarantee that every piece of client information is found.

AI protection hides supported number patterns; register client names and other custom terms in the word list. Company plugins also participate in AI sending, and plugin-load failure disables AI. Running a local toolbox does not mean content already sent to an AI provider remains local.`);
p('limits',`### 恢复路径对照

| 情况 | 推荐恢复方法 |
|---|---|
| WPS 文字普通编辑 | 立即检查 Ctrl+Z 的结果；宿主未提供完整撤销时恢复副本 |
| WPS 表格工具改写 | 从操作前保存的备份恢复；不能依赖 Ctrl+Z |
| 演示全部页操作 | 保留原稿并另存副本；部分接口不能撤销 |
| 生成新 PDF 新文档或图片 | 原文件保留；核对后删除不需要的结果 |
| 公式助手写入 | 用“最近写入与撤回”，最多 10 次；仅移除仍与记录一致的公式 |
| 文件改名 | 用“改回上一次”，依赖文件仍可匹配记录 |
| 脱敏副本 | 用对应表复原；没有对应表时不能从星号或黑块猜回原文 |

自动备份是额外保护，不替代保存。磁盘备份可能早于当前未保存内容。对一份文件连续做多个操作时，每步检查后再进行下一步，不要同时切换选区或更换文档。

### 常见上限

共用正文选区为 5 万字符；基础表格范围不超过 1000 行或 50 列，部分工具要求更小范围。快速表格输入 20 行、10 列、1 万字符。脱敏复制最多 20 万字符、表格最多 2 万单元格。MCP 读取可能截断，应看返回的 truncated 标志。大文件先拆分；不要把超过上限后的部分结果当成全量处理。

文档里的字体需要本机安装；不同宿主和操作系统可能替换字体、重排分页。实际打印用预览核对纸张、方向、缩放和页码。界面切换只改按钮和提示，不改文档正文语言。`,
`### 戻し方の対照

| 状況 | 戻し方 |
|---|---|
| WPS Writer の通常編集 | すぐに Ctrl+Z の結果を確認し、不完全ならコピーから戻します |
| WPS Spreadsheets のツール書き込み | 保存したバックアップから戻します。Ctrl+Z に依存しません |
| Presentation の全ページ操作 | 原稿を残し、別名のコピーで作業します。一部は元に戻せません |
| 新しい PDF 文書 画像 | 元ファイルを残し、不要な出力を確認して削除します |
| 数式アシスタント | 最近の書き込みの取り消しを使います。最大 10 回で、記録と一致する数式だけが対象です |
| ファイル名変更 | 前回の名前に戻す機能を使います。記録と現在のファイルが一致する必要があります |
| マスキングのコピー | 対応表で復元します。表がなければ伏字から原文は推測できません |

自動バックアップは保存の代わりではありません。ディスク上のバックアップには未保存の変更が含まれない場合があります。複数の操作は一つずつ確認し、途中で文書や範囲を切り替えないでください。

### 主な上限

共通の本文選択は 5 万文字、基本のセル範囲は 1000 行・50 列までで、個別に小さい上限があります。簡易表入力は 20 行・10 列・1 万文字です。マスキングコピーは 20 万文字、セルは 2 万個までです。MCP の読み取りは切り詰められる場合があるため truncated を確認します。大きい入力は分割し、部分結果を全件処理と扱わないでください。

文書で使うフォントは PC に必要です。ホストや OS により代替フォントと改ページが変わる場合があります。印刷前に用紙・向き・倍率・ページ番号を確認してください。言語切り替えは画面表示を変えるだけで、本文の言語を変えません。`,
`### Recovery paths

| Situation | Recovery |
|---|---|
| Ordinary WPS Writer edits | Try Ctrl+Z immediately and inspect; use a retained copy if host undo is incomplete |
| WPS spreadsheet tool writes | Restore the saved backup; do not rely on Ctrl+Z |
| Whole-deck changes | Work in a saved copy; some operations have no undo |
| New PDF document or image | Keep the source and delete an unwanted output after checking it |
| Formula assistant writes | Use recent-write rollback, up to 10 records; only unchanged recorded formulas are removed |
| File renaming | Use the last-rename recovery feature while current files still match its record |
| Masked copies | Restore with the map; asterisks or black blocks cannot reconstruct missing originals |

Automatic backup does not replace saving. A disk backup may predate unsaved edits. Apply multiple operations one at a time and check each result; do not change documents or selections during processing.

### Common limits

Shared body-text selection is limited to 50,000 characters. Basic cell ranges are limited to 1,000 rows and 50 columns, with smaller limits in some tools. Quick table input allows 20 rows, 10 columns, and 10,000 characters. Masked copy allows 200,000 characters and 20,000 cells. MCP reads may be truncated: inspect the truncated flag. Split large inputs and do not treat partial output as a complete result.

Fonts used in a document must be available locally. Other hosts or operating systems can substitute fonts and repaginate. Before printing, check paper, orientation, scale, and page numbers. UI language switching changes buttons and messages, not the language of document content.`);
module.exports=require('./manual-core.cjs');
