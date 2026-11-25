// 申请 API Key 图文步骤（中文 / 日文各写一份）。官网按钮名称可能改版，以官网为准；地址核对于 2026-11-11。
(function(){
'use strict';
var P={
 deepseek:{url:'https://platform.deepseek.com/api_keys',name:'DeepSeek',model:'deepseek-flash',
  zh:['打开右边的官网链接，用手机号或邮箱注册并登录。','在左侧菜单进入「API keys」，点击「创建 API key」，起个名字（例如“办公工具箱”）。','复制生成的 Key（以 sk- 开头，只显示一次，请立即复制）。','在「充值」页面充值少量金额即可开始使用，按用量计费。'],
  ja:['右のリンクから公式サイトを開き、電話番号またはメールで登録・ログインします。','左メニューの「API keys」で「Create API key」をクリックし、名前（例：Office Toolbox）を付けます。','表示された Key（sk- で始まる）をコピーします。一度しか表示されないので、すぐにコピーしてください。','「Top up」でチャージすると使えます（従量課金）。']},
 qwen:{url:'https://bailian.console.aliyun.com/model/settings/api-key',name:'通义千问 / Qwen',model:'qwen-plus',
  zh:['打开阿里云百炼的 API Key 页面，用阿里云账号登录（首次使用需要实名认证并开通百炼服务，新用户通常有免费额度）。','点击「创建 API-KEY」，选择默认业务空间。','复制生成的 Key。','注意：Key 区分地域。工具箱默认用“北京”地域的地址，请在北京地域创建 Key。'],
  ja:['Alibaba Cloud Model Studio（百錬）の API Key ページを開き、Alibaba Cloud アカウントでログインします（初回は本人確認とサービスの有効化が必要です）。','「API-KEY を作成」をクリックし、既定のワークスペースを選びます。','作成された Key をコピーします。','Key はリージョンごとに分かれています。ツールボックスの既定は「北京」リージョンの接続先です。']},
 zhipu:{url:'https://bigmodel.cn/usercenter/proj-mgmt/apikeys',name:'智谱 GLM',model:'glm-5.3',
  zh:['打开智谱开放平台，注册并登录。','进入「API Keys」页面，点击「添加新的 API Key」。','复制生成的 Key。','新用户一般有体验额度；用完后在「财务」里充值。'],
  ja:['Zhipu のオープンプラットフォームを開き、登録してログインします。','「API Keys」ページで新しい API Key を追加します。','作成された Key をコピーします。','無料枠を使い切ったら、料金ページでチャージしてください。']},
 kimi:{url:'https://platform.kimi.com/',name:'Kimi',model:'',
  zh:['打开 Kimi 开放平台，注册并登录。','进入控制台的「API Key 管理」，新建一个 Key。','复制生成的 Key，按需充值。','模型名会更新：保存 Key 后，在工具箱里点“获取模型列表”选择。'],
  ja:['Kimi のプラットフォームを開き、登録してログインします。','コンソールの API Key 管理で新しい Key を作成します。','Key をコピーし、必要に応じてチャージします。','モデル名は更新されます。Key を保存したら「モデル一覧を取得」から選んでください。']},
 doubao:{url:'https://console.volcengine.com/ark',name:'豆包 / Doubao',model:'',
  zh:['打开火山引擎控制台的「火山方舟」，注册并实名认证。','在「API Key 管理」中创建 Key 并复制。','在「开通管理」中开通你要用的豆包模型。','保存 Key 后，在工具箱里点“获取模型列表”选择模型。'],
  ja:['Volcengine コンソールの「火山方舟（Ark）」を開き、登録して本人確認をします。','API Key 管理で Key を作成し、コピーします。','使いたい Doubao モデルを有効化します。','Key を保存したら「モデル一覧を取得」からモデルを選びます。']},
 openai:{url:'https://platform.openai.com/api-keys',name:'OpenAI',model:'',
  zh:['打开 OpenAI 开发者平台并登录（与 ChatGPT 账号可通用，但 API 单独计费）。','点击「Create new secret key」，复制 Key（以 sk- 开头，只显示一次）。','在「Billing」里绑定信用卡并预充值。','保存 Key 后，点“获取模型列表”选择模型。注意：OpenAI 不向中国大陆提供服务。'],
  ja:['OpenAI の開発者プラットフォームを開いてログインします（ChatGPT と同じアカウントで可。API は別料金です）。','「Create new secret key」をクリックし、Key（sk- で始まる、一度だけ表示）をコピーします。','「Billing」でクレジットカードを登録し、前払いでチャージします。','Key を保存したら「モデル一覧を取得」でモデルを選びます。']},
 claude:{url:'https://platform.claude.com/settings/keys',name:'Anthropic Claude',model:'claude-sonnet-5-5',
  zh:['打开 Claude 开发者平台并登录。','在「API Keys」页面点击「Create Key」，复制 Key（只显示一次）。','在「Billing」里购买额度。','工具箱通过 Claude 的 OpenAI 兼容接口调用；推荐模型 claude-sonnet-5-5。'],
  ja:['Claude の開発者プラットフォーム（Claude Console）を開いてログインします。','「API Keys」で「Create Key」をクリックし、Key をコピーします（一度だけ表示）。','「Billing」でクレジットを購入します。','ツールボックスは Claude の OpenAI 互換 API を使います。おすすめは claude-sonnet-5-5 です。']},
 gemini:{url:'https://aistudio.google.com/apikey',name:'Google Gemini',model:'gemini-3.8-flash',
  zh:['打开 Google AI Studio，用 Google 账号登录。','点击「Create API key」，复制 Key。','有免费额度（有次数限制）；需要更多时在 Google Cloud 开通计费。'],
  ja:['Google AI Studio を開き、Google アカウントでログインします。','「Create API key」をクリックして Key をコピーします。','無料枠があります（回数制限あり）。足りない場合は Google Cloud で課金を有効にします。']},
 ollama:{url:'https://ollama.com/download',name:'Ollama',model:'',
  zh:['下载并安装 Ollama（免费，在自己电脑上运行，不需要 Key，不上传内容）。','在 Ollama 里下载一个模型（例如通义千问 Qwen 系列），电脑内存建议 16GB 以上。','在工具箱里选“本机 Ollama”，点“获取模型列表”选择刚下载的模型。'],
  ja:['Ollama をダウンロードしてインストールします（無料。自分のパソコンで動き、Key 不要、内容は送信されません）。','Ollama でモデルをダウンロードします（例：Qwen シリーズ。メモリ 16GB 以上推奨）。','ツールボックスで「ローカル Ollama」を選び、「モデル一覧を取得」からモデルを選びます。']}
};
var T={
 zh:{title:'申请 AI 的 API Key：一步一步教你',lead:'API Key 就像一把“钥匙”，让办公工具箱用你自己的 AI 账户帮你写 PPT 大纲等。申请只需 3～5 分钟。',
  pickH:'先选一家',pick:[['中国大陆用户','推荐 DeepSeek、通义千问或智谱：注册方便、价格低、国内网络直连。'],['日本及海外用户','推荐 Claude、OpenAI 或 Gemini：日文表现好，需要信用卡（Gemini 有免费额度）。'],['不想付费 / 不能联网','使用本机 Ollama：免费离线，但需要配置较好的电脑。']],
  allH:'通用步骤（每家都一样）',all:['在官网注册并登录（点下面每家卡片里的按钮直达）。','找到“API Keys / API Key 管理”，新建一个 Key。','复制 Key（很多平台只显示一次）。','回到工具箱「设置 → AI 服务」，选服务商，粘贴 Key，点“保存 AI 设置”。','点“测试连接”，看到“连接成功”就可以用了。'],
  provH:'各家申请步骤',go:'打开官网 ›',model:'工具箱默认模型：',modelNone:'保存 Key 后点“获取模型列表”选择',
  note:'安全提醒：API Key 等同于账户密码，不要发给别人、不要贴在群里。工具箱只把它保存在本机，不会上传到我们这里。如果怀疑泄露，到官网删除这个 Key 再新建一个。官网按钮名称可能改版，以官网为准。',
  faqH:'常见问题',faq:[['提示“API Key 不正确”','重新到官网复制一次，注意不要多复制空格；确认 Key 没有被删除。'],['提示“余额不足 / 额度用完”','到服务商官网充值或领取免费额度。工具箱本身不收费。'],['提示“连不上 AI 服务”','检查网络。公司网络可能拦截外部网站，请联系网络管理员；海外服务在中国大陆可能无法访问，改用国内服务商。'],['提示“找不到这个模型”','点“获取模型列表”，从列表里选一个。'],['要花多少钱？','生成一次 PPT 大纲通常只用几千个字（token），国内服务一般不到几分钱，以官网价格为准。']],
  back:'‹ 返回 AI 设置'},
 ja:{title:'AI の API Key を取得する手順',lead:'API Key はツールボックスがあなたの AI アカウントを使うための「鍵」です。PPT の構成案づくりなどに使います。3〜5 分で取得できます。',
  pickH:'まずサービスを選ぶ',pick:[['日本・海外のかた','Claude、OpenAI、Gemini がおすすめです。日本語に強く、クレジットカードが必要です（Gemini は無料枠あり）。'],['中国本土のかた','DeepSeek、通義千問（Qwen）、智譜（GLM）がおすすめです。登録しやすく低価格です。'],['無料・オフラインで使いたい','ローカルの Ollama を使います。無料ですが、性能の高いパソコンが必要です。']],
  allH:'共通の手順',all:['公式サイトで登録・ログインします（下の各カードのボタンから直接開けます）。','「API Keys」の画面で新しい Key を作成します。','Key をコピーします（多くのサービスで一度しか表示されません）。','ツールボックスの「設定 → AI サービス」でサービスを選び、Key を貼り付けて「AI 設定を保存」をクリックします。','「接続テスト」で「接続成功」と出れば完了です。'],
  provH:'サービス別の手順',go:'公式サイトを開く ›',model:'ツールボックスの既定モデル：',modelNone:'Key を保存したら「モデル一覧を取得」から選択',
  note:'セキュリティ：API Key はパスワードと同じです。他人に送ったり、チャットに貼ったりしないでください。ツールボックスはこのパソコンの中だけに保存し、外部には送信しません。漏れた疑いがあるときは、公式サイトでその Key を削除し、作り直してください。ボタン名は公式サイトの変更により異なる場合があります。',
  faqH:'よくある質問',faq:[['「API Key が正しくありません」と出る','公式サイトからもう一度コピーしてください。前後の空白に注意し、Key が削除されていないか確認します。'],['「残高不足／上限に達した」と出る','サービスの公式サイトでチャージするか、無料枠を確認してください。ツールボックス自体は無料です。'],['「AI サービスに接続できません」と出る','ネットワークを確認してください。社内ネットワークで外部サイトが制限されている場合は管理者に相談してください。'],['「モデルが見つかりません」と出る','「モデル一覧を取得」をクリックし、一覧から選んでください。'],['費用はどのくらい？','PPT の構成案 1 回で使うのは数千トークン程度で、多くの場合ごく少額です。料金は各公式サイトでご確認ください。']],
  back:'‹ AI 設定に戻る'}
};
var ORDER={zh:['deepseek','qwen','zhipu','kimi','doubao','claude','openai','gemini','ollama'],ja:['claude','openai','gemini','deepseek','qwen','zhipu','kimi','doubao','ollama']};
function el(tag,cls,text){var e=document.createElement(tag);if(cls)e.className=cls;if(text!=null)e.textContent=text;return e;}
function lang(){var i=window.OfficeI18n;return i&&i.lang&&i.lang()==='ja'?'ja':'zh';}
function render(){var l=lang(),t=T[l],g=document.getElementById('guide');g.textContent='';document.documentElement.lang=l==='ja'?'ja':'zh-CN';document.title=l==='ja'?'API Key の取得手順 · Office Toolbox':'申请 AI API Key 教程 · 办公工具箱';
 document.getElementById('back').textContent=t.back;
 document.querySelectorAll('.lang button').forEach(function(b){b.setAttribute('aria-pressed',String(b.dataset.lang===l));});
 g.append(el('h1','',t.title),el('p','lead',t.lead),el('h2','',t.pickH));var pk=el('div','pick');t.pick.forEach(function(x){var d=el('div');d.append(el('b','',x[0]),el('span','',x[1]));pk.append(d);});g.append(pk);
 g.append(el('h2','',t.allH));var ol=el('ol','steps');t.all.forEach(function(x){ol.append(el('li','',x));});g.append(ol,el('p','note',t.note));
 g.append(el('h2','',t.provH));
 ORDER[l].forEach(function(id){var p=P[id],box=el('section','provider');box.id=id;var top=el('div','top'),a=el('a','go',t.go);a.href=p.url;a.target='_blank';a.rel='noopener';top.append(el('h3','',p.name),a);
  var o=el('ol');p[l].forEach(function(x){o.append(el('li','',x));});var meta=el('p','meta',t.model);var c=el('code','',p.model||t.modelNone);meta.append(c);var u=el('p','meta',p.url);box.append(top,o,meta,u);g.append(box);});
 g.append(el('h2','',t.faqH));var dl=el('dl','faq');t.faq.forEach(function(x){dl.append(el('dt','',x[0]),el('dd','',x[1]));});g.append(dl);
 if(location.hash){var target=document.getElementById(location.hash.slice(1));if(target)target.scrollIntoView();}}
document.querySelectorAll('.lang button').forEach(function(b){b.onclick=function(){if(window.OfficeI18n)window.OfficeI18n.set(b.dataset.lang);render();};});
if(window.OfficeI18n&&window.OfficeI18n.onChange)window.OfficeI18n.onChange(render);
if(window.OfficeI18n&&window.OfficeI18n.start)window.OfficeI18n.start();
render();
})();
