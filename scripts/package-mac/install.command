#!/bin/bash
# 办公工具箱 Mac 版安装：双击运行。会把程序复制到“~/Library/Application Support/OfficeToolbox”，
# 登记到 WPS 和 / 或 Microsoft Office，并设置登录后自动启动本机服务（http://127.0.0.1:39127）。
set -e
cd "$(dirname "$0")"
HERE="$(pwd)"
case "${LANG:-}${TOOLBOX_LANG:-}" in ja*) JA=1;; *) JA=0;; esac
say(){ if [ "$JA" = 1 ]; then echo "$2"; else echo "$1"; fi; }
DEST="$HOME/Library/Application Support/OfficeToolbox"
APP="$DEST/app"
NODE_VERSION="v22.11.0"

find_node(){
  for n in "$APP/runtime/bin/node" "$HERE/app/runtime/bin/node" /opt/homebrew/bin/node /usr/local/bin/node "$(command -v node 2>/dev/null || true)"; do
    if [ -n "$n" ] && [ -x "$n" ]; then
      major="$("$n" -p 'process.versions.node.split(".")[0]' 2>/dev/null || echo 0)"
      # 本机服务要 Node.js 22 以上（package.json 的 engines）；旧版本跳过，没有合适的就下载 22。
      if [ "$major" -ge 22 ] 2>/dev/null; then echo "$n"; return 0; fi
    fi
  done
  return 1
}

say "== 办公工具箱 Mac 版安装 ==" "== オフィスツールボックス Mac 版インストール =="
mkdir -p "$DEST"
# 复制程序；保留已有的设置、AI 配置、证书（generated 文件夹）和本机 Node（runtime）。
mkdir -p "$APP"
rsync -a --delete --exclude generated --exclude runtime "$HERE/app/" "$APP/"
if [ -d "$HERE/app/runtime" ] && [ ! -d "$APP/runtime" ]; then cp -R "$HERE/app/runtime" "$APP/runtime"; fi

NODE="$(find_node || true)"
if [ -z "$NODE" ]; then
  ARCH="$(uname -m)"; [ "$ARCH" = "x86_64" ] && ARCH=x64
  FILE="node-$NODE_VERSION-darwin-$ARCH.tar.gz"
  say "这台 Mac 上没有 Node.js（运行本机服务需要）。" "この Mac には Node.js（ローカルサービスに必要）がありません。"
  say "是否从官方网站 nodejs.org 下载 $FILE（约 45 MB，免费）？[Y/n]" "公式サイト nodejs.org から $FILE（約 45 MB、無料）をダウンロードしますか？[Y/n]"
  read -r ans
  case "$ans" in n|N|no|NO) say "已取消。也可以自己从 https://nodejs.org 安装 LTS 版后，再双击本文件。" "キャンセルしました。https://nodejs.org から LTS 版を入れてから、もう一度このファイルをダブルクリックしてもかまいません。"; exit 1;; esac
  TMP="$(mktemp -d)"
  curl -fL --progress-bar -o "$TMP/$FILE" "https://nodejs.org/dist/$NODE_VERSION/$FILE"
  curl -fsSL -o "$TMP/SHASUMS256.txt" "https://nodejs.org/dist/$NODE_VERSION/SHASUMS256.txt"
  WANT="$(grep " $FILE\$" "$TMP/SHASUMS256.txt" | awk '{print $1}')"
  GOT="$(shasum -a 256 "$TMP/$FILE" | awk '{print $1}')"
  if [ -z "$WANT" ] || [ "$WANT" != "$GOT" ]; then say "下载的文件校验不通过，已停止安装。" "ダウンロードしたファイルの検証に失敗したため、インストールを止めました。"; rm -rf "$TMP"; exit 1; fi
  mkdir -p "$APP/runtime" && tar -xzf "$TMP/$FILE" -C "$APP/runtime" --strip-components 1
  rm -rf "$TMP"
  NODE="$APP/runtime/bin/node"
fi
say "使用 Node.js：$NODE" "使用する Node.js：$NODE"

DETECT="$("$NODE" "$APP/lib/mac-install.mjs" detect)"
HAS_WPS=0; HAS_OFFICE=0
echo "$DETECT" | grep -q '"wps":"' && HAS_WPS=1
echo "$DETECT" | grep -q '"office":\["' && HAS_OFFICE=1
TARGETS="${TOOLBOX_TARGETS:-}"
if [ -z "$TARGETS" ]; then
  if [ $HAS_WPS = 1 ] && [ $HAS_OFFICE = 1 ]; then
    say "找到了 WPS Office 和 Microsoft Office。装到哪里？ 1) 两个都装  2) 只装 WPS  3) 只装 Microsoft Office  [1]" "WPS Office と Microsoft Office が見つかりました。どちらに入れますか？ 1) 両方  2) WPS のみ  3) Microsoft Office のみ  [1]"
    read -r c; case "$c" in 2) TARGETS=wps;; 3) TARGETS=office;; *) TARGETS=wps,office;; esac
  elif [ $HAS_WPS = 1 ]; then TARGETS=wps
  elif [ $HAS_OFFICE = 1 ]; then TARGETS=office
  else say "没有找到 WPS Office 或 Microsoft Office，先只装本机服务；装好 Office 后再运行一次本文件。" "WPS Office も Microsoft Office も見つかりません。ローカルサービスだけ入れます。Office を入れた後にもう一度実行してください。"; TARGETS=none; fi
fi
"$NODE" "$APP/lib/mac-install.mjs" install --app "$APP" --targets "$TARGETS" --node "$NODE"

# 桌面上放一个“办公工具箱.command”：确认服务在运行并打开设置页。
LAUNCH="$HOME/Desktop/办公工具箱.command"
cat > "$LAUNCH" <<LAUNCHER
#!/bin/bash
launchctl kickstart "gui/\$(id -u)/com.officetoolbox.server" >/dev/null 2>&1 || true
sleep 1
open "http://127.0.0.1:39127/settings.html"
LAUNCHER
chmod +x "$LAUNCH"
sleep 2
if curl -fsS --max-time 5 http://127.0.0.1:39127/api/health >/dev/null 2>&1; then
  say "完成。本机服务已启动。请完全退出并重新打开 WPS 或 Word / Excel / PowerPoint。" "完了しました。ローカルサービスが起動しています。WPS または Word / Excel / PowerPoint を完全に終了して開き直してください。"
else
  say "已安装，但本机服务还没响应，请稍等几秒后双击桌面上的“办公工具箱.command”。日志：$APP/generated/server.log" "インストールしましたが、ローカルサービスがまだ応答しません。数秒後にデスクトップの「办公工具箱.command」をダブルクリックしてください。ログ：$APP/generated/server.log"
fi
say "（WPS Mac 版对加载项的支持随版本不同；顶部没有出现“办公工具箱”时，请把 WPS 更新到最新版。）" "（WPS の Mac 版はバージョンによってアドイン対応が異なります。上部に表示されないときは WPS を最新版に更新してください。）"
