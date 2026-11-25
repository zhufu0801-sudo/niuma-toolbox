#!/bin/bash
# 办公工具箱 Mac 版卸载：从 WPS 和 Microsoft Office 里移除、停止本机服务。设置、AI 配置和备份文件保留，需要时可手动删除。
cd "$(dirname "$0")"
APP="$HOME/Library/Application Support/OfficeToolbox/app"
case "${LANG:-}${TOOLBOX_LANG:-}" in ja*) JA=1;; *) JA=0;; esac
NODE=""
for n in "$APP/runtime/bin/node" /opt/homebrew/bin/node /usr/local/bin/node "$(command -v node 2>/dev/null || true)"; do if [ -n "$n" ] && [ -x "$n" ] && [ "$("$n" -p 'process.versions.node.split(".")[0]' 2>/dev/null || echo 0)" -ge 22 ] 2>/dev/null; then NODE="$n"; break; fi; done
if [ -z "$NODE" ] || [ ! -f "$APP/lib/mac-install.mjs" ]; then
  launchctl bootout "gui/$(id -u)" "$HOME/Library/LaunchAgents/com.officetoolbox.server.plist" 2>/dev/null; rm -f "$HOME/Library/LaunchAgents/com.officetoolbox.server.plist"
  if [ "$JA" = 1 ]; then echo "ローカルサービスを停止しました。"; else echo "已停止本机服务。"; fi; exit 0
fi
"$NODE" "$APP/lib/mac-install.mjs" uninstall --app "$APP"
rm -f "$HOME/Desktop/办公工具箱.command"
if [ "$JA" = 1 ]; then echo "プログラム本体を削除するには、このフォルダーをゴミ箱に入れてください：$HOME/Library/Application Support/OfficeToolbox"; else echo "如果要连程序一起删除，请把这个文件夹移到废纸篓：$HOME/Library/Application Support/OfficeToolbox"; fi
