#!/bin/bash
# GitHub 的 macOS 机器上：打 Mac 安装包，再像用户双击“install.command”那样装一次（不需要人操作）。
# 之后由 scripts/ci/mac-check.mjs 检查本机服务、WPS 登记、钥匙串、中文和带空格的路径、Office 侧栏、卸载。
set -euo pipefail
cd "$(dirname "$0")/../.."
ROOT="$(pwd)"
# 1. 打包：底包用项目自己（npm ci 之后 node_modules 里是 Mac 的组件）；Mac 包不带 Node，runtime 放空文件夹。
mkdir -p runtime
node scripts/package/build-release.mjs --base . --out release --only mac
PKG="$(ls -d release/OfficeToolbox-*-macOS | head -1)"
echo "安装包：$PKG"
test -x "$PKG/install.command" || chmod +x "$PKG/install.command"
# 2. 解压后的 zip 也要能用：检查 .command 的可执行权限保留了
ZIP="$PKG.zip"; rm -rf /tmp/unzip-check && mkdir -p /tmp/unzip-check && (cd /tmp/unzip-check && unzip -q "$ROOT/$ZIP")
test -x /tmp/unzip-check/"$(basename "$PKG")"/install.command || { echo "zip 里的 install.command 没有可执行权限"; exit 1; }
# 3. 安装：装到 WPS（这台机器上没有 WPS，只写登记文件）；Office 的证书信任需要输密码，留给 mac-check.mjs 单独测。
#    Node 用 GitHub 装好的（真实用户没有 Node 时安装脚本会从 nodejs.org 下载）。
TOOLBOX_TARGETS=wps bash /tmp/unzip-check/"$(basename "$PKG")"/install.command </dev/null
