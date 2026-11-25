办公工具箱 0.16.0 · macOS 版

安装
1. 解压后，在“访达”里右键点“install.command”→“打开”（第一次需要这样打开，因为它不是从 App Store 下载的）。
2. 按提示选择装到 WPS 还是 Microsoft Office。没有 Node.js 时会询问是否从官网 nodejs.org 下载（免费）。
3. 装 Microsoft Office 版时，macOS 会要求输入登录密码，用来信任只给本机 localhost 用的证书。
4. 完全退出并重新打开 WPS 或 Word / Excel / PowerPoint：
   - WPS：顶部出现“办公工具箱”标签。
   - Microsoft Office：在“插入 → 加载项 → 我的加载项”里打开“办公工具箱”。

说明
- 图片转文字使用 macOS 自带的文字识别：中文需要 macOS 11 或更新版本，日文需要 macOS 13 或更新版本。
- 压缩图片使用 macOS 自带的 sips。
- Mac 版 WPS 对加载项的支持随版本不同；没有出现标签时，请把 WPS 更新到最新版。
- 读取 PDF 里的图片需要的图像组件只有 Windows 版带，Mac 上“提取 PDF 里的图片”不可用；PDF 转 Word 的文字、扫描件识别可以用。

卸载
右键点“uninstall.command”→“打开”。设置、AI 配置和备份文件会保留。

本次没有 Mac 真机，未实测。PDF 图片型加工依赖图像组件，缺少组件时功能不可用并给出提示。重要文件先备份；AI 为可选，启用后内容发送给你选择的服务商。
