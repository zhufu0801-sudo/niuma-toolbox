# 说明书维护

用户阅读原稿在 `docs/manual/{zh,ja,en}.md`。从项目根目录运行：

```text
node scripts/docs/build-manual.mjs
```

生成 `release/docs` 下的三语 PDF、DOCX 和构建报告。中文、日文非表格行的引号和全角括号必须成对；构建时会检查。

本轮说明书的编写脚本及输入快照保存在 `scripts/docs/authoring`。需要从这些输入重新生成 Markdown 时运行：

```text
node scripts/docs/authoring/write-manuals.cjs
```

这会覆盖三语 Markdown，先保存其中的手工修改。`inventory.json` 是编写时的功能和翻译快照，不会自动随产品代码更新；功能改动需同时维护说明书和编写输入。步骤解析仅在引号外的 ` → ` 处分隔，不会拆开引号内的菜单路径。

打包先在临时目录生成说明书。仅在缺少本机 CJK 字体（生成器退出码 78）时，优先使用 `release/docs`，其次使用仓库 `docs/manual/prebuilt` 中已有的中文、日文 PDF，并输出警告；其他生成错误仍会中止。没有可用的现成 PDF 也会明确报错。修改手册后，请把生成的中日 PDF 更新到 `docs/manual/prebuilt`，用于不上传 `release/` 的 GitHub 构建。可通过 `MANUAL_FONT` 和 `MANUAL_FONT_NAME` 指定已安装字体。

Windows 的成功构建与模拟缺字体测试不代表 macOS 已验证。发布到 GitHub 后还需在 macOS runner 上实际执行 Mac 打包；本次未执行该步骤。
