# GitHub 自动测试的工作流

这里是 GitHub Actions 工作流的原件，内容和 `.github/workflows/` 里的完全一样（自动测试会核对两边是否一致）。

上传到 GitHub 前，把这个文件夹里的两个 `.yml` 复制到仓库的 `.github/workflows/` 文件夹：
- `test.yml`：每次推送时在 Windows、Mac、Linux 上运行自动测试，并在 Mac 上模拟一次安装、检查、卸载。
- `release-mac.yml`：在 GitHub 的 Mac 上打 Mac 安装包（中文版、日文版）。

改工作流时，两边一起改。
