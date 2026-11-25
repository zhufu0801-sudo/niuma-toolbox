# 隐私保护插件接口（接口版本 1）

办公工具箱发给 AI 的每一次请求，都会先经过“隐私保护关卡”（`lib/privacy.mjs`）：

```
页面 → 本机服务 → 内置保护（自动隐藏个人信息）→ 插件 beforeSend（按文件名顺序）→ AI 服务商
页面 ← 本机服务 ← 内置还原 ← 插件 afterReceive（倒序）            ← AI 服务商
```

## 放在哪里

- 把插件文件（`.mjs`）放进本文件夹 `app\plugins\`。
- 以 `_` 开头的文件不会加载（例如本文件夹里的 `_example-privacy.mjs`，改名去掉 `_` 才会生效）。
- 第一次使用 AI 时加载；修改插件后，重启办公工具箱，或调用 `POST /api/ai/plugins/reload`。
- 设置页“AI 服务”下会显示已加载的插件名和版本。

## 插件格式

```js
export default {
  name: '公司隐私保护',        // 必填，显示在设置页
  version: '1.0.0',            // 可选
  apiVersion: 1,               // 可选；填写时必须等于工具箱的接口版本
  async beforeSend(ctx) { ... },   // 可选，发送前
  async afterReceive(ctx) { ... }  // 可选，收到回答后
};
```

### beforeSend(ctx)

`ctx` 内容：

| 字段 | 说明 |
|---|---|
| `messages` | 即将发送的消息数组 `[{role, content}]`（已是副本，可随意修改） |
| `feature` | 功能：`ppt-outline`（写大纲）、`translate`（翻译润色）、`formula`（帮写公式）、`test`（测试连接） |
| `provider` / `model` | 服务商编号和模型名 |
| `region` | `cn` 国内、`global` 海外、`local` 本机（Ollama）、`any` 自定义 |
| `mode` | 内置保护方式：`mask` / `remind` / `off` |
| `state` | 本次请求专用的空对象，可存放还原用的对应表，`afterReceive` 时原样拿回 |

返回值：

- `undefined`：不修改，继续发送。
- `{ messages }`：用新的消息数组发送（必须是 1～50 条 `{role: 'system'|'user'|'assistant', content: 字符串}`）。
- `{ block: true, reason: '原因' }`：拦截，不发送；`reason` 会显示给用户。

### afterReceive(ctx)

`ctx` 和上面相同，但没有 `messages`，多了 `answer`（AI 的回答文字）。返回 `undefined`、`{ answer }` 或 `{ block: true, reason }`。

## 安全规则（工具箱强制执行）

- 插件抛出错误、超过 5 秒没有返回、返回格式不对：**这次请求不发送**，并告诉用户是哪个插件。
- 插件文件加载失败：**AI 功能整体暂停**，避免用户以为有保护、其实没有。
- 内置保护先于插件运行；插件看到的是已经隐藏过的内容。若需要插件自己处理原文，请在设置里把内置保护设为“关闭”。
- 对应表只存在内存里，不写文件、不写日志。插件也应遵守这一点。

## 示例

见 `_example-privacy.mjs`：拦截含“机密”字样的内容，并把公司内部项目代号替换成占位符、收到回答后换回。
