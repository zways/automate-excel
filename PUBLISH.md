# 发布到 ClawHub

本 skill 已按 ClawHub 规范准备好，可按以下步骤上传。

**若你想先把代码推到 GitHub，再发布到 ClawHub**：请直接看 **[GITHUB_AND_CLAWHUB.md](GITHUB_AND_CLAWHUB.md)**，里面有「推到 GitHub → 本地用 CLI 发布到 ClawHub」的完整步骤。

---

## 前置条件

- 已注册 [ClawHub](https://clawhub.ai) 开发者账号（GitHub 账号需至少一周）
- 已安装 Node.js

## 1. 安装 CLI

```bash
npm i -g clawhub
# 或
pnpm add -g clawhub
```

若全局安装有权限问题，可用 npx：

```bash
npx clawhub --help
```

## 2. 登录

```bash
clawhub login
```

按提示在浏览器中完成登录；或使用 token：

```bash
clawhub login --token <你的 API Token>
```

## 3. 发布本 skill

在**本仓库根目录**（即 `skills` 的上一级，或包含 `automate-excel` 的目录）执行：

```bash
clawhub publish ./automate-excel --slug automate-excel --name "Automate Excel" --version 1.0.0 --tags latest
```

若在 `automate-excel` 目录内：

```bash
clawhub publish . --slug automate-excel --name "Automate Excel" --version 1.0.0 --tags latest
```

可选：附带 changelog

```bash
clawhub publish ./automate-excel --slug automate-excel --name "Automate Excel" --version 1.0.0 --tags latest --changelog "Initial release: Excel read/write/merge/validate and scripts"
```

## 4. 后续更新

修改 skill 后递增版本号（如 1.0.1），再执行：

```bash
clawhub publish ./automate-excel --slug automate-excel --name "Automate Excel" --version 1.0.1 --tags latest --changelog "描述本次更新"
```

或使用同步（会扫描本地 skill 并发布有变动的）：

```bash
clawhub sync --dry-run   # 先预览
clawhub sync --all       # 确认后执行
```

## 参考

- [ClawHub 文档](https://docs.openclaw.ai/tools/clawdhub)
- [发布指南](https://www.openclawexperts.io/guides/custom-dev/how-to-publish-a-skill-to-clawhub)
