# 把 skill 推到 GitHub 并在 ClawHub 发布

说明：**ClawHub 不会从 GitHub 自动拉取**，而是你从**本地上传** skill 包。因此流程是：**先把代码推到 GitHub 做源码托管 → 再在本地用 ClawHub CLI 发布**。下面按顺序做即可。

---

## 一、把代码传到 GitHub

### 1. 在 GitHub 上建仓库

1. 打开 [GitHub](https://github.com) 并登录。
2. 右上角 **"+"** → **New repository**。
3. 填写：
   - **Repository name**：例如 `automate-excel` 或 `my-excel-skill`。
   - **Public**，**不要**勾选 “Add a README”（本地已有文件）。
4. 点 **Create repository**，记下仓库地址，例如：  
   `https://github.com/你的用户名/automate-excel.git`

### 2. 在本地把本 skill 做成 Git 仓库并推送

在**包含 `automate-excel` 文件夹的那一层**打开终端（例如你的目录是 `skills`，里面有 `automate-excel`，就在 `skills` 的上一级，或直接把 `automate-excel` 当作仓库根）。

**方式 A：整个 automate-excel 作为仓库根（推荐）**

```bash
cd /Users/z/skills/automate-excel
git init
git add .
git commit -m "Initial commit: Excel automation skill"
git branch -M main
git remote add origin https://github.com/你的用户名/automate-excel.git
git push -u origin main
```

**方式 B：上层目录是仓库，automate-excel 是子目录**

```bash
cd /Users/z/skills
git init
git add automate-excel
git commit -m "Add automate-excel skill"
git branch -M main
git remote add origin https://github.com/你的用户名/你的仓库名.git
git push -u origin main
```

- 若已有 `.git`（例如整个 `skills` 已是仓库），只需在对应目录里 `git add` / `git commit`，再 `git remote add origin ...` 和 `git push`。
- 若提示没有权限，改用 SSH 地址：  
  `git remote add origin git@github.com:你的用户名/automate-excel.git`  
  并确保本机已配置 SSH key。

### 3. 建议加的 Git 忽略

在仓库根建 `.gitignore`，例如：

```
# Python
__pycache__/
*.py[cod]
.venv/
venv/
*.egg-info/

# OS
.DS_Store
```

然后：

```bash
git add .gitignore
git commit -m "Add .gitignore"
git push
```

---

## 二、从本地上传到 ClawHub（“通过 GitHub 导入”的实质步骤）

ClawHub 是**从你本机的文件夹上传**，不会去 GitHub 拉代码。所以你要在**已经包含该 skill 的本地目录**执行发布。

### 1. 安装并登录 ClawHub CLI

```bash
npm i -g clawhub
# 或: pnpm add -g clawhub
clawhub login
```

按提示在浏览器里登录（或使用 `clawhub login --token <token>`）。

### 2. 发布本 skill

**若仓库根就是 automate-excel 目录**（方式 A）：

```bash
cd /Users/z/skills/automate-excel
clawhub publish . --slug automate-excel --name "Automate Excel" --version 1.0.0 --tags latest --changelog "Initial release from GitHub repo"
```

**若 automate-excel 是子目录**（方式 B）：

```bash
cd /Users/z/skills
clawhub publish ./automate-excel --slug automate-excel --name "Automate Excel" --version 1.0.0 --tags latest --changelog "Initial release from GitHub repo"
```

- `--slug`：在 ClawHub 里的唯一标识，建议和仓库名一致。
- `--name`：页面上显示的名称。
- `--version`：语义化版本，例如 1.0.0。
- `--tags latest`：打上 `latest` 标签，方便别人 `clawhub install automate-excel`。

### 3. 确认已发布

- 打开 [clawhub.ai](https://clawhub.ai)，搜索 `automate-excel` 或你的 skill 名称。
- 或本地执行：`clawhub search "excel"` 看是否出现你的 skill。

---

## 三、以后更新：改代码 → 推 GitHub → 再发布到 ClawHub

1. **改 skill 代码**（在本地）。
2. **推送到 GitHub**：
   ```bash
   git add .
   git commit -m "Add format_columns_as_text / update docs"
   git push
   ```
3. **在 ClawHub 发布新版本**（仍在本地执行）：
   ```bash
   # 在 automate-excel 所在目录
   clawhub publish . --slug automate-excel --name "Automate Excel" --version 1.0.1 --tags latest --changelog "列格式为文本、文档更新"
   ```
   - 每次发布请把 `--version` 递增（如 1.0.0 → 1.0.1）。
   - 已安装的用户可用 `clawhub update automate-excel` 拉取新版本。

---

## 四、别人如何用你的 skill

- **从 ClawHub 安装**（你发布后）：
  ```bash
  clawhub install automate-excel
  ```
- **从 GitHub 安装**（未发布或想用最新源码）：
  ```bash
  git clone https://github.com/你的用户名/automate-excel.git
  # 把 clone 下来的文件夹里的内容放到工作区的 skills/ 下，或按 OpenClaw/Cursor 的 skill 加载方式配置
  ```

---

## 五、小结

| 步骤 | 做什么 |
|------|--------|
| 1 | GitHub 建仓 → 本地 `git init` / `add` / `commit` / `remote` / `push` |
| 2 | 本地安装并登录 `clawhub` |
| 3 | 在**本地**执行 `clawhub publish <skill 目录> --slug ... --version ...` |
| 4 | 更新时：改代码 → `git push` → 再执行一次 `clawhub publish` 并提高版本号 |

**ClawHub 不会通过 GitHub 自动导入**，需要你在本机用 CLI 上传；GitHub 只负责保存源码和协作，发布到 ClawHub 始终是“本地发布”。
