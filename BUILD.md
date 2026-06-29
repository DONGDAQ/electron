# 打包构建流程文档

## 概述

报价工具使用 **Electron + Flask** 架构，运行时 Electron 负责桌面窗口（UI），Flask 提供 HTTP 后端服务。打包时需要将两者合二为一，打包流程分为三个阶段：

1. **Flask 后端打包** — 用 PyInstaller 将 Python 源码编译成独立的 `flask_server.exe`
2. **文件预置** — 将 flask_server 复制到 `app/` 目录下
3. **Electron 打包** — 用 electron-builder 将整个项目打包成 `报价工具.exe`（Windows 便携版）

最终产出的 `报价工具.exe` 是一个自包含的 Portable exe，双击运行即可，无需安装 Python 或 Node.js。

---

⚠️ 重要注意事项（打包前必读）
------------

### 1. PyInstaller 必须能检测到 pywin32

`app/settlement/_com_utils.py` 中的 `import pythoncom` 和 `import win32com.client` **必须写在文件顶部**，不能写在函数内部。

原因：PyInstaller 只扫描顶层 import，函数内部的 import 会被忽略，导致打包后的 `flask_server.exe` 里没有 pywin32，公式缓存功能（`_ensure_formula_cached`）会静默失败。

验证方法（打包后检查）：
```bash
ls "D:/baojia/electron/dist/flask_build/flask_server/_internal/" | grep -i "pythoncom\|pywin\|win32com"
```

### 2. electron-builder 必须用系统 Node 运行

不要用 `npm run build`（managed Node 路径有问题），直接用系统 Node：
```bash
cd "D:/baojia/electron"
"C:/Program Files/nodejs/npx.cmd" electron-builder --win portable
```

---

## 架构总图

```
┌──────────────────────────────────────────────────────────┐
│                   报价工具.exe (Portable)                  │
│                                                          │
│  ┌──────────┐        ┌──────────────────────────────┐    │
│  │ Electron │        │  resources/app/               │    │
│  │  main.js │ spawn  │                              │    │
│  │          │───────▶│  ┌────────────────────────┐  │    │
│  │ ┌──────┐ │        │  │  flask_server/         │  │    │
│  │ │ win- │ │        │  │  flask_server.exe      │  │    │
│  │ │ dow  │ │        │  │  _internal/            │  │    │
│  │ │      │ │        │  └────────────────────────┘  │    │
│  │ │ 5000 │◀─────── │                              │    │
│  │ └──────┘ │ HTTP   │  ┌────────────────────────┐  │    │
│  └──────────┘        │  │  quote_system/ (源码)   │  │    │
│                       │  │  templates/            │  │    │
│                       │  │  static/               │  │    │
│                       │  │  config/               │  │    │
│                       │  └────────────────────────┘  │    │
│                       │                              │    │
│                       │  ┌────────────────────────┐  │    │
│                       │  │  outputs/              │  │    │
│                       │  │  logs/ /uploads/       │  │    │
│                       │  └────────────────────────┘  │    │
│                       └──────────────────────────────┘    │
└──────────────────────────────────────────────────────────┘
```

**运行原理：**
1. Electron 启动，执行 `main.js`
2. `main.js` 调用 `startFlaskServer()`，spawn 子进程运行 `flask_server.exe`
3. `flask_server.exe` 启动 Flask Web 服务监听 `127.0.0.1:5000`
4. Electron 监听 localhost:5000，就绪后加载页面
5. 用户操作通过 HTTP 请求与 Flask 后端交互

---

## 涉及的工具

| 工具 | 用途 | 版本 | 安装方式 |
|------|------|------|----------|
| **Node.js** | 运行 electron-builder 和构建脚本 | ≥18 | [nodejs.org](https://nodejs.org) |
| **npm** | 包管理 | 随 Node.js | 自动 |
| **electron-builder** | Electron 桌面应用打包 | 24.13.3 | `npm install` |
| **Electron** | 桌面应用框架 | 28.3.3 | `npm install` |
| **Python 3.13.12** | 运行 Flask 源码的 Python 环境 | 3.13.12 | workbuddy 管理 |
| **PyInstaller** | 将 Python 源码编译为 exe | 6.21.0 | `pip install pyinstaller` |
| **pip** (清华镜像) | Python 包管理 | 最新 | 配置 index-url |

### 关键依赖（Python，通过 pip 安装）

所有安装使用清华镜像：
```bash
pip install <包名> -i https://pypi.tuna.tsinghua.edu.cn/simple
```

| 包名 | 用途 |
|------|------|
| flask | Web 框架 |
| openpyxl | Excel 读写 |
| python-docx | Word 文档生成 |
| lxml | HTML/XML 解析 |
| lark-oapi | 飞书 API 客户端 |
| markupsafe | 模板安全转义 |
| requests, httpx | HTTP 客户端 |

---

## 打包流程详解

### 整体构建命令

```bash
npm run build
```

> ⚠️ **注意**：如果运行 `npm run build` 报 `MODULE_NOT_FOUND`，说明 managed Node 路径有问题。
> 请改用系统 Node 直接运行：
> ```bash
> cd "D:/baojia/electron"
> "C:/Program Files/nodejs/npx.cmd" electron-builder --win portable
> ```

这条命令在 `package.json` 中定义为：

```json
"build": "npm run build-flask && node scripts/prebuild.js && electron-builder --win portable && node scripts/postbuild.js"
```

即四个步骤顺序执行：

```
Step 1: npm run build-flask      ├── PyInstaller 编译 Python → flask_server.exe
Step 2: node scripts/prebuild.js ├── 复制 flask_server → app/flask_server/
Step 3: electron-builder         ├── 打包 Electron → dist/报价工具.exe
Step 4: node scripts/postbuild.js└── 后处理：删除 asar、同步文件
```

---

### Step 1: 编译 Flask 后端 → flask_server.exe

**入口脚本：`app/run_flask.py`**

```python
import sys, os
from pathlib import Path

if getattr(sys, 'frozen', False):
    base = Path(sys._MEIPASS)  # PyInstaller 打包后，解压目录
else:
    base = Path(__file__).parent

sys.path.insert(0, str(base))
os.chdir(str(base))

from quote_system.web_app import app  # Flask 应用入口

# 启动时后台刷新缓存
import threading
def _startup_refresh():
    try:
        from quote_system.auto_fill_tk import sync_tk_data
        sync_tk_data()
    except Exception:
        pass
    try:
        from quote_system.report_cache import sync_report_cache
        sync_report_cache()
    except Exception:
        pass
threading.Thread(target=_startup_refresh, daemon=True).start()

app.run(host="127.0.0.1", port=5000, debug=False)
```

`run_flask.py` 是整个 Flask 服务的入口，`sys.frozen` 判断是 PyInstaller 打包后运行还是源码直接运行。

**配置文件中定义编译参数：`app/flask_server.spec`**

```python
# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['run_flask.py'],                                         # 入口脚本
    pathex=[],
    binaries=[],
    datas=[                                                     # 需要打包的资源目录
        ('quote_system/templates', 'quote_system/templates'),
        ('quote_system/static', 'quote_system/static'),
        ('config', 'config'),
        ('模板', '模板'),
    ],
    hiddenimports=[                                             # 显式声明的隐藏导入
        'quote_system.web_app',
        'quote_system.generator',
        'quote_system.memoq_html',
        'quote_system.projects',
        'quote_system.config',
        'quote_system.save_path_config',
        'quote_system.paths',
        'quote_system.utils',
        'quote_system.feishu_client',
        'quote_system.auto_quote',
        'quote_system.auto_quote_zhan_shuang',
        'quote_system.auto_quote_zhan_shuang_feishu',
        'quote_system.auto_quote_4399',
        'quote_system.auto_fill_tk',
        'settlement.generate_settlement',
        'settlement.generate_settlement_bilibili',
        'settlement.generate_settlement_zhan_shuang',
        'settlement.generate_settlement_zhan_shuang_faxing',
        'settlement.generate_settlement_4399',
        'settlement.generate_settlement_diezhi',
        'settlement.settlement_tracker',
        'lxml', 'openpyxl', 'docx',
    ],
    excludes=['matplotlib', 'PyQt5', 'IPython', 'pytest'],      # 排除不必要的包
)
exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name='flask_server',
    console=True,                                               # 保留控制台窗口便于调试
    upx=True,                                                   # UPX 压缩
)
coll = COLLECT(
    exe, a.binaries, a.datas,
    name='flask_server',                                        # onedir 模式输出目录
)
```

**执行命令：**

```bash
cd app
C:/Users/admin/.workbuddy/binaries/python/versions/3.13.12/python.exe \
  -m PyInstaller \
  --distpath ../dist/flask_build \
  --workpath ../dist/flask_work \
  -y flask_server.spec
```

| 参数 | 含义 |
|------|------|
| `--distpath` | 输出目录，生产物 `dist/flask_build/flask_server/` |
| `--workpath` | 工作目录，存放临时文件 |
| `-y` | 覆盖已存在的输出目录 |
| `flask_server.spec` | 编译配置文件 |

**PyInstaller 内部工作机制：**

```
run_flask.py
    │
    ▼
Analysis (模块依赖分析)
    │  ├── 扫描 import 语句
    │  ├── 追踪所有依赖模块
    │  ├── 合并 datas 资源文件
    │  └── 处理 hiddenimports / excludes
    ▼
PYZ (压缩 Python 字节码)
    │  将所有 .py 编译为 .pyc 并打包成 ZIP
    ▼
PKG (构建归档)
    │  将 PYZ、二进制文件、资源文件合并为一个归档
    ▼
EXE (生成 exe 引导程序)
    │  生成 flask_server.exe（含 bootloader + 归档）
    ▼
COLLECT (onedir 模式)
    │  将 exe + _internal/ 放入 flask_server/ 目录
    ▼
dist/flask_build/flask_server/
    ├── flask_server.exe        ← 入口（约 8MB）
    └── _internal/              ← 运行时依赖（约 70MB）
        ├── python313.dll       ← Python 解释器
        ├── quote_system/       ← 业务代码
        ├── flask/              ← Flask 框架
        ├── openpyxl/           ← Excel 库
        ├── lxml/               ← XML 库
        ├── site-packages/      ← 第三方包
        └── ...
```

**PYInstaller 两种输出模式：**

| 模式 | 参数 | 输出 | 优点 | 缺点 |
|------|------|------|------|------|
| **onedir**（当前使用） | `COLLECT` | 目录 `flask_server/`，内含 exe + `_internal/` | 启动快，调试方便 | 目录略大 |
| onefile | 无 COLLECT | 单个 exe `flask_server.exe` | 只有一个文件 | 启动慢（需临时解压） |

---

### Step 2: 预置 flask_server → app/flask_server/

**脚本：`scripts/prebuild.js`**

```javascript
const srcDir = path.join(__dirname, '..', 'dist', 'flask_build', 'flask_server');
const destDir = path.join(__dirname, '..', 'app', 'flask_server');

// 删除旧的
if (fs.existsSync(destDir)) fs.rmSync(destDir, { recursive: true });

// 复制整个 flask_server 目录到 app/flask_server/
fs.mkdirSync(destDir, { recursive: true });
const entries = fs.readdirSync(srcDir, { withFileTypes: true });
for (const entry of entries) {
  const srcPath = path.join(srcDir, entry.name);
  const destPath = path.join(destDir, entry.name);
  if (entry.isDirectory()) {
    fs.cpSync(srcPath, destPath, { recursive: true });
  } else {
    fs.copyFileSync(srcPath, destPath);
  }
}
```

**为什么需要这一步？**

electron-builder 的 `extraResources` 只会打包已在项目目录中的文件。如果不先把 `flask_server` 复制到 `app/` 下，electron-builder 打包时它就不在 extraResources 的 `from: "app"` 扫描范围里，打包出来的 exe 里就没有 Flask 后端，启动就会报 "Flask 服务器意外断开"。

---

### Step 3: 打包 Electron 桌面应用

**配置：`package.json` 中的 `build` 字段**

```json
"build": {
  "appId": "com.example.translation-quote",
  "productName": "报价工具",
  "icon": "icon.ico",
  "directories": { "output": "dist" },
  "extraResources": [
    {
      "from": "app",           ← 扫描目录
      "to": "app",             ← 打包后路径 (resources/app/)
      "filter": [
        "**/*",                ← 包含所有文件
        "!outputs/**",         ← 排除：运行时输出
        "!报价/**",            ← 排除：报价单历史
        "!结算/**",            ← 排除：结算文件
        "!汇总/**",            ← 排除：汇总文件
        "!**/*.log",           ← 排除：日志
        "!**/__pycache__/**"   ← 排除：Python 缓存
      ]
    }
  ],
  "win": {
    "target": [
      { "target": "nsis", "arch": ["x64"] },
      { "target": "portable", "arch": ["x64"] }
    ],
    "requestedExecutionLevel": "asInvoker"
  },
  "nsis": {
    "oneClick": false,
    "perMachine": false,
    "allowToChangeInstallationDirectory": false,
    "artifactName": "${productName}-${version}-setup.exe"
  },
  "portable": {
    "artifactName": "${productName}.exe",
    "requestExecutionLevel": "highest"
  }
}
```

**extraResources 详细说明：**

`extraResources` 是 electron-builder 的核心机制。它告诉 electron-builder：**将本地 `app/` 目录下的文件（应用代码、flask_server、模板等）复制到打包后 exe 内 `resources/app/` 目录下**，但这些文件不作为 app.asar 压缩包的一部分，保持为独立目录。

这么做是因为：
- `flask_server.exe` 是二进制文件（约 80MB），在 asar 里无法直接运行
- 需要保留文件系统路径，让 `main.js` 的 `spawn()` 能找到并执行它
- 配置文件、模板文件等也需要保持文件系统可读写性

打包后解压目录结构：
```
报价工具.exe (自解压)
  └── resources/
      └── app/
          ├── flask_server/
          │   ├── flask_server.exe    ← Python 编译后的后端
          │   └── _internal/          ← Python 运行时依赖
          ├── quote_system/
          │   ├── web_app.py          ← Flask 入口
          │   ├── templates/          ← Jinja2 模板
          │   └── static/             ← CSS/JS 静态资源
          ├── config/                 ← 配置文件
          ├── 模板/                   ← Excel/Word 模板
          ├── main.js                 ← Electron 主进程
          └── package.json
```

**主进程启动逻辑：`main.js` 中的 `startFlaskServer()`**

```javascript
function startFlaskServer() {
  const appDir = getAppDir();  // 打包后: resources/app/ ; 开发时: app/

  // 优先 onedir 模式 (flask_server/flask_server.exe)
  let flaskExe = path.join(appDir, 'flask_server', 'flask_server.exe');
  // 其次 onefile 模式 (flask_server.exe)
  if (!fs.existsSync(flaskExe)) {
    flaskExe = path.join(appDir, 'flask_server.exe');
  }

  if (fs.existsSync(flaskExe)) {
    // 运行编译好的 flask_server.exe
    pythonProcess = spawn(flaskExe, [], { cwd: appDir, ... });
  } else {
    // 开发环境：直接运行 Python 源码
    const pythonPath = findPython();
    pythonProcess = spawn(pythonPath, ['-m', 'quote_system.web_app'], {
      cwd: appDir,
      env: { ...process.env, PYTHONPATH: appDir },
    });
  }

  // Flask 意外退出时自动重试（最多 3 次）
  pythonProcess.on('close', (code) => {
    if (flaskRetryCount < 3) {
      flaskRetryCount++;
      setTimeout(startFlaskServer, 2000);
    } else {
      dialog.showMessageBox(mainWindow, {
        type: 'warning',
        title: '服务器断开',
        message: 'Flask 服务器多次断开，应用将退出',
      }).then(() => app.quit());
    }
  });
}
```

**electron-builder 执行过程：**

```
electron-builder --win portable
    │
    ├── 1. 读取 package.json 中的 "build" 配置
    ├── 2. 准备打包环境
    │       ├── 定位 Electron 28.3.3
    │       ├── 处理 extraResources
    │       │   └── 从 app/ 复制文件到 resources/app/
    │       └── 处理 asar 打包
    ├── 3. 生成 win-unpacked/
    │       ├── resources/
    │       │   ├── app.asar          ← 压缩的 Electron 代码
    │       │   └── app/              ← extraResources 解压目录
    │       ├── electron.exe          ← Electron 主程序
    │       └── 其他 DLL/运行时文件
    └── 4. 压缩为 Portable exe
            └── dist/报价工具.exe      ← 最终输出
```

**便携版 vs 安装版：**

`package.json` 配置了两个 target：

```json
"win": {
  "target": [
    { "target": "nsis", "arch": ["x64"] },       ← 安装版（需安装）
    { "target": "portable", "arch": ["x64"] }    ← 便携版（双击即用）
  ]
}
```

| 特性 | Portable（当前使用） | NSIS 安装版 |
|------|-------------------|-------------|
| 产物品名 | `报价工具.exe` | `报价工具-2.0.0-setup.exe` |
| 是否需要安装 | 不需要，双击即可 | 需要运行安装程序 |
| 文件写入 | 运行时会解压到 Temp 目录 | 安装到 Program Files |
| 便携性 | 可随身携带 U 盘 | 只能安装在本机 |

当前构建命令 `npm run build` 只生成 Portable 版。如果有安装版需求，可修改命令为：
```json
"build": "npm run build-flask && node scripts/prebuild.js && electron-builder --win portable --win nsis && node scripts/postbuild.js"
```

---

### Step 4: 后处理

**脚本：`scripts/postbuild.js`**

```javascript
const distApp = path.join(__dirname, '..', 'dist', 'win-unpacked', 'resources', 'app');

// 1. 删除 app.asar（强制 Electron 使用 app/ 目录，避免 flask_server 被 asar 隔离）
//    因为 flask_server.exe 是二进制文件，在 asar 压缩包中无法被 spawn 执行
const asarPath = path.join(__dirname, '..', 'dist', 'win-unpacked', 'resources', 'app.asar');
if (fs.existsSync(asarPath)) {
  fs.unlinkSync(asarPath);
  console.log('已删除 app.asar');
}

// 2. 同步静态文件（覆盖 extraResources 中可能被过滤或版本不一致的文件）
const filesToSync = [
  'quote_system/feishu_client.py',
  'quote_system/web_app.py',
  'quote_system/auto_fill_tk.py',
  // ... 更多文件
];
for (const f of filesToSync) {
  const src = path.join(srcApp, f);
  const dest = path.join(distApp, f);
  if (fs.existsSync(src)) {
    copy(src, dest);
  }
}

// 3. 创建运行时需要的空目录
const outputDirs = [
  'outputs/logs/auto_quote',
  'outputs/settlement_tracker',
  'outputs/uploads',
];
for (const d of outputDirs) {
  ensureDir(path.join(distApp, d));
}
```

**为什么删除 app.asar？**

- electron-builder 默认将 `node_modules/`、`main.js`、`preload.js` 等打包进 `app.asar`
- `extraResources` 中的文件不会进入 asar，但一旦 asar 存在，Electron 的资源查找优先走 asar
- 删除 asar 后，Electron 直接从 `resources/app/` 目录加载所有文件
- 确保 `main.js` 中 `path.join(process.resourcesPath, 'app')` 能找到 `flask_server/flask_server.exe`

---

## 输出目录结构

```
dist/
├── 报价工具.exe                  ← 最终产出！111MB，双击运行
├── win-unpacked/                 ← 解压版目录（用于调试或直接运行）
│   ├── electron.exe
│   └── resources/
│       └── app/
│           ├── flask_server/     ← 编译好的 Flask 后端
│           ├── quote_system/     ← Python 源码
│           ├── settlement/       ← 结算模块
│           ├── config/           ← 配置文件
│           ├── 模板/             ← Excel/Word 模板
│           ├── main.js           ← Electron 主进程
│           ├── preload.js        ← 预加载脚本
│           └── package.json
├── flask_build/                  ← 临时：PyInstaller 产物的原始版本
│   └── flask_server/
│       ├── flask_server.exe
│       └── _internal/
└── flask_work/                   ← 临时：PyInstaller 工作目录
```

---

## Mac/Linux 打包说明

当前 `package.json` 的 `build` 命令只支持 Windows。Electron 本身是跨平台的，但以下变更才能支持其他平台：

### macOS
```json
"mac": {
  "target": ["dmg", "zip"],
  "arch": ["x64", "arm64"]
}
```
- 需在 macOS 上运行
- Flask server 编译需 `pyinstaller --onedir --name flask_server run_flask.py`
- Python 依赖与 Windows 一致

### Linux
```json
"linux": {
  "target": ["AppImage", "deb"],
  "arch": ["x64"]
}
```
- 需在 Linux 上运行
- Flask server 编译同上，但产物为 `flask_server`（无 .exe 后缀）
- `main.js` 需自动适配 `flask_server` vs `flask_server.exe`

---

## 故障排查指南

### 1. Flask 装不上 / 缺少模块

**症状**：`python.exe -c "import flask"` 报 `ModuleNotFoundError`
**修复**：
```bash
# workbuddy 的 Python 环境
C:/Users/admin/.workbuddy/binaries/python/versions/3.13.12/python.exe \
  -m pip install flask lxml openpyxl python-docx lark-oapi markupsafe \
  -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 2. PyInstaller 找不到隐藏导入

**症状**：打包后运行 `flask_server.exe` 报 `No module named 'xxx'`
**修复**：在 `app/flask_server.spec` 的 `hiddenimports` 列表中添加对应模块名

### 3. electron-builder 报 "Device or resource busy"

**症状**：`rm -rf dist` 报权限错误，文件被占用
**原因**：Windows Explorer 打开了 `dist/` 目录，或之前的 `报价工具.exe` 正在运行
**修复**：
```bash
taskkill /F /IM 报价工具.exe
taskkill /F /IM flask_server.exe
taskkill /F /IM explorer.exe && start explorer  # 重启资源管理器
```

### 4. 打包后 Flask 启动报错

第一步诊断：`flask_server.exe` 能否单独正常运行？
```bash
cd dist/win-unpacked/resources/app/flask_server
flask_server.exe
# 应看到 "Running on http://127.0.0.1:5000"
```

如果独立运行正常，检查 main.js 的 `startFlaskServer()` 路径查找逻辑。

### 5. `dist/flask_build` 被锁

**原因**：前一次编译的 `flask_server.exe` 仍在运行
**修复**：
```bash
taskkill /F /IM flask_server.exe
rm -rf dist/flask_build dist/flask_work
```

---

## Git 相关

### .gitignore

```gitignore
dist/                     ← 忽略所有打包产物
app/flask_server/         ← 忽略构建时复制的 flask_server
```

### 构建相关的变更文件清单

| 文件 | 用途 | 是否提交 |
|------|------|----------|
| `package.json` | npm 脚本 + electron-builder 配置 | ✅ |
| `scripts/prebuild.js` | 预构建：复制 flask_server 到 app/ | ✅ |
| `scripts/postbuild.js` | 后处理：删除 asar + 同步文件 | ✅ |
| `app/flask_server.spec` | PyInstaller 编译配置 | ✅ |
| `app/run_flask.py` | Flask 入口脚本 | ✅ |
| `main.js` | Electron 主进程，启动 Flask | ✅ |

---

## 完整构建命令速查

```bash
# 1. 完整构建（推荐）
npm run build

# 2. 仅编译 Flask 后端
npm run build-flask

# 3. 测试 Flask exe 是否正常
cd dist/flask_build/flask_server
./flask_server.exe                # 控制台启动
curl http://127.0.0.1:5000/      # 应返回 HTML

# 4. 仅测试 Python 源码模式开发
cd app
python -m quote_system.web_app

# 5. 清理构建产物
rm -rf dist app/flask_server
```
