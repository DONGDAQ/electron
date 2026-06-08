# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 启动/运行命令

```bash
# Flask 开发模式（推荐日常开发）
cd app && python -m quote_system.web_app

# Electron 桌面应用
npm install && npm start

# 构建 Windows 安装包
npm run build

# 测试（只通过 test.bat，不用浏览器）
test.bat
```

## 架构概览

**技术栈：Electron + Flask + openpyxl**

- `main.js` — Electron 主进程，启动时 spawn Flask 子进程（`python -m quote_system.web_app`），Flask 就绪后加载 `http://127.0.0.1:5000`
- `preload.js` — 通过 contextBridge 暴露 `electronAPI`（openFile/openFileLocation/openFolder）给渲染进程
- `app/quote_system/web_app.py` — Flask 应用入口，所有路由在此。前端页面使用 Jinja2 模板渲染（`templates/`），非 SPA 架构
- `app/quote_system/projects.py` — 项目定义中心。`PROJECTS` 字典注册所有项目（马娘、战双、幻塔等），每个项目绑定模板、语种/单价、generator 类型
- `app/quote_system/generator.py` — 报价单生成核心。根据 `project.generator` 字段分发到不同的填充策略（`fill_punctuation_project` / `fill_bang2` / `fill_zhan_shuang` / `fill_huan_ta`）
- `app/quote_system/memoq_html.py` — MemoQ HTML 解析器，提取各匹配段的字数统计
- `app/quote_system/feishu_client.py` — 飞书 API 客户端（表格读写、附件下载/上传、权限管理）
- `app/quote_system/config.py` — 项目语种/单价配置持久化（`config/language_config.json`）
- `app/settlement/` — 结算模块，按项目/月份从飞书读取数据生成结算单 Excel + 验收单 Word

**数据流：**
1. 用户上传 MemoQ HTML → 解析字数统计
2. 按项目模板生成 Excel 报价单（openpyxl 填充公式和数值）
3. 文件保存到 `报价单历史/公司/项目/`，前端点击通过 `/open-file` 接口直接打开

**自动报价（非交互式）：**
- `auto_quote.py` → 完美世界/幻塔：从飞书表格拉取需求 → 下载 HTML → 生成报价单 → 回传飞书
- `auto_quote_zhan_shuang.py` → 战双：从邮件解析需求自动报价
- `auto_quote_diezhi.py` → 叠纸：批次报价模式
- `auto_fill_tk.py` → TK项目：API 方式从飞书下载 N 列 HTML → 解析 → 回填 O-Y 列

**前端文件：**
- 图标库 `lucide.min.js` 从本地 static 目录引用，不能使用 CDN
- 生成的文件用 `/open-file` 接口直接打开，不弹出保存对话框

## pip 安装

中国国内环境，所有 pip install 必须用清华镜像：
```
pip install <包名> -i https://pypi.tuna.tsinghua.edu.cn/simple
```

## Python 路径

编写涉及 app 目录下模块的脚本时，自动加上：
```python
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
```

## 输出路径

结算文件统一按 `结算/年月/公司/项目/` 存放。

## 文件编码

所有 Python 文件默认 UTF-8，运行时设置环境变量 `PYTHONIOENCODING=utf-8`。

## 代码风格

- 不写注释（除非逻辑非直观）
- 不做不必要的抽象
- 简洁优先

## Git 提交

中文提交信息，格式：`类型: 描述`（如 `feat:` `fix:` `chore:` 等）

## 测试方式

只用 `test.bat` 启动测试，不用浏览器。

## 文件操作

所有生成的文件前端点击后直接打开（用 `/open-file` 接口），不要弹出保存对话框。文件会自动保存到对应目录。

## 添加新项目结算功能

结算流程：读取数据 → 生成对账单 xlsx → 客户确认 → 生成盖章版结算单 PDF（+ Invoice PDF）

### 1. 创建结算生成模块

在 `app/settlement/` 下新建 `generate_settlement_<project>.py`，参考 `generate_settlement.py`（幻塔）或 `generate_settlement_mamian.py`（马娘）：

```python
# 必须实现的函数
def read_feishu_data(year, month) -> list[dict]: ...
def generate_settlement_excel(records, year, month, output_dir) -> Path: ...
def generate_acceptance_docx(records, year, month, output_dir) -> Path: ...
```

### 2. 准备模板文件

Excel 结算单模板和 Word 验收单模板放到 `app/模板/`。

### 3. 注册 Flask 路由

在 `web_app.py` 中添加 3 个路由：
- `GET /api/settlement_<project>/preview` — 读取飞书数据预览
- `POST /api/settlement_<project>/generate` — 生成结算文件
- 如有盖章版需求，加 `POST /api/settlement_<project>/generate-sealed`

### 4. 注册前端配置

在 `index.html` 的 `SETTLEMENT_PROJECTS` 对象中添加项目：
```javascript
'<project_key>': { preview: '/api/settlement_<project>/preview', generate: '/api/settlement_<project>/generate' },
```

### 5. 结算文件输出目录

统一按 `结算/{year}年{month}月/{公司}/{项目}/` 存放。外币结算项目（TK、4399）不需要开票请求。
