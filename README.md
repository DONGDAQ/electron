# 翻译报价系统

## 使用说明

### 方式一：直接运行 Flask 服务（推荐）

1. 双击运行 `run_flask.bat`
2. 在浏览器中访问 `http://127.0.0.1:5000`

### 方式二：开发模式

```bash
cd app
python -m quote_system.web_app
```

### 方式三：Electron 模式（需要安装依赖）

```bash
npm install
npm start
```

## 项目结构

```
electron/
├── app/                    # Flask 应用目录
│   ├── quote_system/       # 报价系统核心代码
│   ├── config/             # 配置文件
│   ├── 模板/               # Excel 模板
│   ├── 报价单历史/         # 生成的报价单
│   └── outputs/            # 输出文件
├── main.js                 # Electron 主进程
├── package.json            # Electron 依赖配置
├── run_flask.bat           # Flask 启动脚本
└── start.bat               # Electron 启动脚本
```

## 功能特性

- 支持多个项目模板（马娘、战双、BANG2等）
- 从 MemoQ HTML 文件提取字数统计
- 自动生成 Excel 报价单
- 支持自定义语种和单价配置
- 支持多文件批量处理
- 支持网页自动填表模式：从飞书表格 N 列下载 MemoQ HTML，并回填 O-Y 列字数统计

## TK 飞书表格自动填表

通过飞书 API 直接操作 TK 表格副本。

运行：

```bash
cd app
python -m quote_system.auto_fill_tk
```

脚本会自动扫描 N 列为空的行，下载 HTML 解析后填写 O-Y 列。

## 注意事项

1. 需要安装 Python 3.9+
2. 需要安装依赖：`pip install -r app/requirements.txt`
3. 首次运行前请确保配置文件存在
