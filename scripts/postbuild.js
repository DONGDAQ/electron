const fs = require('fs');
const path = require('path');

const distApp = path.join(__dirname, '..', 'dist', 'win-unpacked', 'resources', 'app');
const srcApp = path.join(__dirname, '..', 'app');

function copy(src, dest) {
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  fs.copyFileSync(src, dest);
}

function ensureDir(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

// 1. 删除 app.asar（强制 Electron 使用 app/ 目录）
const asarPath = path.join(__dirname, '..', 'dist', 'win-unpacked', 'resources', 'app.asar');
if (fs.existsSync(asarPath)) {
  fs.unlinkSync(asarPath);
  console.log('已删除 app.asar');
}

// 2. 复制 flask_server.exe（如果存在）
const flaskExe = path.join(__dirname, '..', 'dist', 'flask_build', 'flask_server.exe');
if (fs.existsSync(flaskExe)) {
  copy(flaskExe, path.join(distApp, 'flask_server.exe'));
  console.log('已复制 flask_server.exe');
}

// 3. 同步关键文件（这些文件可能被 extraResources filter 排除或需要覆盖）
const filesToSync = [
  'quote_system/feishu_client.py',
  'quote_system/web_app.py',
  'quote_system/auto_fill_tk.py',
  'quote_system/auto_quote.py',
  'quote_system/auto_quote_zhan_shuang.py',
  'quote_system/auto_quote_zhan_shuang_feishu.py',
  'quote_system/auto_quote_4399.py',
  'quote_system/static/styles.css',
  'quote_system/templates/index.html',
  'quote_system/templates/report.html',
  'settlement/settlement_tracker.py',
  'auto_quote_scheduled.py',
  'watchdog_task.vbs',
];

for (const f of filesToSync) {
  const src = path.join(srcApp, f);
  const dest = path.join(distApp, f);
  if (fs.existsSync(src)) {
    copy(src, dest);
  }
}
console.log(`已同步 ${filesToSync.length} 个文件到 dist`);

// 4. 确保 outputs 目录结构存在（不复制临时文件）
const outputDirs = [
  'outputs/logs/auto_quote',
  'outputs/settlement_tracker',
  'outputs/uploads',
];
for (const d of outputDirs) {
  ensureDir(path.join(distApp, d));
}
console.log('已创建 outputs 目录结构');
