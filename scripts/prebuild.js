const fs = require('fs');
const path = require('path');

const srcDir = path.join(__dirname, '..', 'dist', 'flask_build', 'flask_server');
const destDir = path.join(__dirname, '..', 'app', 'flask_server');

if (!fs.existsSync(srcDir)) {
  console.error('flask_server build not found, run "npm run build-flask" first');
  process.exit(1);
}

// 删除旧的 app/flask_server/
if (fs.existsSync(destDir)) {
  fs.rmSync(destDir, { recursive: true });
}

// 复制 flask_server 到 app/ 下，让 extraResources 打包进去
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
console.log('已复制 flask_server 到 app/flask_server/');
