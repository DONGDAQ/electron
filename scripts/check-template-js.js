const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const os = require('os');

const templatePath = path.join(__dirname, '..', 'app', 'quote_system', 'templates', 'index.html');
const content = fs.readFileSync(templatePath, 'utf-8');
const lines = content.split('\n');

// 找到内联 <script> 块（跳过 src= 的外部脚本）
let scriptStart = -1, scriptEnd = -1;
for (let i = 0; i < lines.length; i++) {
  if (lines[i].trim() === '<script>' && scriptStart === -1) {
    scriptStart = i;
  }
  if (lines[i].trim() === '</script>' && scriptStart !== -1) {
    scriptEnd = i;
    break;
  }
}

if (scriptStart === -1 || scriptEnd === -1) {
  console.error('未找到内联 <script> 块');
  process.exit(1);
}

// 提取 JS 内容，保留原始行号
const jsLines = lines.slice(scriptStart + 1, scriptEnd);
const cleaned = jsLines.map(line => {
  // 去掉 {% if ... %} / {% endif %} 等 Jinja2 控制流
  if (/^\s*\{%.*%\}\s*$/.test(line)) return '';
  // {{ ... }} 替换为合法 JS 表达式
  return line.replace(/\{\{.*?\}\}/g, "'__JINJA__'");
}).join('\n');

// 写入临时文件
const tmpFile = path.join(os.tmpdir(), 'check-template-js.js');
fs.writeFileSync(tmpFile, cleaned, 'utf-8');

try {
  execSync(`node --check "${tmpFile}"`, { stdio: 'pipe' });
  console.log(`✓ JS 语法检查通过（第 ${scriptStart + 1}-${scriptEnd + 1} 行）`);
} catch (e) {
  // 把临时文件的行号映射回原始模板行号
  const stderr = e.stderr.toString();
  const mapped = stderr.replace(/check-template-js\.js:(\d+)/g, (_, line) => {
    return `index.html:${scriptStart + 1 + parseInt(line)}`;
  });
  console.error(mapped);
  process.exit(1);
} finally {
  fs.unlinkSync(tmpFile);
}
