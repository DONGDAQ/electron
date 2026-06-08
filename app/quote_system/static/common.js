/* common.js — 共享工具函数 */
(function() {
  var t = localStorage.getItem('theme') || 'dark';
  if (t === 'light') document.body.classList.add('light-theme');
})();

function setTheme(theme) {
  if (theme === 'light') document.body.classList.add('light-theme');
  else document.body.classList.remove('light-theme');
  localStorage.setItem('theme', theme);
  updateThemeBtns();
}

function updateThemeBtns() {
  var isLight = document.body.classList.contains('light-theme');
  var db = document.getElementById('themeDarkBtn');
  var lb = document.getElementById('themeLightBtn');
  if (db) { db.style.borderColor = isLight ? '' : 'var(--accent)'; db.style.color = isLight ? '' : 'var(--accent)'; }
  if (lb) { lb.style.borderColor = isLight ? 'var(--accent)' : ''; lb.style.color = isLight ? 'var(--accent)' : ''; }
}

function escHtml(text) {
  if (!text) return '';
  var d = document.createElement('div');
  d.textContent = text;
  return d.innerHTML;
}

function showToast(message, isError) {
  var existing = document.querySelector('.copy-toast');
  if (existing) existing.remove();
  var toast = document.createElement('div');
  toast.className = 'copy-toast';
  toast.style.cssText = 'position:fixed;top:20px;left:50%;transform:translateX(-50%);background:' + (isError ? '#ef4444' : '#10b981') + ';color:white;padding:12px 24px;border-radius:8px;font-size:14px;z-index:9999;box-shadow:0 4px 12px rgba(0,0,0,.15);';
  toast.textContent = message;
  document.body.appendChild(toast);
  setTimeout(function() { toast.style.transition = 'opacity .3s'; toast.style.opacity = '0'; setTimeout(function() { toast.remove(); }, 300); }, 1000);
}

function openFileInExcel(filePath) {
  if (window.electronAPI) {
    window.electronAPI.openFile(filePath);
  } else {
    fetch('/open-file', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ path: filePath }),
    }).catch(function(e) { console.error('打开文件失败:', e); });
  }
}

function openFileLocation(p) {
  if (window.electronAPI) window.electronAPI.openFileLocation(p);
  else window.open('/open-folder?path=' + encodeURIComponent(p));
}

function openTargetFolder(p) {
  if (window.electronAPI) window.electronAPI.openFolder(p);
  else window.open('/open-folder?path=' + encodeURIComponent(p));
}

function copyFileTo(p) {
  if (window.electronAPI) {
    window.electronAPI.copyToClipboard(p).then(function(r) {
      if (r.success) showToast('已复制到剪贴板，可粘贴到微信');
      else showToast('复制失败: ' + (r.error || ''), true);
    });
  } else {
    showToast('此功能仅在桌面版可用', true);
  }
}

function dragFile(p, name) {
  if (window.electronAPI) window.electronAPI.requestDrag(p);
}

function openSettings() {
  var modal = document.getElementById('settingsModal');
  if (modal) {
    modal.style.display = 'flex';
    if (typeof updateThemeBtns === 'function') updateThemeBtns();
    fetch('/api/base-paths')
      .then(function(r) { return r.json(); })
      .then(function(data) {
        if (data.status === 'success') {
          var qp = document.getElementById('settingsQuotePath');
          var sp = document.getElementById('settingsSettlePath');
          if (qp) qp.value = data.quote_history_base || '';
          if (sp) sp.value = data.settlement_base || '';
        }
      }).catch(function() {});
  }
}

function closeSettings() {
  var modal = document.getElementById('settingsModal');
  if (modal) modal.style.display = 'none';
}

function browseFolder(inputId) {
  if (window.electronAPI) {
    window.electronAPI.selectFolder().then(function(result) {
      if (!result.canceled && result.filePaths.length > 0) {
        document.getElementById(inputId).value = result.filePaths[0];
      }
    });
  }
}

function saveSettings() {
  var quotePath = document.getElementById('settingsQuotePath').value.trim();
  var settlePath = document.getElementById('settingsSettlePath').value.trim();
  fetch('/api/base-paths', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ quote_history_base: quotePath, settlement_base: settlePath }),
  }).then(function(r) { return r.json(); })
    .then(function(d) {
      if (d.status === 'success') { showToast('设置已保存'); closeSettings(); }
      else showToast('保存失败: ' + d.message, true);
    }).catch(function(e) { showToast('保存失败', true); });
}

function fmtDateTime(d) {
  var pad = function(n) { return String(n).padStart(2, '0'); };
  return d.getFullYear() + '-' + pad(d.getMonth()+1) + '-' + pad(d.getDate()) + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes());
}

function fmtSize(b) {
  if (b < 1024) return b + ' B';
  if (b < 1048576) return (b/1024).toFixed(1) + ' KB';
  return (b/1048576).toFixed(1) + ' MB';
}

document.addEventListener('click', function(e) {
  var modal = document.getElementById('settingsModal');
  if (modal && e.target === modal) closeSettings();
});
