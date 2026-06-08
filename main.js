const { app, BrowserWindow, Menu, shell, dialog, ipcMain } = require('electron');
const path = require('path');
const { spawn } = require('child_process');
const fs = require('fs');
const http = require('http');

let mainWindow;
let pythonProcess = null;

function getAppDir() {
  if (app.isPackaged) {
    return path.join(process.resourcesPath, 'app');
  }
  return path.join(__dirname, 'app');
}

function getHistoryDir() {
  return path.join(getAppDir(), '报价单历史');
}

function findWpsExecutable() {
  const possiblePaths = [];
  const localAppData = process.env.LOCALAPPDATA || '';
  if (localAppData) {
    const kingsoftDir = path.join(localAppData, 'Kingsoft', 'WPS Office');
    if (fs.existsSync(kingsoftDir)) {
      const dirs = fs.readdirSync(kingsoftDir).filter(d => /^\d/.test(d));
      for (const d of dirs) {
        const wpsPath = path.join(kingsoftDir, d, 'office6', 'wps.exe');
        if (fs.existsSync(wpsPath)) possiblePaths.push(wpsPath);
      }
    }
  }
  for (const base of [process.env['PROGRAMFILES(X86)'], process.env.PROGRAMFILES, 'C:\\Program Files (x86)', 'C:\\Program Files']) {
    if (!base) continue;
    const kingsoftDir = path.join(base, 'Kingsoft', 'WPS Office');
    if (fs.existsSync(kingsoftDir)) {
      const entries = fs.readdirSync(kingsoftDir).filter(d => /^\d/.test(d));
      for (const d of entries) {
        const wpsPath = path.join(kingsoftDir, d, 'office6', 'wps.exe');
        if (fs.existsSync(wpsPath)) possiblePaths.push(wpsPath);
      }
    }
  }
  return possiblePaths.length > 0 ? possiblePaths[0] : null;
}

const wpsExecutablePath = findWpsExecutable();

function openWithWps(filePath) {
  const { execFile } = require('child_process');
  return new Promise((resolve, reject) => {
    execFile(wpsExecutablePath, [filePath], (error) => {
      if (error) reject(error);
      else resolve();
    });
  });
}

ipcMain.on('open-file', (event, filePath) => {
  const fs = require('fs');
  const absolutePath = path.resolve(filePath);
  if (!fs.existsSync(absolutePath)) {
    dialog.showErrorBox('打开失败', '文件不存在: ' + filePath);
    return;
  }
  if (wpsExecutablePath) {
    openWithWps(absolutePath).catch(() => {
      shell.openPath(absolutePath).catch((err) => {
        dialog.showErrorBox('打开失败', '无法打开文件: ' + filePath);
      });
    });
  } else {
    shell.openPath(absolutePath).catch((err) => {
      dialog.showErrorBox('打开失败', '无法打开文件: ' + filePath);
    });
  }
});

ipcMain.on('open-file-location', (event, filePath) => {
  console.log('open-file-location received:', filePath);
  if (!filePath || typeof filePath !== 'string') {
    dialog.showErrorBox('路径无效', '无法打开文件夹：路径为空或格式错误');
    return;
  }
  const fs = require('fs');
  const absolutePath = path.resolve(filePath);
  if (!fs.existsSync(absolutePath)) {
    dialog.showErrorBox('文件不存在', `无法打开文件所在位置，文件不存在:\n${absolutePath}`);
    return;
  }
  shell.showItemInFolder(absolutePath);
});

ipcMain.on('open-folder', (event, folderPath) => {
  console.log('open-folder received:', folderPath);
  if (!folderPath || typeof folderPath !== 'string') {
    dialog.showErrorBox('路径无效', '无法打开文件夹：路径为空或格式错误');
    return;
  }
  const fs = require('fs');
  const absolutePath = path.resolve(folderPath);
  if (!fs.existsSync(absolutePath)) {
    dialog.showErrorBox('文件夹不存在', `文件夹不存在:\n${absolutePath}`);
    return;
  }
  shell.openPath(absolutePath);
});

ipcMain.handle('select-folder', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openDirectory'],
    title: '选择文件夹',
  });
  if (result.canceled || result.filePaths.length === 0) {
    return { canceled: true, filePath: null };
  }
  return { canceled: false, filePath: result.filePaths[0] };
});

ipcMain.handle('copy-file-to', async (event, filePath) => {
  if (!fs.existsSync(filePath)) {
    return { success: false, error: '文件不存在' };
  }
  const result = await dialog.showSaveDialog({
    title: '复制文件到...',
    defaultPath: path.basename(filePath),
  });
  if (result.canceled || !result.filePath) {
    return { success: false, error: '已取消' };
  }
  try {
    fs.copyFileSync(filePath, result.filePath);
    return { success: true, destPath: result.filePath };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('copy-to-clipboard', async (event, filePath) => {
  if (!fs.existsSync(filePath)) {
    return { success: false, error: '文件不存在' };
  }
  
  try {
    const { exec } = require('child_process');
    const absolutePath = path.resolve(filePath);
    
    return new Promise((resolve) => {
      exec(`powershell -Command "Set-Clipboard -Path '${absolutePath.replace(/'/g, "''")}'"`, (err) => {
        if (err) {
          resolve({ success: false, error: err.message });
        } else {
          resolve({ success: true });
        }
      });
    });
  } catch (err) {
    return { success: false, error: err.message };
  }
});

ipcMain.on('request-drag', (event, filePath) => {
  if (fs.existsSync(filePath)) {
    mainWindow.webContents.startDrag({
      file: filePath
    });
  }
});

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 1024,
    minHeight: 600,
    title: '翻译报价系统',
    fullscreen: false,
    maximizable: true,
    show: false,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      enableRemoteModule: true,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  mainWindow.maximize();

  const menuTemplate = [
    {
      label: '文件',
      submenu: [
        {
          label: '打开报价单目录',
          click: () => {
            shell.openPath(getHistoryDir());
          }
        },
        { type: 'separator' },
        {
          label: '退出',
          role: 'quit'
        }
      ]
    },
    {
      label: '帮助',
      submenu: [
        {
          label: '开发者工具',
          accelerator: 'F12',
          click: () => {
            if (mainWindow) mainWindow.webContents.toggleDevTools();
          }
        },
        { type: 'separator' },
        {
          label: '关于',
          click: () => {
            dialog.showMessageBox({
              title: '关于',
              message: '翻译报价系统 v1.0.0',
              detail: '基于 Electron + Flask 构建的翻译报价单生成工具'
            });
          }
        }
      ]
    }
  ];

  const menu = Menu.buildFromTemplate(menuTemplate);
  Menu.setApplicationMenu(menu);

  startFlaskServer();
  waitForFlask(() => {
    if (mainWindow) {
      mainWindow.loadURL('http://127.0.0.1:5000');
      mainWindow.show();
    }
  });

  mainWindow.webContents.on('will-navigate', (event, url) => {
    const parsedUrl = new URL(url);
    if (parsedUrl.origin !== 'http://127.0.0.1:5000') {
      event.preventDefault();
    }
  });

  mainWindow.webContents.on('did-create-window', (childWindow) => {
    childWindow.on('close', (event) => {
      event.preventDefault();
      childWindow.destroy();
    });
  });

  mainWindow.webContents.session.setPermissionRequestHandler((webContents, permission, callback) => {
    callback(false);
  });

  mainWindow.webContents.on('start-drag', (event, info) => {
    if (info.file && info.file.path) {
      event.preventDefault();
      mainWindow.webContents.startDrag({
        file: info.file.path,
        icon: path.join(__dirname, 'icon.png')
      });
    }
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
    if (pythonProcess) {
      pythonProcess.kill();
    }
  });
}

function startFlaskServer() {
  const appDir = getAppDir();
  console.log('应用目录:', appDir);

  const flaskExe = path.join(appDir, 'flask_server.exe');
  const useExe = fs.existsSync(flaskExe);

  if (useExe) {
    console.log('使用 flask_server.exe');
    pythonProcess = spawn(flaskExe, [], {
      cwd: appDir,
      detached: false,
      stdio: ['pipe', 'pipe', 'pipe']
    });
  } else {
    const pythonPath = findPython();
    console.log('Python路径:', pythonPath);
    const env = Object.create(process.env);
    env.PYTHONPATH = appDir;
    pythonProcess = spawn(pythonPath, ['-m', 'quote_system.web_app'], {
      cwd: appDir,
      env: env,
      detached: false,
      stdio: ['pipe', 'pipe', 'pipe']
    });
  }

  pythonProcess.stdout.on('data', (data) => {
    console.log(`Flask stdout: ${data.toString()}`);
  });

  pythonProcess.stderr.on('data', (data) => {
    console.error(`Flask stderr: ${data.toString()}`);
  });

  pythonProcess.on('close', (code) => {
    console.log(`Flask 进程退出，代码: ${code}`);
    if (mainWindow && !app.isQuitting) {
      dialog.showMessageBox({
        type: 'warning',
        title: '服务器断开',
        message: 'Flask 服务器意外断开，应用将退出',
        buttons: ['确定']
      }).then(() => {
        app.quit();
      });
    }
  });

  pythonProcess.on('error', (err) => {
    console.error('启动 Flask 失败:', err);
    dialog.showErrorBox('启动失败', '无法启动 Flask 服务器: ' + err.message);
    app.quit();
  });
}

function waitForFlask(callback) {
  const maxAttempts = 50;
  let attempt = 0;

  function check() {
    attempt++;
    const req = http.get('http://127.0.0.1:5000', (res) => {
      res.resume();
      callback();
    });
    req.on('error', () => {
      if (attempt < maxAttempts) {
        setTimeout(check, 100);
      } else {
        callback();
      }
    });
    req.setTimeout(500, () => {
      req.destroy();
      if (attempt < maxAttempts) {
        setTimeout(check, 100);
      } else {
        callback();
      }
    });
  }
  check();
}

function findPython() {
  const possiblePaths = [
    path.join(process.env.LOCALAPPDATA, 'Python', 'pythoncore-3.14-64', 'python.exe'),
    'python',
    'python3',
    'py',
    path.join(process.env.PROGRAMFILES, 'Python311', 'python.exe'),
    path.join(process.env.PROGRAMFILES, 'Python310', 'python.exe'),
    path.join(process.env.PROGRAMFILES, 'Python39', 'python.exe'),
    path.join(process.env.LOCALAPPDATA, 'Programs', 'Python', 'Python311', 'python.exe'),
    path.join(process.env.LOCALAPPDATA, 'Programs', 'Python', 'Python310', 'python.exe'),
    path.join(process.env.LOCALAPPDATA, 'Programs', 'Python', 'Python39', 'python.exe')
  ];

  for (const p of possiblePaths) {
    try {
      if (fs.existsSync(p)) {
        return p;
      }
    } catch (e) {
      continue;
    }
  }
  
  return 'python';
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

app.on('before-quit', () => {
  app.isQuitting = true;
  if (pythonProcess) {
    pythonProcess.kill();
  }
});