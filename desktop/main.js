const { app, BrowserWindow, dialog } = require("electron");
const { spawn } = require("child_process");
const http = require("http");
const net = require("net");
const path = require("path");

let backendProcess;
let mainWindow;
let isQuitting = false;

function findFreePort() {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.unref();
    server.on("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const { port } = server.address();
      server.close(() => resolve(port));
    });
  });
}

function bundledBackendPath() {
  return path.join(process.resourcesPath, "backend", "inventory_backend.exe");
}

function startBackend(port) {
  const environment = {
    ...process.env,
    INVENTORY_NO_BROWSER: "1",
    INVENTORY_PORT: String(port),
    INVENTORY_DATA_ROOT: path.join(app.getPath("userData"), "data"),
    INVENTORY_INSTALL_ROOT: app.isPackaged ? path.dirname(process.resourcesPath) : "",
  };

  if (app.isPackaged) {
    return spawn(bundledBackendPath(), [], { env: environment, windowsHide: true, stdio: "ignore" });
  }

  const projectRoot = path.join(__dirname, "..");
  const python = process.env.INVENTORY_PYTHON || "py";
  return spawn(python, ["-3", path.join(projectRoot, "app.py")], {
    cwd: projectRoot,
    env: environment,
    windowsHide: true,
    stdio: "ignore",
  });
}

function stopBackend() {
  if (backendProcess && !backendProcess.killed) backendProcess.kill();
  backendProcess = undefined;
}

function waitForBackend(address, attempts = 80) {
  return new Promise((resolve, reject) => {
    const probe = (remaining) => {
      const request = http.get(`${address}/api/state`, (response) => {
        response.resume();
        if (response.statusCode === 200) {
          resolve();
        } else if (remaining > 0) {
          setTimeout(() => probe(remaining - 1), 150);
        } else {
          reject(new Error(`后端服务返回状态 ${response.statusCode}`));
        }
      });
      request.on("error", () => {
        if (remaining > 0) {
          setTimeout(() => probe(remaining - 1), 150);
        } else {
          reject(new Error("本地服务未能启动"));
        }
      });
      request.setTimeout(1000, () => request.destroy());
    };
    probe(attempts);
  });
}

async function createWindow() {
  let address;
  let lastError;
  for (let attempt = 0; attempt < 3; attempt += 1) {
    const port = await findFreePort();
    const candidateAddress = `http://127.0.0.1:${port}`;
    backendProcess = startBackend(port);
    try {
      await waitForBackend(candidateAddress);
      address = candidateAddress;
      break;
    } catch (error) {
      lastError = error;
      stopBackend();
    }
  }
  if (!address) throw lastError || new Error("本地服务未能启动");

  backendProcess.on("error", (error) => {
    if (isQuitting) return;
    dialog.showErrorBox("库存工作台无法启动", `无法启动库存服务：${error.message}`);
    app.quit();
  });
  backendProcess.on("exit", (code) => {
    if (isQuitting) return;
    dialog.showErrorBox("库存服务已停止", `库存服务意外退出（代码 ${code ?? "未知"}），请重新启动库存工作台。`);
    app.quit();
  });

  mainWindow = new BrowserWindow({
    title: "库存工作台",
    width: 1440,
    height: 920,
    minWidth: 1180,
    minHeight: 720,
    show: false,
    autoHideMenuBar: true,
    backgroundColor: "#f4f6f2",
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });
  mainWindow.once("ready-to-show", () => mainWindow.show());
  await mainWindow.loadURL(address);
}

const hasSingleInstanceLock = app.requestSingleInstanceLock();
if (!hasSingleInstanceLock) {
  app.quit();
} else {
  app.on("second-instance", () => {
    if (!mainWindow) return;
    if (mainWindow.isMinimized()) mainWindow.restore();
    mainWindow.focus();
  });

  app.whenReady().then(createWindow).catch((error) => {
    stopBackend();
    dialog.showErrorBox("库存工作台无法启动", error.message);
    app.quit();
  });
}

app.on("window-all-closed", () => app.quit());
app.on("before-quit", () => {
  isQuitting = true;
  stopBackend();
});
