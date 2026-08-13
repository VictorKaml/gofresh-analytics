const { app, BrowserWindow, dialog, Menu, ipcMain } = require("electron");
const path = require("path");
const fs = require("fs");
const { spawn, execFileSync } = require("child_process");
const { autoUpdater } = require("electron-updater");

let mainWindow = null;
let splashWindow = null;
let currentFilePath = null;

const TEMPLATE_PATH = path.join(__dirname, "assets", "dashboard-template.html");
const DEFAULT_XLSX = app.isPackaged
  ? path.join(process.resourcesPath, "app.asar.unpacked", "data", "default.xlsx")
  : path.join(__dirname, "data", "default.xlsx");
const AGGREGATE_SCRIPT = path.join(__dirname, "src", "aggregate_workbook.py");
const APP_ICON = path.join(__dirname, "assets", "icons", "icon.png");

// ---------------------------------------------------------------------------
// Aggregator resolution
// ---------------------------------------------------------------------------

const BUNDLED_EXE_NAME =
  process.platform === "win32"
    ? "aggregate_workbook.exe"
    : "aggregate_workbook";

function bundledPythonPath() {
  const resourcesDir = app.isPackaged
    ? process.resourcesPath
    : path.join(__dirname, "build-resources");
  return path.join(
    resourcesDir,
    "python",
    "aggregate_workbook",
    BUNDLED_EXE_NAME,
  );
}

const SYSTEM_PYTHON_CANDIDATES =
  process.platform === "win32" ? ["python", "py"] : ["python3", "python"];

let resolvedSystemPython = null;

function resolveSystemPython() {
  if (resolvedSystemPython) return resolvedSystemPython;
  for (const candidate of SYSTEM_PYTHON_CANDIDATES) {
    try {
      execFileSync(candidate, ["--version"], { stdio: "ignore" });
      resolvedSystemPython = candidate;
      return resolvedSystemPython;
    } catch (_) {
      // try the next candidate
    }
  }
  return null;
}

/**
 * Returns true if the bundled binary exists AND is newer than the source
 * script. In development we ALWAYS prefer the live script so edits are
 * picked up without running `npm run build:python`.
 */
function shouldUseBundledBinary() {
  const bundled = bundledPythonPath();
  if (!fs.existsSync(bundled)) return false;

  // In dev, always prefer the script so Python fixes are immediate
  if (!app.isPackaged) {
    console.log("[aggregator] Dev mode — preferring live Python script");
    return false;
  }

  // In packaged builds, only use the binary if it is newer than the script
  try {
    const binMtime = fs.statSync(bundled).mtimeMs;
    const scriptMtime = fs.statSync(AGGREGATE_SCRIPT).mtimeMs;
    if (binMtime < scriptMtime) {
      console.warn(
        `[aggregator] Bundled binary is OLDER than ${path.basename(
          AGGREGATE_SCRIPT,
        )}. Falling back to system Python so the latest script runs.`,
      );
      return false;
    }
  } catch (_) {
    // If we can't stat either file, play it safe and use the binary
  }
  return true;
}

/** Returns { cmd, args } describing how to invoke the aggregator. */
function resolveAggregatorInvocation(xlsxPath) {
  if (shouldUseBundledBinary()) {
    const bundled = bundledPythonPath();
    console.log("[aggregator] Using bundled binary:", bundled);
    return { cmd: bundled, args: [xlsxPath] };
  }

  const python = resolveSystemPython();
  if (python) {
    console.log("[aggregator] Using system Python:", python, AGGREGATE_SCRIPT);
    return { cmd: python, args: [AGGREGATE_SCRIPT, xlsxPath] };
  }

  throw new Error(
    `Bundled aggregator not found at ${bundledPythonPath()}, and no system Python ` +
      `interpreter was found either (tried: ${SYSTEM_PYTHON_CANDIDATES.join(", ")}). ` +
      `Run "npm run build:python" to build the bundled binary, or install ` +
      `Python 3 + "pip install python-calamine" for development.`,
  );
}

/**
 * Runs the aggregator as a subprocess (streaming, non-blocking) and
 * resolves with the parsed payload.
 */
function runAggregator(xlsxPath, onProgress) {
  return new Promise((resolve, reject) => {
    let cmd, args;
    try {
      ({ cmd, args } = resolveAggregatorInvocation(xlsxPath));
    } catch (err) {
      reject(err);
      return;
    }

    if (onProgress) onProgress({ percent: 0, status: "Starting…" });

    const child = spawn(cmd, args);
    let stdout = "";
    let stderrLineBuf = "";
    let lastErrorLine = "";

    child.stdout.on("data", (chunk) => {
      stdout += chunk;
    });

    child.stderr.on("data", (chunk) => {
      stderrLineBuf += chunk.toString();
      const lines = stderrLineBuf.split("\n");
      stderrLineBuf = lines.pop(); // keep the trailing partial line buffered
      for (const line of lines) {
        if (line.startsWith("PROGRESS:")) {
          const rest = line.slice("PROGRESS:".length);
          const sep = rest.indexOf(":");
          const pctRaw = sep === -1 ? rest : rest.slice(0, sep);
          const status = sep === -1 ? "" : rest.slice(sep + 1);
          const percent = pctRaw === "-" ? null : parseInt(pctRaw, 10);
          if (onProgress) onProgress({ percent, status });
        } else if (line.startsWith("ERROR:")) {
          lastErrorLine = line.slice("ERROR:".length).trim();
        }
      }
    });

    child.on("error", (err) => {
      reject(new Error(`Could not start aggregator (${cmd}): ${err.message}`));
    });

    child.on("close", (code) => {
      if (code === 0) {
        try {
          const payload = JSON.parse(stdout);
          // Diagnostic: log budget counts so DevTools shows them immediately
          const bva = payload.bva || {};
          console.log(
            `[aggregator] rows=${payload.n_rows} invoices=${payload.n_invoices} ` +
              `bva.entries=${bva.entries?.length ?? 0} bva.detail=${bva.detail?.length ?? 0}`,
          );
          resolve(payload);
        } catch (err) {
          reject(
            new Error(`Failed to parse aggregator output: ${err.message}`),
          );
        }
      } else {
        reject(
          new Error(lastErrorLine || `Aggregator exited with code ${code}`),
        );
      }
    });
  });
}

async function buildDashboardHtml(xlsxPath, onProgress) {
  const payload = await runAggregator(xlsxPath, onProgress);
  const template = fs.readFileSync(TEMPLATE_PATH, "utf-8");
  return template.replace("__PAYLOAD__", JSON.stringify(payload));
}

/**
 * Loads and renders a workbook, notifying the given window(s) of progress
 * and of the final ready/error result via IPC.
 */
async function renderFile(filePath, { onProgress, targetWindow } = {}) {
  try {
    const html = await buildDashboardHtml(filePath, onProgress);

    // Cache-bust: write to a unique filename so the renderer never loads
    // a stale cached version from a previous run.
    const outDir = app.getPath("userData");
    const outName = `dashboard-rendered-${Date.now()}.html`;
    const outPath = path.join(outDir, outName);
    fs.writeFileSync(outPath, html, "utf-8");

    // Clean up old rendered files (keep last 10)
    try {
      const files = fs
        .readdirSync(outDir)
        .filter((f) => f.startsWith("dashboard-rendered-") && f.endsWith(".html"))
        .map((f) => ({ name: f, mtime: fs.statSync(path.join(outDir, f)).mtimeMs }))
        .sort((a, b) => b.mtime - a.mtime);
      for (const old of files.slice(10)) {
        fs.unlinkSync(path.join(outDir, old.name));
      }
    } catch (_) {
      // ignore cleanup errors
    }

    currentFilePath = filePath;
    if (targetWindow && !targetWindow.isDestroyed()) {
      targetWindow.webContents.send("dashboard-ready", {
        outPath,
        fileName: path.basename(filePath),
      });
    }
    return outPath;
  } catch (err) {
    dialog.showErrorBox("Could not load workbook", err.message);
    if (targetWindow && !targetWindow.isDestroyed()) {
      targetWindow.webContents.send("dashboard-error", {
        message: err.message,
      });
    }
    return null;
  }
}

async function openFileDialog() {
  const { canceled, filePaths } = await dialog.showOpenDialog(mainWindow, {
    title: "Select GoFresh sales workbook",
    filters: [{ name: "Excel workbooks", extensions: ["xlsx", "xlsm"] }],
    properties: ["openFile"],
  });
  if (canceled || !filePaths.length) return;
  await renderFile(filePaths[0], {
    targetWindow: mainWindow,
    onProgress: (p) =>
      mainWindow &&
      !mainWindow.isDestroyed() &&
      mainWindow.webContents.send("load-progress", p),
  });
}

function buildMenu() {
  const template = [
    {
      label: "File",
      submenu: [
        {
          label: "Open Excel File...",
          accelerator: "CmdOrCtrl+O",
          click: openFileDialog,
        },
        {
          label: "Reload Current File",
          accelerator: "CmdOrCtrl+R",
          click: () =>
            currentFilePath &&
            renderFile(currentFilePath, {
              targetWindow: mainWindow,
              onProgress: (p) =>
                mainWindow && mainWindow.webContents.send("load-progress", p),
            }),
        },
        { type: "separator" },
        { role: "quit" },
      ],
    },
    {
      label: "View",
      submenu: [
        { role: "reload" },
        { role: "toggleDevTools" },
        { type: "separator" },
        { role: "resetZoom" },
        { role: "zoomIn" },
        { role: "zoomOut" },
        { type: "separator" },
        { role: "togglefullscreen" },
      ],
    },
    {
      label: "Help",
      submenu: [{ label: "Check for Updates...", click: checkForUpdates }],
    },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

// ---------------------------------------------------------------------------
// Auto-update
// ---------------------------------------------------------------------------

function sendUpdateStatus(data) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send("update-status", data);
  }
}

let updateReadyToInstall = false;
let installTriggered = false;

function silentInstallAndRelaunch() {
  if (installTriggered) return;
  installTriggered = true;
  autoUpdater.quitAndInstall(true, true);
}

function setupAutoUpdater() {
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = false;

  autoUpdater.on("checking-for-update", () => {
    sendUpdateStatus({ state: "checking" });
  });
  autoUpdater.on("update-available", (info) => {
    sendUpdateStatus({ state: "available", version: info.version });
  });
  autoUpdater.on("update-not-available", () => {
    sendUpdateStatus({ state: "not-available" });
  });
  autoUpdater.on("error", (err) => {
    sendUpdateStatus({ state: "error", message: err.message });
  });
  autoUpdater.on("download-progress", (progress) => {
    sendUpdateStatus({
      state: "downloading",
      percent: Math.round(progress.percent),
    });
  });
  autoUpdater.on("update-downloaded", (info) => {
    updateReadyToInstall = true;
    sendUpdateStatus({ state: "downloaded", version: info.version });
  });
}

app.on("before-quit", (event) => {
  if (updateReadyToInstall && !installTriggered) {
    event.preventDefault();
    silentInstallAndRelaunch();
  }
});

function checkForUpdates() {
  autoUpdater.checkForUpdates().catch((err) => {
    sendUpdateStatus({ state: "error", message: err.message });
  });
}

function createSplashWindow() {
  const win = new BrowserWindow({
    width: 460,
    height: 460,
    frame: false,
    resizable: false,
    movable: true,
    transparent: true,
    alwaysOnTop: true,
    center: true,
    show: false,
    skipTaskbar: true,
    icon: fs.existsSync(APP_ICON) ? APP_ICON : undefined,
    webPreferences: {
      preload: path.join(__dirname, "splash-preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.loadFile(path.join(__dirname, "assets", "splash.html"));
  win.once("ready-to-show", () => win.show());
  return win;
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 900,
    minHeight: 600,
    show: false,
    icon: fs.existsSync(APP_ICON) ? APP_ICON : undefined,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile(path.join(__dirname, "index.html"));
  buildMenu();
  return mainWindow;
}

const SPLASH_MIN_VISIBLE_MS = 700;

async function startup() {
  splashWindow = createSplashWindow();
  const splashShownAt = Date.now();

  createWindow();

  const finishSplash = async () => {
    const elapsed = Date.now() - splashShownAt;
    const remaining = SPLASH_MIN_VISIBLE_MS - elapsed;
    if (remaining > 0) await new Promise((r) => setTimeout(r, remaining));
    if (splashWindow && !splashWindow.isDestroyed()) {
      splashWindow.webContents.send("splash-done");
      setTimeout(() => {
        if (splashWindow && !splashWindow.isDestroyed()) splashWindow.close();
        splashWindow = null;
      }, 260);
    }
    if (mainWindow && !mainWindow.isDestroyed()) mainWindow.show();
  };

  mainWindow.once("ready-to-show", async () => {
    if (fs.existsSync(DEFAULT_XLSX)) {
      await renderFile(DEFAULT_XLSX, {
        targetWindow: mainWindow,
        onProgress: (p) =>
          splashWindow &&
          !splashWindow.isDestroyed() &&
          splashWindow.webContents.send("splash-progress", p),
      });
    }
    await finishSplash();
  });

  if (app.isPackaged) {
    setTimeout(() => {
      checkForUpdates();
    }, 5000);
  }
}

ipcMain.handle("open-file-dialog", openFileDialog);
ipcMain.handle("get-current-file", () => currentFilePath);
ipcMain.handle("load-dropped-file", async (_event, filePath) => {
  if (!filePath || !/\.(xlsx|xlsm)$/i.test(filePath)) {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send("dashboard-error", {
        message: "Please drop a .xlsx or .xlsm file.",
      });
    }
    return;
  }
  await renderFile(filePath, {
    targetWindow: mainWindow,
    onProgress: (p) =>
      mainWindow &&
      !mainWindow.isDestroyed() &&
      mainWindow.webContents.send("load-progress", p),
  });
});

ipcMain.handle("restart-and-install", () => {
  silentInstallAndRelaunch();
});

setupAutoUpdater();

app.whenReady().then(startup);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) startup();
});

setInterval(checkForUpdates, 4 * 60 * 60 * 1000);