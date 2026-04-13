import { app, BrowserWindow } from 'electron';
import path from 'path';
import { registerPimHandlers } from './ipc/pimHandlers';
import { initializeTray, refreshTrayMenu } from './tray/trayService';
import { updateService } from './update/updateService';

let mainWindow: BrowserWindow | null = null;

function resolveRendererEntry(): string {
  if (app.isPackaged) {
    return path.join(app.getAppPath(), 'dist', 'renderer', 'index.html');
  }

  return path.join(app.getAppPath(), 'src', 'renderer', 'index.html');
}

function createWindow(): void {
  const appVersionTag = app.getVersion();
  const iconPath = path.join(app.getAppPath(), 'assets', 'icons', 'app.ico');
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 980,
    minHeight: 640,
    title: `PIMp MY ACCESS | ${appVersionTag} | Ankur Vishwakarma (Microsoft)`,
    icon: iconPath,
    frame: false,
    roundedCorners: false,
    titleBarStyle: 'hidden',
    backgroundColor: '#0f1726',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  const entry = resolveRendererEntry();
  mainWindow.loadFile(entry).catch((error) => {
    console.error('Unable to load renderer entry', { entry, error });
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(() => {
  registerPimHandlers();
  createWindow();
  initializeTray();
  updateService.initialize();
  updateService.onStateChange(() => refreshTrayMenu());

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
