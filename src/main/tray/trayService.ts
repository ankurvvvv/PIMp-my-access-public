import { app, BrowserWindow, Menu, Tray, nativeImage } from 'electron';
import path from 'path';
import { updateService } from '../update/updateService';

let tray: Tray | null = null;

function resolveIconPath(): string {
  return path.join(app.getAppPath(), 'assets', 'icons', 'app.png');
}

function showMainWindow(): void {
  const windows = BrowserWindow.getAllWindows();
  if (windows.length === 0) {
    return;
  }

  const mainWindow = windows[0];
  if (mainWindow.isMinimized()) {
    mainWindow.restore();
  }

  mainWindow.show();
  mainWindow.focus();
}

function buildContextMenu(): Menu {
  const updateState = updateService.getState();
  const hasUpdate = updateState.phase === 'downloaded' && updateState.canRestartToUpdate;

  const template: Electron.MenuItemConstructorOptions[] = [
    {
      label: 'Show PIMp my access',
      click: showMainWindow
    },
    { type: 'separator' }
  ];

  if (hasUpdate) {
    template.push({
      label: `Install update ${updateState.version || ''}`.trim(),
      click: () => updateService.restartToInstall()
    });
    template.push({ type: 'separator' });
  }

  template.push({
    label: `v${app.getVersion()}`,
    enabled: false
  });

  template.push({
    label: 'Quit',
    click: () => app.quit()
  });

  return Menu.buildFromTemplate(template);
}

export function initializeTray(): void {
  if (tray) {
    return;
  }

  const iconPath = resolveIconPath();
  const icon = nativeImage.createFromPath(iconPath).resize({ width: 16, height: 16 });
  tray = new Tray(icon);
  tray.setToolTip(`PIMp my access v${app.getVersion()}`);
  tray.setContextMenu(buildContextMenu());

  tray.on('click', showMainWindow);

  app.on('before-quit', () => {
    if (tray) {
      tray.destroy();
      tray = null;
    }
  });
}

export function refreshTrayMenu(): void {
  if (!tray) {
    return;
  }

  tray.setContextMenu(buildContextMenu());
}
