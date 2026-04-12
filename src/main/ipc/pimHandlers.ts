import { app, BrowserWindow, ipcMain } from 'electron';
import { PimService } from '../pim/pimService';
import type { ActivateRoleRequest, PimRoleFilter } from '../pim/types';
import { updateService } from '../update/updateService';

const pimService = new PimService();

export function registerPimHandlers(): void {
  ipcMain.handle('window:minimize', async (event) => {
    const window = BrowserWindow.fromWebContents(event.sender);
    window?.minimize();
  });

  ipcMain.handle('window:toggleMaximize', async (event) => {
    const window = BrowserWindow.fromWebContents(event.sender);
    if (!window) {
      return false;
    }

    if (window.isMaximized()) {
      window.unmaximize();
      return false;
    }

    window.maximize();
    return true;
  });

  ipcMain.handle('window:close', async (event) => {
    const window = BrowserWindow.fromWebContents(event.sender);
    window?.close();
  });

  ipcMain.handle('app:getVersion', async () => {
    return app.getVersion();
  });

  ipcMain.handle('updater:getState', async () => {
    return updateService.getState();
  });

  ipcMain.handle('updater:checkForUpdates', async () => {
    return updateService.checkForUpdates();
  });

  ipcMain.handle('updater:getDiagnostics', async () => {
    return updateService.getDiagnostics();
  });

  ipcMain.handle('updater:restartToInstall', async () => {
    updateService.restartToInstall();
  });

  // NOTE: Auth settings endpoints (managed/BYO/tenant override) are intentionally
  // disabled for now to enforce one-click Azure sign-in UX.
  // Keep PimService auth settings methods in place for future public-release hardening.

  ipcMain.handle('auth:login', async (_event, tenantKey?: string) => {
    return pimService.login(tenantKey);
  });

  ipcMain.handle('pim:listEligibleRoles', async (_event, filter?: PimRoleFilter) => {
    return pimService.listEligibleRoles(filter);
  });

  ipcMain.handle('pim:activateRole', async (_event, request: ActivateRoleRequest) => {
    return pimService.activateRole(request);
  });
}
