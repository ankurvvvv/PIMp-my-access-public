import { app, BrowserWindow, ipcMain, shell } from 'electron';
import { PimService } from '../pim/pimService';
import type { ActivateRoleRequest, DeactivateRoleRequest, PimRoleFilter } from '../pim/types';
import { updateService } from '../update/updateService';
import { readChangelog } from '../changelog/changelogReader';
import { ChangelogStateStore } from '../settings/changelogStateStore';

const pimService = new PimService();
const changelogStateStore = new ChangelogStateStore();

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

  ipcMain.handle('pim:deactivateRole', async (_event, request: DeactivateRoleRequest) => {
    return pimService.deactivateRole(request);
  });

  // ── Changelog / "What's new" ──
  // Renderer fetches the parsed CHANGELOG.md to render in the in-app dialog,
  // and reads/writes the `lastSeenVersion` so we can auto-open the dialog
  // exactly once after each version bump.

  ipcMain.handle('changelog:get', async () => {
    const payload = readChangelog();
    return {
      ...payload,
      currentVersion: app.getVersion()
    };
  });

  ipcMain.handle('changelog:getLastSeenVersion', async () => {
    const state = changelogStateStore.load();
    return state.lastSeenVersion;
  });

  ipcMain.handle('changelog:markSeen', async (_event, version?: string) => {
    const versionToStore = typeof version === 'string' && version.length > 0
      ? version
      : app.getVersion();
    changelogStateStore.save({ lastSeenVersion: versionToStore });
  });

  // Open external URLs (e.g. Keep-A-Changelog reference) safely in the user's
  // default browser. Renderer cannot call shell.openExternal directly because
  // the renderer is sandboxed (contextIsolation: true, nodeIntegration: false).
  ipcMain.handle('shell:openExternal', async (_event, url: string) => {
    if (typeof url !== 'string') {
      return;
    }
    // Allowlist: only http/https. Prevents file:// or javascript: scheme abuse
    // if a malicious changelog ever sneaks in.
    if (!/^https?:\/\//i.test(url)) {
      return;
    }
    await shell.openExternal(url);
  });
}
