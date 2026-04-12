import { contextBridge, ipcRenderer } from 'electron';
import type {
  ActivateRoleRequest,
  AuthLoginResult,
  PimRole,
  PimRoleFilter
} from './pim/types';

const api = {
  windowMinimize: (): Promise<void> =>
    ipcRenderer.invoke('window:minimize'),

  windowToggleMaximize: (): Promise<boolean> =>
    ipcRenderer.invoke('window:toggleMaximize'),

  windowClose: (): Promise<void> =>
    ipcRenderer.invoke('window:close'),

  getAppVersion: (): Promise<string> =>
    ipcRenderer.invoke('app:getVersion'),

  getUpdateState: (): Promise<{
    phase: string;
    message: string;
    version?: string;
    channel?: string;
    progressPercent?: number;
    canRestartToUpdate: boolean;
    checkedAt?: string;
    nextCheckAt?: string;
  }> =>
    ipcRenderer.invoke('updater:getState'),

  checkForUpdates: (): Promise<{
    phase: string;
    message: string;
    version?: string;
    channel?: string;
    progressPercent?: number;
    canRestartToUpdate: boolean;
    checkedAt?: string;
    nextCheckAt?: string;
  }> =>
    ipcRenderer.invoke('updater:checkForUpdates'),

  getUpdateDiagnostics: (): Promise<{
    isPackaged: boolean;
    updatesDisabled: boolean;
    channel: string;
    startupDelayMs: number;
    checkIntervalMs: number;
    warnings: string[];
    metadataPath: string;
    metadataValid: boolean;
    metadataIssue?: string;
  }> =>
    ipcRenderer.invoke('updater:getDiagnostics'),

  restartToInstall: (): Promise<void> =>
    ipcRenderer.invoke('updater:restartToInstall'),

  onUpdateState: (listener: (state: {
    phase: string;
    message: string;
    version?: string;
    channel?: string;
    progressPercent?: number;
    canRestartToUpdate: boolean;
    checkedAt?: string;
    nextCheckAt?: string;
  }) => void): (() => void) => {
    const wrapped = (_event: Electron.IpcRendererEvent, state: {
      phase: string;
      message: string;
      version?: string;
      channel?: string;
      progressPercent?: number;
      canRestartToUpdate: boolean;
      checkedAt?: string;
      nextCheckAt?: string;
    }) => {
      listener(state);
    };

    ipcRenderer.on('updater:state', wrapped);
    return () => {
      ipcRenderer.removeListener('updater:state', wrapped);
    };
  },

  // NOTE: BYO/managed/tenant-override settings API is intentionally hidden from renderer for now.
  login: (tenantKey?: string): Promise<AuthLoginResult> =>
    ipcRenderer.invoke('auth:login', tenantKey),

  listEligibleRoles: (filter?: PimRoleFilter): Promise<PimRole[]> =>
    ipcRenderer.invoke('pim:listEligibleRoles', filter),

  activateRole: (request: ActivateRoleRequest): Promise<{ requestId: string; status: string }> =>
    ipcRenderer.invoke('pim:activateRole', request)
};

try {
  contextBridge.exposeInMainWorld('pimClient', api);
} catch (error) {
  console.error('Failed to expose preload bridge', error);
}
