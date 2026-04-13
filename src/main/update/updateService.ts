import { app, BrowserWindow } from 'electron';
import fs from 'fs';
import path from 'path';
import { autoUpdater, type ProgressInfo, type UpdateDownloadedEvent, type UpdateInfo } from 'electron-updater';

export type UpdatePhase =
  | 'idle'
  | 'disabled'
  | 'checking'
  | 'available'
  | 'downloading'
  | 'downloaded'
  | 'up-to-date'
  | 'error';

export interface UpdateState {
  phase: UpdatePhase;
  message: string;
  version?: string;
  channel?: string;
  progressPercent?: number;
  canRestartToUpdate: boolean;
  checkedAt?: string;
  nextCheckAt?: string;
}

export interface UpdateDiagnostics {
  isPackaged: boolean;
  updatesDisabled: boolean;
  channel: string;
  startupDelayMs: number;
  checkIntervalMs: number;
  warnings: string[];
  metadataPath: string;
  metadataValid: boolean;
  metadataIssue?: string;
}

interface UpdateRuntimeConfig {
  channel: string;
  startupDelayMs: number;
  checkIntervalMs: number;
  disabled: boolean;
}

function parsePositiveInt(value: string | undefined, fallback: number): number {
  const raw = Number.parseInt(String(value ?? ''), 10);
  return Number.isFinite(raw) && raw > 0 ? raw : fallback;
}

function parseDisableFlag(value: string | undefined): boolean {
  const normalized = String(value ?? '').trim().toLowerCase();
  return normalized === '1' || normalized === 'true' || normalized === 'yes';
}

class UpdateService {
  private initialized = false;
  private periodicTimer: NodeJS.Timeout | null = null;
  private checkInFlight: Promise<UpdateState> | null = null;
  private config: UpdateRuntimeConfig;
  private configWarnings: string[];
  private metadataPath: string;
  private metadataIssue: string | null = null;

  private state: UpdateState = {
    phase: 'idle',
    message: 'Update service is idle.',
    canRestartToUpdate: false
  };

  constructor() {
    const configuredChannel = String(process.env.PIMMEFAST_UPDATE_CHANNEL ?? '').trim().toLowerCase();

    this.config = {
      channel: configuredChannel || 'latest',
      startupDelayMs: parsePositiveInt(process.env.PIMMEFAST_UPDATE_STARTUP_DELAY_MS, 12000),
      checkIntervalMs: parsePositiveInt(process.env.PIMMEFAST_UPDATE_CHECK_INTERVAL_MS, 4 * 60 * 60 * 1000),
      disabled: parseDisableFlag(process.env.PIMMEFAST_DISABLE_UPDATES)
    };
    this.configWarnings = this.getConfigWarnings(this.config);
    this.metadataPath = path.join(process.resourcesPath, 'app-update.yml');
  }

  initialize(): void {
    if (this.initialized) {
      return;
    }

    this.initialized = true;

    if (!app.isPackaged) {
      this.setState({
        phase: 'disabled',
        message: 'Updates are available only in packaged builds.',
        canRestartToUpdate: false
      });
      return;
    }

    if (this.config.disabled) {
      this.setState({
        phase: 'disabled',
        message: 'Updates are disabled by configuration.',
        canRestartToUpdate: false
      });
      return;
    }

    for (const warning of this.configWarnings) {
      console.warn(`[updater-config-warning] ${warning}`);
    }

    const metadataError = this.validatePackagedFeedMetadata();
    if (metadataError) {
      this.metadataIssue = metadataError;
      this.setState({
        phase: 'error',
        message: metadataError,
        canRestartToUpdate: false
      });
      return;
    }

    autoUpdater.autoDownload = true;
    autoUpdater.autoInstallOnAppQuit = true;

    autoUpdater.channel = this.config.channel;
    autoUpdater.allowPrerelease = false;
    this.setState({
      channel: this.config.channel,
      message: `Update channel: ${this.config.channel}`,
      canRestartToUpdate: false
    });

    autoUpdater.on('checking-for-update', () => {
      this.setState({
        phase: 'checking',
        message: 'Checking for updates...',
        checkedAt: new Date().toISOString(),
        canRestartToUpdate: false,
        progressPercent: undefined
      });
    });

    autoUpdater.on('update-available', (info: UpdateInfo) => {
      this.setState({
        phase: 'available',
        message: `Update ${info.version} found. Downloading in background...`,
        version: info.version,
        canRestartToUpdate: false,
        progressPercent: 0
      });
    });

    autoUpdater.on('download-progress', (progress: ProgressInfo) => {
      this.setState({
        phase: 'downloading',
        message: `Downloading update... ${progress.percent.toFixed(1)}%`,
        progressPercent: progress.percent,
        canRestartToUpdate: false
      });
    });

    autoUpdater.on('update-downloaded', (event: UpdateDownloadedEvent) => {
      this.setState({
        phase: 'downloaded',
        message: `Update ${event.version} is ready. Restart to install.`,
        version: event.version,
        progressPercent: 100,
        canRestartToUpdate: true
      });
    });

    autoUpdater.on('update-not-available', () => {
      this.setState({
        phase: 'up-to-date',
        message: 'You already have the latest version.',
        progressPercent: undefined,
        canRestartToUpdate: false
      });
    });

    autoUpdater.on('error', (error: unknown) => {
      const raw = this.getErrorMessage(error);
      this.setState({
        phase: 'error',
        message: `Update check failed: ${raw}`,
        canRestartToUpdate: false,
        progressPercent: undefined
      });
    });

    const startupDelay = Math.max(0, this.config.startupDelayMs);
    this.setNextCheckAt(Date.now() + startupDelay);
    setTimeout(() => {
      void this.checkForUpdates();
    }, startupDelay);

    this.periodicTimer = setInterval(() => {
      this.setNextCheckAt(Date.now() + this.config.checkIntervalMs);
      void this.checkForUpdates();
    }, this.config.checkIntervalMs);

    this.setNextCheckAt(Date.now() + this.config.checkIntervalMs);

    app.on('before-quit', () => {
      if (this.periodicTimer) {
        clearInterval(this.periodicTimer);
        this.periodicTimer = null;
      }
    });
  }

  getState(): UpdateState {
    return { ...this.state };
  }

  getDiagnostics(): UpdateDiagnostics {
    return {
      isPackaged: app.isPackaged,
      updatesDisabled: !app.isPackaged || this.config.disabled,
      channel: this.config.channel,
      startupDelayMs: this.config.startupDelayMs,
      checkIntervalMs: this.config.checkIntervalMs,
      warnings: [...this.configWarnings],
      metadataPath: this.metadataPath,
      metadataValid: !this.metadataIssue,
      metadataIssue: this.metadataIssue || undefined
    };
  }

  async checkForUpdates(): Promise<UpdateState> {
    if (this.checkInFlight) {
      return this.checkInFlight;
    }

    if (!app.isPackaged) {
      this.setState({
        phase: 'disabled',
        message: 'Updates are available only in packaged builds.',
        canRestartToUpdate: false
      });
      return this.getState();
    }

    if (this.config.disabled) {
      this.setState({
        phase: 'disabled',
        message: 'Updates are disabled by configuration.',
        canRestartToUpdate: false
      });
      return this.getState();
    }

    this.checkInFlight = (async () => {
      try {
        await autoUpdater.checkForUpdates();
      } catch (error) {
        const raw = this.getErrorMessage(error);
        this.setState({
          phase: 'error',
          message: `Update check failed: ${raw}`,
          canRestartToUpdate: false
        });
      } finally {
        this.setNextCheckAt(Date.now() + this.config.checkIntervalMs);
        this.checkInFlight = null;
      }

      return this.getState();
    })();

    return this.checkInFlight;
  }

  restartToInstall(): void {
    if (!this.state.canRestartToUpdate) {
      return;
    }

    autoUpdater.quitAndInstall(false, true);
  }

  private setState(next: Partial<UpdateState>): void {
    this.state = {
      ...this.state,
      ...next
    };

    for (const window of BrowserWindow.getAllWindows()) {
      window.webContents.send('updater:state', this.state);
    }
  }

  private getErrorMessage(error: unknown): string {
    if (error instanceof Error && error.message) {
      return error.message;
    }

    return String(error || 'Unknown updater error');
  }

  private validatePackagedFeedMetadata(): string | null {
    if (!fs.existsSync(this.metadataPath)) {
      return 'Updater is misconfigured: app-update.yml is missing in packaged resources.';
    }

    let content = '';
    try {
      content = fs.readFileSync(this.metadataPath, 'utf8');
    } catch {
      return 'Updater is misconfigured: unable to read app-update.yml.';
    }

    const hasProvider = /(^|\n)\s*provider\s*:\s*\S+/m.test(content);
    if (!hasProvider) {
      return 'Updater is misconfigured: app-update.yml is missing provider.';
    }

    const hasUrl = /(^|\n)\s*url\s*:\s*\S+/m.test(content);
    const hasGithubOwner = /(^|\n)\s*owner\s*:\s*\S+/m.test(content);
    const hasGithubRepo = /(^|\n)\s*repo\s*:\s*\S+/m.test(content);
    if (!hasUrl && !(hasGithubOwner && hasGithubRepo)) {
      return 'Updater is misconfigured: app-update.yml must define url or owner/repo source.';
    }

    this.metadataIssue = null;

    return null;
  }

  private setNextCheckAt(timestamp: number): void {
    this.setState({
      nextCheckAt: new Date(timestamp).toISOString()
    });
  }

  private getConfigWarnings(config: UpdateRuntimeConfig): string[] {
    const warnings: string[] = [];

    if (config.checkIntervalMs < config.startupDelayMs) {
      warnings.push('Update check interval is lower than startup delay; scheduled checks may begin before the intended first delay window.');
    }

    if (config.checkIntervalMs < 5 * 60 * 1000) {
      warnings.push('Update check interval is very low (< 5 minutes); this can create unnecessary update traffic.');
    }

    return warnings;
  }
}

export const updateService = new UpdateService();
