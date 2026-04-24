export {};

declare global {
  interface Window {
    pimClient: {
      windowMinimize: () => Promise<void>;
      windowToggleMaximize: () => Promise<boolean>;
      windowClose: () => Promise<void>;
      getAppVersion: () => Promise<string>;
      getUpdateState: () => Promise<{
        phase: string;
        message: string;
        version?: string;
        channel?: string;
        progressPercent?: number;
        canRestartToUpdate: boolean;
        checkedAt?: string;
        nextCheckAt?: string;
      }>;
      checkForUpdates: () => Promise<{
        phase: string;
        message: string;
        version?: string;
        channel?: string;
        progressPercent?: number;
        canRestartToUpdate: boolean;
        checkedAt?: string;
        nextCheckAt?: string;
      }>;
      getUpdateDiagnostics: () => Promise<{
        isPackaged: boolean;
        updatesDisabled: boolean;
        channel: string;
        startupDelayMs: number;
        checkIntervalMs: number;
        warnings: string[];
        metadataPath: string;
        metadataValid: boolean;
        metadataIssue?: string;
      }>;
      restartToInstall: () => Promise<void>;
      onUpdateState: (listener: (state: {
        phase: string;
        message: string;
        version?: string;
        channel?: string;
        progressPercent?: number;
        canRestartToUpdate: boolean;
        checkedAt?: string;
        nextCheckAt?: string;
      }) => void) => () => void;
      login: (tenantKey?: string) => Promise<{ accountId: string; tenantId: string; username: string }>;
      listEligibleRoles: (filter?: { tenantKey?: 'nuance' | 'healthcareCloud'; family?: 'entra' | 'azureResource' | 'group' }) => Promise<
        Array<{
          id: string;
          tenantKey: 'nuance' | 'healthcareCloud';
          tenantLabel: string;
          family: 'entra' | 'azureResource' | 'group';
          displayName: string;
          scope: string;
          membership: string;
          endTime: string;
          resource: string;
          resourceType: string;
          state: 'eligible' | 'active' | 'pending';
        }>
      >;
      activateRole: (request: {
        tenantKey: 'nuance' | 'healthcareCloud';
        family: 'entra' | 'azureResource' | 'group';
        roleId: string;
        durationHours: number;
        justification: string;
      }) => Promise<{ requestId: string; status: string }>;
      deactivateRole: (request: {
        tenantKey: 'nuance' | 'healthcareCloud';
        family: 'entra' | 'azureResource' | 'group';
        roleId: string;
      }) => Promise<{ requestId: string; status: string }>;
      getChangelog: () => Promise<{
        entries: Array<{
          version: string;
          date: string;
          sections: Array<{ heading: string; items: string[] }>;
        }>;
        found: boolean;
        sourcePath: string;
        currentVersion: string;
      }>;
      getLastSeenChangelogVersion: () => Promise<string>;
      markChangelogSeen: (version?: string) => Promise<void>;
      openExternal: (url: string) => Promise<void>;
    };
  }
}
