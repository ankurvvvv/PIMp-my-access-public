import { app, safeStorage } from 'electron';
import fs from 'fs';
import path from 'path';

export type AuthMode = 'managed' | 'byo';

export interface AuthSettings {
  authMode: AuthMode;
  managedClientId: string;
  byoClientId: string;
  byoTenantId: string;
}

const DEFAULT_SETTINGS: AuthSettings = {
  authMode: 'managed',
  managedClientId: process.env.PIMMEFAST_CLIENT_ID ?? '',
  byoClientId: '',
  byoTenantId: 'organizations'
};

interface PersistedAuthSettings {
  payload: string;
  encrypted: boolean;
}

export class AuthSettingsStore {
  private readonly settingsFilePath = path.join(app.getPath('userData'), 'auth-settings.json');

  load(): AuthSettings {
    if (!fs.existsSync(this.settingsFilePath)) {
      return { ...DEFAULT_SETTINGS };
    }

    try {
      const raw = fs.readFileSync(this.settingsFilePath, 'utf8');
      const parsed = JSON.parse(raw) as PersistedAuthSettings;
      const payload = this.decodePayload(parsed.payload, parsed.encrypted);
      const loaded = JSON.parse(payload) as Partial<AuthSettings>;

      return {
        authMode: loaded.authMode === 'byo' ? 'byo' : 'managed',
        managedClientId: String(loaded.managedClientId ?? '').trim() || DEFAULT_SETTINGS.managedClientId,
        byoClientId: loaded.byoClientId ?? '',
        byoTenantId: loaded.byoTenantId ?? 'organizations'
      };
    } catch {
      return { ...DEFAULT_SETTINGS };
    }
  }

  save(input: AuthSettings): AuthSettings {
    const normalized: AuthSettings = {
      authMode: input.authMode,
      managedClientId: input.managedClientId.trim(),
      byoClientId: input.byoClientId.trim(),
      byoTenantId: input.byoTenantId.trim() || 'organizations'
    };

    const payload = JSON.stringify(normalized);
    const encrypted = safeStorage.isEncryptionAvailable();
    const data: PersistedAuthSettings = {
      payload: this.encodePayload(payload, encrypted),
      encrypted
    };

    fs.mkdirSync(path.dirname(this.settingsFilePath), { recursive: true });
    fs.writeFileSync(this.settingsFilePath, JSON.stringify(data, null, 2), 'utf8');
    return normalized;
  }

  validate(input: AuthSettings): { valid: boolean; errors: string[] } {
    const errors: string[] = [];

    if (input.authMode !== 'managed' && input.authMode !== 'byo') {
      errors.push('Auth mode must be managed or byo.');
    }

    if (input.authMode === 'managed' && !this.isGuid(input.managedClientId.trim())) {
      errors.push('Managed mode requires a valid Application (client) ID GUID. Open Settings and set Managed client ID, or set PIMMEFAST_CLIENT_ID.');
    }

    if (input.authMode === 'byo') {
      if (!this.isGuid(input.byoClientId.trim())) {
        errors.push('BYO mode requires a valid Application (client) ID GUID.');
      }

      const tenant = input.byoTenantId.trim();
      const isTenantAccepted = tenant === 'organizations' || this.isGuid(tenant);
      if (!isTenantAccepted) {
        errors.push('BYO tenant must be organizations or a valid tenant GUID.');
      }
    }

    return {
      valid: errors.length === 0,
      errors
    };
  }

  private isGuid(value: string): boolean {
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
  }

  private encodePayload(payload: string, encrypted: boolean): string {
    if (!encrypted) {
      return payload;
    }

    return safeStorage.encryptString(payload).toString('base64');
  }

  private decodePayload(payload: string, encrypted: boolean): string {
    if (!encrypted) {
      return payload;
    }

    return safeStorage.decryptString(Buffer.from(payload, 'base64'));
  }
}
