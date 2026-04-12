export type PimRoleFamily = 'entra' | 'azureResource' | 'group';
export type PimTenantKey = 'nuance' | 'healthcareCloud';
export type AuthMode = 'managed' | 'byo';

export interface AuthSettings {
  authMode: AuthMode;
  managedClientId: string;
  byoClientId: string;
  byoTenantId: string;
}

export interface AuthSettingsValidationResult {
  valid: boolean;
  errors: string[];
}

export interface AuthLoginResult {
  accountId: string;
  tenantId: string;
  username: string;
}

export interface PimRole {
  id: string;
  tenantKey: PimTenantKey;
  tenantLabel: string;
  family: PimRoleFamily;
  displayName: string;
  scope: string;
  membership: string;
  endTime: string;
  endTimeIso?: string;
  resource: string;
  resourceType: string;
  state: 'eligible' | 'active' | 'pending';
  source: Record<string, string>;
}

export interface PimRoleFilter {
  tenantKey?: PimTenantKey;
  family?: PimRoleFamily;
}

export interface ActivateRoleRequest {
  tenantKey: PimTenantKey;
  family: PimRoleFamily;
  roleId: string;
  durationHours: number;
  justification: string;
}

export interface GraphActivationResult {
  id: string;
  status: string;
}

export interface ArmActivationResult {
  name: string;
  properties?: {
    status?: string;
  };
}
