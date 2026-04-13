import type { PimTenantKey } from './types';

// TODO(public-release): Move these runtime auth constants into a secure, versioned config model.
export const BUNDLED_CLIENT_ID = '77e5ef01-979f-4959-8737-2df0d3f7a9b0';

export const TENANT_KEY_TO_ID: Record<PimTenantKey, string> = {
  nuance: '29208c38-8fc5-4a03-89e2-9b6e8e4b388b',
  healthcareCloud: 'ed5693bc-117f-4001-a14c-5f50a530d5df'
};

export const TENANT_KEY_TO_LABEL: Record<PimTenantKey, string> = {
  nuance: 'Nuance',
  healthcareCloud: 'HealthCare Cloud'
};

export const ALL_TENANTS: PimTenantKey[] = ['nuance', 'healthcareCloud'];

export function toTenantKey(value?: string): PimTenantKey | undefined {
  if (value === 'nuance' || value === 'healthcareCloud') {
    return value;
  }

  return undefined;
}
