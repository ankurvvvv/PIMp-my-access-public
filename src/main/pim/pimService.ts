import { DeviceCodeAuth } from '../auth/deviceCodeAuth';
import { ArmClient } from '../api/armClient';
import { GraphClient } from '../api/graphClient';
import { AuthSettingsStore } from '../settings/authSettingsStore';
import { BUNDLED_CLIENT_ID, TENANT_KEY_TO_ID, ALL_TENANTS, toTenantKey } from './tenantConfig';
import { listEntraEligibleRoles, listGroupEligibleRoles, listAzureResourceEligibleRoles } from './roleListing';
import { activateRole as performActivation, deactivateRole as performDeactivation } from './activation';
import type {
  ActivateRoleRequest,
  AuthLoginResult,
  AuthSettings,
  AuthSettingsValidationResult,
  DeactivateRoleRequest,
  PimRole,
  PimRoleFilter,
  PimTenantKey
} from './types';

export class PimService {
  private readonly auth = new DeviceCodeAuth();
  private readonly authSettingsStore = new AuthSettingsStore();

  async login(tenantKey?: string): Promise<AuthLoginResult> {
    const settings = this.authSettingsStore.load();
    const managedClientId = [
      process.env.PIMMEFAST_CLIENT_ID,
      settings.managedClientId,
      BUNDLED_CLIENT_ID
    ]
      .map((value) => String(value ?? '').trim())
      .find((value) => value.length > 0) ?? '';

    if (!managedClientId) {
      throw new Error('Application (client) ID is not configured. Set BUNDLED_CLIENT_ID in tenantConfig.ts or set PIMMEFAST_CLIENT_ID.');
    }

    const resolvedTenantKey = toTenantKey(tenantKey);
    const selectedTenant = resolvedTenantKey ? TENANT_KEY_TO_ID[resolvedTenantKey] : '';
    const resolvedTenant = selectedTenant || 'organizations';

    return this.auth.login(managedClientId, resolvedTenant);
  }

  getAuthSettings(): AuthSettings {
    return this.authSettingsStore.load();
  }

  saveAuthSettings(input: AuthSettings): AuthSettings {
    const validation = this.authSettingsStore.validate(input);
    if (!validation.valid) {
      throw new Error(validation.errors.join(' '));
    }

    return this.authSettingsStore.save(input);
  }

  validateAuthSettings(input: AuthSettings): AuthSettingsValidationResult {
    return this.authSettingsStore.validate(input);
  }

  async listEligibleRoles(filter?: PimRoleFilter): Promise<PimRole[]> {
    if (filter?.tenantKey) {
      return this.listEligibleRolesForTenant(filter.tenantKey, filter.family);
    }

    const allTenantRoles = await Promise.allSettled(
      ALL_TENANTS.map((tenantKey) => this.listEligibleRolesForTenant(tenantKey, filter?.family))
    );

    const roles: PimRole[] = [];
    for (const result of allTenantRoles) {
      if (result.status === 'fulfilled') {
        roles.push(...result.value);
      }
    }

    if (roles.length > 0) {
      return roles;
    }

    const reasons = allTenantRoles
      .filter((result): result is PromiseRejectedResult => result.status === 'rejected')
      .map((result) => String(result.reason));
    throw new Error(reasons.join(' | '));
  }

  async activateRole(request: ActivateRoleRequest): Promise<{ requestId: string; status: string }> {
    const tenantKey = toTenantKey(request.tenantKey) ?? 'nuance';
    const graph = this.getGraphClientForTenant(tenantKey);
    const arm = this.getArmClientForTenant(tenantKey);

    return performActivation(request, graph, arm, this.auth, tenantKey);
  }

  // Self-deactivate an active PIM role. Mirrors activateRole's tenant routing.
  async deactivateRole(request: DeactivateRoleRequest): Promise<{ requestId: string; status: string }> {
    const tenantKey = toTenantKey(request.tenantKey) ?? 'nuance';
    const graph = this.getGraphClientForTenant(tenantKey);
    const arm = this.getArmClientForTenant(tenantKey);

    return performDeactivation(request, graph, arm, this.auth, tenantKey);
  }

  private async listEligibleRolesForTenant(tenantKey: PimTenantKey, family?: PimRoleFilter['family']): Promise<PimRole[]> {
    const graph = this.getGraphClientForTenant(tenantKey);
    const arm = this.getArmClientForTenant(tenantKey);

    if (family === 'entra') {
      return listEntraEligibleRoles(tenantKey, graph);
    }

    if (family === 'group') {
      return listGroupEligibleRoles(tenantKey, graph);
    }

    if (family === 'azureResource') {
      return listAzureResourceEligibleRoles(tenantKey, arm);
    }

    const [entra, groups, azureResources] = await Promise.allSettled([
      listEntraEligibleRoles(tenantKey, graph),
      listGroupEligibleRoles(tenantKey, graph),
      listAzureResourceEligibleRoles(tenantKey, arm)
    ]);

    const roles: PimRole[] = [];
    if (entra.status === 'fulfilled') {
      roles.push(...entra.value);
    }
    if (groups.status === 'fulfilled') {
      roles.push(...groups.value);
    }
    if (azureResources.status === 'fulfilled') {
      roles.push(...azureResources.value);
    }

    if (roles.length === 0) {
      const reasons = [entra, groups, azureResources]
        .filter((result): result is PromiseRejectedResult => result.status === 'rejected')
        .map((result) => String(result.reason));
      throw new Error(reasons.join(' | '));
    }

    return roles;
  }

  private getGraphClientForTenant(tenantKey: PimTenantKey): GraphClient {
    const tenantId = TENANT_KEY_TO_ID[tenantKey];
    return new GraphClient(() => this.auth.getGraphTokenForTenant(tenantId));
  }

  private getArmClientForTenant(tenantKey: PimTenantKey): ArmClient {
    const tenantId = TENANT_KEY_TO_ID[tenantKey];
    return new ArmClient(() => this.auth.getArmTokenForTenant(tenantId));
  }
}
