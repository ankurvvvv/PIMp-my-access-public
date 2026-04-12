import { randomUUID } from 'crypto';
import { DeviceCodeAuth } from '../auth/deviceCodeAuth';
import { ArmClient } from '../api/armClient';
import { GraphClient } from '../api/graphClient';
import { AuthSettingsStore } from '../settings/authSettingsStore';
import type {
  ActivateRoleRequest,
  ArmActivationResult,
  AuthLoginResult,
  AuthSettings,
  AuthSettingsValidationResult,
  GraphActivationResult,
  PimRole,
  PimRoleFilter,
  PimTenantKey
} from './types';

interface GraphCollection<T> {
  value: T[];
}

interface GraphEntraEligibility {
  id: string;
  principalId?: string;
  directoryScopeId?: string;
  endDateTime?: string;
  roleDefinition?: {
    displayName?: string;
  };
  roleDefinitionId?: string;
}

interface GraphGroupEligibility {
  id: string;
  principalId?: string;
  groupId?: string;
  accessId?: 'owner' | 'member' | 'unknownFutureValue';
  endDateTime?: string;
  group?: {
    displayName?: string;
  };
}

interface ArmSubscription {
  subscriptionId: string;
  displayName: string;
}

interface ArmRoleEligibility {
  id: string;
  name: string;
  properties?: {
    principalId?: string;
    roleDefinitionId?: string;
    roleEligibilityScheduleId?: string;
    linkedRoleEligibilityScheduleId?: string;
    linkedRoleEligibilityScheduleInstanceId?: string;
    scope?: string;
    status?: string;
    assignmentType?: string;
    endDateTime?: string;
    expandedProperties?: {
      roleDefinition?: {
        displayName?: string;
      };
    };
  };
}

interface ArmCollection<T> {
  value: T[];
}

// TODO(public-release): Move these runtime auth constants into a secure, versioned config model.
const BUNDLED_CLIENT_ID = '77e5ef01-979f-4959-8737-2df0d3f7a9b0';

const TENANT_KEY_TO_ID: Record<PimTenantKey, string> = {
  nuance: '29208c38-8fc5-4a03-89e2-9b6e8e4b388b',
  healthcareCloud: 'ed5693bc-117f-4001-a14c-5f50a530d5df'
};

const TENANT_KEY_TO_LABEL: Record<PimTenantKey, string> = {
  nuance: 'Nuance',
  healthcareCloud: 'HealthCare Cloud'
};

const ALL_TENANTS: PimTenantKey[] = ['nuance', 'healthcareCloud'];

export class PimService {
  private readonly auth = new DeviceCodeAuth();
  private readonly authSettingsStore = new AuthSettingsStore();
  private azureRolesCache: Partial<Record<PimTenantKey, { expiresAt: number; roles: PimRole[] }>> = {};

  private readonly knownAzureRoleNames: Record<string, string> = {
    'b24988ac-6180-42a0-ab88-20f7382dd24c': 'Contributor',
    '8e3af657-a8ff-443c-a75c-2fe8c4bcb635': 'Owner',
    'acdd72a7-3385-48ef-bd42-f606fba81ae7': 'Reader',
    'f58310d9-a9f6-439a-9e8d-f62e7b41a168': 'User Access Administrator'
  };

  async login(tenantKey?: string): Promise<AuthLoginResult> {

    // NOTE: Managed/BYO/tenant-override auth variants are intentionally disabled.
    // Current behavior is one-click Azure sign-in against organizations tenant only.
    const settings = this.authSettingsStore.load();
    const managedClientId = [
      process.env.PIMMEFAST_CLIENT_ID,
      settings.managedClientId,
      BUNDLED_CLIENT_ID
    ]
      .map((value) => String(value ?? '').trim())
      .find((value) => value.length > 0) ?? '';

    if (!managedClientId) {
      throw new Error('Application (client) ID is not configured. Set BUNDLED_CLIENT_ID in src/main/pim/pimService.ts or set PIMMEFAST_CLIENT_ID.');
    }

    const resolvedTenantKey = this.toTenantKey(tenantKey);
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

  private async listEligibleRolesForTenant(tenantKey: PimTenantKey, family?: PimRoleFilter['family']): Promise<PimRole[]> {
    const graph = this.getGraphClientForTenant(tenantKey);
    const arm = this.getArmClientForTenant(tenantKey);

    if (family === 'entra') {
      return this.listEntraEligibleRoles(tenantKey, graph);
    }

    if (family === 'group') {
      return this.listGroupEligibleRoles(tenantKey, graph);
    }

    if (family === 'azureResource') {
      return this.listAzureResourceEligibleRoles(tenantKey, arm);
    }

    const [entra, groups, azureResources] = await Promise.allSettled([
      this.listEntraEligibleRoles(tenantKey, graph),
      this.listGroupEligibleRoles(tenantKey, graph),
      this.listAzureResourceEligibleRoles(tenantKey, arm)
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

  async activateRole(request: ActivateRoleRequest): Promise<{ requestId: string; status: string }> {
    const tenantKey = this.toTenantKey(request.tenantKey) ?? 'nuance';
    const graph = this.getGraphClientForTenant(tenantKey);
    const arm = this.getArmClientForTenant(tenantKey);

    if (request.family === 'entra') {
      const [principalId, roleDefinitionId, directoryScopeId] = request.roleId.split('|');
      const result = await graph.post<GraphActivationResult>('/roleManagement/directory/roleAssignmentScheduleRequests', {
        action: 'selfActivate',
        principalId,
        roleDefinitionId,
        directoryScopeId: directoryScopeId || '/',
        justification: request.justification,
        scheduleInfo: {
          expiration: {
            type: 'AfterDuration',
            duration: `PT${request.durationHours}H`
          }
        }
      });

      return { requestId: result.id, status: result.status };
    }

    if (request.family === 'group') {
      const [principalId, groupId, accessId] = request.roleId.split('|');
      const result = await graph.post<GraphActivationResult>('/identityGovernance/privilegedAccess/group/assignmentScheduleRequests', {
        action: 'selfActivate',
        principalId,
        groupId,
        accessId,
        justification: request.justification,
        scheduleInfo: {
          expiration: {
            type: 'afterDuration',
            duration: `PT${request.durationHours}H`
          }
        }
      });

      return { requestId: result.id, status: result.status };
    }

    const azureTarget = this.parseAzureActivationTarget(request.roleId);
    const { scope } = azureTarget;

    if (!scope.startsWith('/')) {
      throw new Error('Invalid Azure role scope. Refresh the role list and retry.');
    }

    const resolved = await this.resolveAzureActivationTarget(arm, azureTarget);
    const requestorPrincipalId = await this.getArmRequestorPrincipalId(tenantKey);

    if (!resolved.roleDefinitionId) {
      throw new Error('Missing Azure role definition ID. Refresh the role list and retry.');
    }

    const requestId = randomUUID();
    let result: ArmActivationResult;
    try {
      result = await arm.put<ArmActivationResult>(
        `${scope}/providers/Microsoft.Authorization/roleAssignmentScheduleRequests/${requestId}?api-version=2020-10-01`,
        {
          properties: {
            requestType: 'SelfActivate',
            principalId: requestorPrincipalId,
            roleDefinitionId: resolved.roleDefinitionId,
            justification: request.justification,
            scheduleInfo: {
              startDateTime: new Date().toISOString(),
              expiration: {
                type: 'AfterDuration',
                duration: `PT${request.durationHours}H`
              }
            }
          }
        }
      );
    } catch (error) {
      const raw = String(error ?? '');
      if (raw.includes('"code":"RoleAssignmentExists"')) {
        throw new Error(`PIM role is already activated.\n${raw}`);
      }

      throw error;
    }

    this.azureRolesCache[tenantKey] = undefined;
    return { requestId: result.name ?? requestId, status: result.properties?.status ?? 'submitted' };
  }

  private async listEntraEligibleRoles(tenantKey: PimTenantKey, graph: GraphClient): Promise<PimRole[]> {
    const [eligibleResult, activeResult] = await Promise.allSettled([
      graph.get<GraphCollection<GraphEntraEligibility>>(
        "/roleManagement/directory/roleEligibilityScheduleInstances/filterByCurrentUser(on='principal')?$expand=roleDefinition"
      ),
      graph.get<GraphCollection<GraphEntraEligibility>>(
        "/roleManagement/directory/roleAssignmentScheduleInstances/filterByCurrentUser(on='principal')?$expand=roleDefinition"
      )
    ]);

    const roles = new Map<string, PimRole>();

    if (eligibleResult.status === 'fulfilled') {
      for (const entry of eligibleResult.value.value) {
        const role: PimRole = {
          id: `${entry.principalId ?? ''}|${entry.roleDefinitionId ?? ''}|${entry.directoryScopeId ?? '/'}`,
          tenantKey,
          tenantLabel: TENANT_KEY_TO_LABEL[tenantKey],
          family: 'entra',
          displayName: entry.roleDefinition?.displayName ?? 'Entra Role',
          scope: entry.directoryScopeId ?? '/',
          membership: 'Direct',
          endTime: 'Permanent',
          resource: entry.directoryScopeId === '/' || !entry.directoryScopeId ? 'Directory' : entry.directoryScopeId,
          resourceType: 'Directory',
          state: 'eligible',
          source: {
            scheduleInstanceId: entry.id,
            principalId: entry.principalId ?? '',
            roleDefinitionId: entry.roleDefinitionId ?? ''
          }
        };

        roles.set(role.id, role);
      }
    }

    if (activeResult.status === 'fulfilled') {
      for (const entry of activeResult.value.value) {
        const roleId = `${entry.principalId ?? ''}|${entry.roleDefinitionId ?? ''}|${entry.directoryScopeId ?? '/'}`;
        const existing = roles.get(roleId);
        if (!existing) {
          continue;
        }
        const endTimeIso = entry.endDateTime;

        roles.set(roleId, {
          id: roleId,
          tenantKey,
          tenantLabel: TENANT_KEY_TO_LABEL[tenantKey],
          family: 'entra',
          displayName: entry.roleDefinition?.displayName ?? existing.displayName,
          scope: entry.directoryScopeId ?? existing.scope,
          membership: existing.membership,
          endTime: endTimeIso ? new Date(endTimeIso).toLocaleString() : existing.endTime,
          endTimeIso: endTimeIso ?? existing.endTimeIso,
          resource:
            (entry.directoryScopeId === '/' || !entry.directoryScopeId ? 'Directory' : entry.directoryScopeId) ??
            existing.resource ??
            'Directory',
          resourceType: existing.resourceType,
          state: 'active',
          source: {
            scheduleInstanceId: entry.id,
            principalId: entry.principalId ?? existing.source.principalId ?? '',
            roleDefinitionId: entry.roleDefinitionId ?? existing.source.roleDefinitionId ?? ''
          }
        });
      }
    }

    return Array.from(roles.values());
  }

  private async listGroupEligibleRoles(tenantKey: PimTenantKey, graph: GraphClient): Promise<PimRole[]> {
    const [eligibleResult, activeResult] = await Promise.allSettled([
      graph.get<GraphCollection<GraphGroupEligibility>>(
        "/identityGovernance/privilegedAccess/group/eligibilityScheduleInstances/filterByCurrentUser(on='principal')?$expand=group"
      ),
      graph.get<GraphCollection<GraphGroupEligibility>>(
        "/identityGovernance/privilegedAccess/group/assignmentScheduleInstances/filterByCurrentUser(on='principal')?$expand=group"
      )
    ]);

    const roles = new Map<string, PimRole>();

    if (eligibleResult.status === 'fulfilled') {
      for (const entry of eligibleResult.value.value) {
        const role: PimRole = {
          id: `${entry.principalId ?? ''}|${entry.groupId ?? ''}|${entry.accessId ?? 'member'}`,
          tenantKey,
          tenantLabel: TENANT_KEY_TO_LABEL[tenantKey],
          family: 'group',
          displayName: entry.group?.displayName ? `Group: ${entry.group.displayName}` : 'PIM Group',
          scope: entry.groupId ?? '-',
          membership: entry.accessId === 'owner' ? 'Owner' : 'Member',
          endTime: 'Permanent',
          resource: entry.group?.displayName ?? entry.groupId ?? '-',
          resourceType: 'Group',
          state: 'eligible',
          source: {
            assignmentInstanceId: entry.id,
            principalId: entry.principalId ?? '',
            groupId: entry.groupId ?? ''
          }
        };

        roles.set(role.id, role);
      }
    }

    if (activeResult.status === 'fulfilled') {
      for (const entry of activeResult.value.value) {
        const roleId = `${entry.principalId ?? ''}|${entry.groupId ?? ''}|${entry.accessId ?? 'member'}`;
        const existing = roles.get(roleId);
        if (!existing) {
          continue;
        }
        const endTimeIso = entry.endDateTime;

        roles.set(roleId, {
          id: roleId,
          tenantKey,
          tenantLabel: TENANT_KEY_TO_LABEL[tenantKey],
          family: 'group',
          displayName:
            (entry.group?.displayName ? `Group: ${entry.group.displayName}` : undefined) ?? existing.displayName,
          scope: entry.groupId ?? existing.scope,
          membership: entry.accessId === 'owner' ? 'Owner' : existing.membership,
          endTime: endTimeIso ? new Date(endTimeIso).toLocaleString() : existing.endTime,
          endTimeIso: endTimeIso ?? existing.endTimeIso,
          resource: entry.group?.displayName ?? existing.resource ?? entry.groupId ?? '-',
          resourceType: existing.resourceType,
          state: 'active',
          source: {
            assignmentInstanceId: entry.id,
            principalId: entry.principalId ?? existing.source.principalId ?? '',
            groupId: entry.groupId ?? existing.source.groupId ?? ''
          }
        });
      }
    }

    return Array.from(roles.values());
  }

  private async listAzureResourceEligibleRoles(tenantKey: PimTenantKey, arm: ArmClient): Promise<PimRole[]> {
    const now = Date.now();
    const cached = this.azureRolesCache[tenantKey];
    if (cached && cached.expiresAt > now) {
      return cached.roles;
    }

    const subscriptionsResponse = await arm.get<ArmCollection<ArmSubscription>>('/subscriptions?api-version=2020-01-01');
    const roles: PimRole[] = [];
    const bySubscription = new Map(subscriptionsResponse.value.map((sub) => [sub.subscriptionId, sub.displayName]));
    const subscriptionIds = subscriptionsResponse.value.map((sub) => sub.subscriptionId);

    const results = await this.mapWithConcurrency(subscriptionIds, 10, async (subscriptionId) => {
      const [eligibilityResponse, assignmentResponse] = await Promise.allSettled([
        arm.get<ArmCollection<ArmRoleEligibility>>(
          `/subscriptions/${subscriptionId}/providers/Microsoft.Authorization/roleEligibilityScheduleInstances?$filter=asTarget()&api-version=2020-10-01`
        ),
        arm.get<ArmCollection<ArmRoleEligibility>>(
          `/subscriptions/${subscriptionId}/providers/Microsoft.Authorization/roleAssignmentScheduleInstances?$filter=asTarget()&api-version=2020-10-01`
        )
      ]);

      const subscriptionLabel = bySubscription.get(subscriptionId) ?? subscriptionId;
      const roleMap = new Map<string, PimRole>();
      const eligibleByScheduleId = new Map<string, PimRole>();
      const eligibleBySemanticWithPrincipal = new Map<string, PimRole>();
      const eligibleBySemanticWithoutPrincipal = new Map<string, PimRole>();

      if (eligibilityResponse.status === 'fulfilled') {
        for (const entry of eligibilityResponse.value.value) {
          const scope = entry.properties?.scope ?? `/subscriptions/${subscriptionId}`;
          const parsedScope = this.parseAzureScope(scope, subscriptionLabel);
          const roleDefinitionId = entry.properties?.roleDefinitionId ?? '';
          const principalId = entry.properties?.principalId ?? '';
          const normalizedScheduleId = this.normalizeScheduleId(entry.properties?.roleEligibilityScheduleId || entry.id || entry.name);
          const roleGuid = this.extractRoleGuid(roleDefinitionId);
          const roleName =
            entry.properties?.expandedProperties?.roleDefinition?.displayName ||
            (roleGuid ? this.knownAzureRoleNames[roleGuid] : undefined) ||
            'Azure Role';
          const linkedRoleEligibilityScheduleId = normalizedScheduleId || entry.name;
          const encodedTarget = encodeURIComponent(
            JSON.stringify({
              scope,
              linkedRoleEligibilityScheduleId,
              principalId,
              roleDefinitionId
            })
          );
          const roleKey = `${principalId}|${roleDefinitionId}|${scope}`;
          const semanticWithPrincipal = `${principalId}|${roleDefinitionId}|${scope}`;
          const semanticWithoutPrincipal = `${roleDefinitionId}|${scope}`;
          const role: PimRole = {
            id: encodedTarget,
            tenantKey,
            tenantLabel: TENANT_KEY_TO_LABEL[tenantKey],
            family: 'azureResource',
            displayName: roleName,
            scope,
            resource: parsedScope.resource,
            resourceType: parsedScope.resourceType,
            membership: 'Group',
            endTime: entry.properties?.endDateTime ? new Date(entry.properties.endDateTime).toLocaleString() : 'Permanent',
            endTimeIso: entry.properties?.endDateTime,
            state: 'eligible',
            source: {
              roleEligibilityScheduleId: linkedRoleEligibilityScheduleId,
              principalId,
              roleDefinitionId,
              subscriptionId
            }
          };

          roleMap.set(roleKey, role);
          if (normalizedScheduleId) {
            eligibleByScheduleId.set(normalizedScheduleId, role);
          }
          eligibleBySemanticWithPrincipal.set(semanticWithPrincipal, role);
          eligibleBySemanticWithoutPrincipal.set(semanticWithoutPrincipal, role);
        }
      }

      if (assignmentResponse.status === 'fulfilled') {
        for (const entry of assignmentResponse.value.value) {
          const scope = entry.properties?.scope ?? `/subscriptions/${subscriptionId}`;
          const parsedScope = this.parseAzureScope(scope, subscriptionLabel);
          const roleDefinitionId = entry.properties?.roleDefinitionId ?? '';
          const principalId = entry.properties?.principalId ?? '';
          const roleGuid = this.extractRoleGuid(roleDefinitionId);
          const roleName =
            entry.properties?.expandedProperties?.roleDefinition?.displayName ||
            (roleGuid ? this.knownAzureRoleNames[roleGuid] : undefined) ||
            'Azure Role';
          const linkedScheduleId = this.normalizeScheduleId(entry.properties?.linkedRoleEligibilityScheduleId);
          const semanticWithPrincipal = `${principalId}|${roleDefinitionId}|${scope}`;
          const semanticWithoutPrincipal = `${roleDefinitionId}|${scope}`;
          const existing =
            (linkedScheduleId ? eligibleByScheduleId.get(linkedScheduleId) : undefined) ||
            eligibleBySemanticWithPrincipal.get(semanticWithPrincipal) ||
            eligibleBySemanticWithoutPrincipal.get(semanticWithoutPrincipal);
          if (!existing) {
            continue;
          }

          const roleKey = `${existing.source.principalId ?? principalId}|${roleDefinitionId}|${scope}`;

          roleMap.set(roleKey, {
            id: existing.id,
            tenantKey,
            tenantLabel: TENANT_KEY_TO_LABEL[tenantKey],
            family: 'azureResource',
            displayName: roleName,
            scope,
            resource: parsedScope.resource,
            resourceType: parsedScope.resourceType,
            membership: existing.membership,
            endTime: entry.properties?.endDateTime ? new Date(entry.properties.endDateTime).toLocaleString() : existing.endTime,
            endTimeIso: entry.properties?.endDateTime ?? existing.endTimeIso,
            state: 'active',
            source: {
              roleEligibilityScheduleId: existing.source.roleEligibilityScheduleId || linkedScheduleId,
              subscriptionId
            }
          });
        }
      }

      return Array.from(roleMap.values());
    });

    for (const result of results) {
      if (result.status === 'fulfilled') {
        roles.push(...result.value);
      }
    }

    const deduped = new Map<string, PimRole>();
    for (const role of roles) {
      const baseKey = role.source.roleEligibilityScheduleId || role.id;
      const key = `${baseKey}|${role.scope}`;
      if (!deduped.has(key)) {
        deduped.set(key, role);
      }
    }

    const uniqueRoles = Array.from(deduped.values()).sort((a, b) => {
      if (a.displayName === b.displayName) {
        return a.resource.localeCompare(b.resource);
      }

      return a.displayName.localeCompare(b.displayName);
    });

    this.azureRolesCache[tenantKey] = {
      expiresAt: now + 45_000,
      roles: uniqueRoles
    };

    return uniqueRoles;
  }

  private extractRoleGuid(roleDefinitionId: string): string | null {
    const match = /\/roleDefinitions\/([0-9a-fA-F-]{36})$/i.exec(roleDefinitionId);
    return match ? match[1].toLowerCase() : null;
  }

  private normalizeScheduleId(value?: string): string {
    if (!value) {
      return '';
    }

    const normalized = value.trim().toLowerCase();
    const parts = normalized.split('/').filter(Boolean);
    return parts.length > 0 ? parts[parts.length - 1] : normalized;
  }

  private parseAzureActivationTarget(roleId: string): {
    scope: string;
    linkedRoleEligibilityScheduleId: string;
    principalId?: string;
    roleDefinitionId?: string;
  } {
    // New format: URL-encoded JSON payload.
    try {
      const decoded = JSON.parse(decodeURIComponent(roleId)) as {
        scope?: string;
        linkedRoleEligibilityScheduleId?: string;
        principalId?: string;
        roleDefinitionId?: string;
      };

      if (decoded.scope && decoded.linkedRoleEligibilityScheduleId) {
        return {
          scope: decoded.scope,
          linkedRoleEligibilityScheduleId: decoded.linkedRoleEligibilityScheduleId,
          principalId: decoded.principalId,
          roleDefinitionId: decoded.roleDefinitionId
        };
      }
    } catch {
      // ignore and try legacy parsing for backward compatibility
    }

    // Legacy format: <encodedScope>|<linkedEligibilityId>
    const separatorIndex = roleId.indexOf('|');
    if (separatorIndex < 0) {
      throw new Error('Invalid Azure role identifier. Refresh the role list and retry.');
    }

    const encodedScope = roleId.slice(0, separatorIndex);
    const linkedRoleEligibilityScheduleId = roleId.slice(separatorIndex + 1);

    return {
      scope: decodeURIComponent(encodedScope),
      linkedRoleEligibilityScheduleId
    };
  }

  private async resolveAzureActivationTarget(arm: ArmClient, target: {
    scope: string;
    linkedRoleEligibilityScheduleId: string;
    principalId?: string;
    roleDefinitionId?: string;
  }): Promise<{
    linkedRoleEligibilityScheduleId: string;
    principalId?: string;
    roleDefinitionId?: string;
  }> {
    try {
      const response = await arm.get<ArmCollection<ArmRoleEligibility>>(
        `${target.scope}/providers/Microsoft.Authorization/roleEligibilitySchedules?$filter=asTarget()&api-version=2020-10-01`
      );

      // Prefer exact id match first, then role+principal match.
      const byId = response.value.find((item) => item.name === target.linkedRoleEligibilityScheduleId);
      const byRoleAndPrincipal = response.value.find(
        (item) =>
          (!target.roleDefinitionId || item.properties?.roleDefinitionId === target.roleDefinitionId) &&
          (!target.principalId || item.properties?.principalId === target.principalId)
      );

      const matched = byId ?? byRoleAndPrincipal;
      if (matched) {
        return {
          linkedRoleEligibilityScheduleId: matched.name,
          principalId: matched.properties?.principalId ?? target.principalId,
          roleDefinitionId: matched.properties?.roleDefinitionId ?? target.roleDefinitionId
        };
      }
    } catch {
      // Fallback to already-carried identifiers when re-resolution fails.
    }

    return {
      linkedRoleEligibilityScheduleId: target.linkedRoleEligibilityScheduleId,
      principalId: target.principalId,
      roleDefinitionId: target.roleDefinitionId
    };
  }

  private parseAzureScope(scope: string, subscriptionDisplayName: string): { resource: string; resourceType: string } {
    const mgMatch = /\/providers\/Microsoft\.Management\/managementGroups\/([^/]+)/i.exec(scope);
    if (mgMatch) {
      return { resource: mgMatch[1], resourceType: 'Management group' };
    }

    const subOnly = /^\/subscriptions\/([^/]+)$/i.exec(scope);
    if (subOnly) {
      return { resource: subscriptionDisplayName, resourceType: 'Subscription' };
    }

    const rgOnly = /^\/subscriptions\/[^/]+\/resourceGroups\/([^/]+)$/i.exec(scope);
    if (rgOnly) {
      return { resource: rgOnly[1], resourceType: 'Resource group' };
    }

    const providerMatch = /\/providers\/([^/]+)\/([^/]+)/i.exec(scope);
    if (providerMatch) {
      const segments = scope.split('/').filter(Boolean);
      const resourceName = segments[segments.length - 1] ?? scope;
      return { resource: resourceName, resourceType: `${providerMatch[1]}/${providerMatch[2]}` };
    }

    return { resource: scope, resourceType: 'Scope' };
  }

  private async mapWithConcurrency<T, R>(items: T[], concurrency: number, worker: (item: T) => Promise<R>): Promise<PromiseSettledResult<R>[]> {
    const results: PromiseSettledResult<R>[] = [];
    let index = 0;

    const runWorker = async (): Promise<void> => {
      while (index < items.length) {
        const current = items[index];
        index += 1;

        try {
          const value = await worker(current);
          results.push({ status: 'fulfilled', value });
        } catch (reason) {
          results.push({ status: 'rejected', reason });
        }
      }
    };

    const workers = Array.from({ length: Math.max(1, Math.min(concurrency, items.length)) }, () => runWorker());
    await Promise.all(workers);
    return results;
  }

  private getGraphClientForTenant(tenantKey: PimTenantKey): GraphClient {
    const tenantId = TENANT_KEY_TO_ID[tenantKey];
    return new GraphClient(() => this.auth.getGraphTokenForTenant(tenantId));
  }

  private getArmClientForTenant(tenantKey: PimTenantKey): ArmClient {
    const tenantId = TENANT_KEY_TO_ID[tenantKey];
    return new ArmClient(() => this.auth.getArmTokenForTenant(tenantId));
  }

  private toTenantKey(value?: string): PimTenantKey | undefined {
    if (value === 'nuance' || value === 'healthcareCloud') {
      return value;
    }

    return undefined;
  }

  private async getArmRequestorPrincipalId(tenantKey: PimTenantKey): Promise<string> {
    const token = await this.auth.getArmTokenForTenant(TENANT_KEY_TO_ID[tenantKey]);
    const parts = token.split('.');
    if (parts.length < 2) {
      throw new Error('Failed to parse ARM token for requestor identity. Sign in again and retry.');
    }

    try {
      const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8')) as {
        oid?: string;
        sub?: string;
      };
      const principalId = payload.oid ?? payload.sub;

      if (!principalId) {
        throw new Error();
      }

      return principalId;
    } catch {
      throw new Error('Failed to resolve requestor principal ID from ARM token. Sign in again and retry.');
    }
  }
}
