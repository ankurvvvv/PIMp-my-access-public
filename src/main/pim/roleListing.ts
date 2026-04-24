import { GraphClient } from '../api/graphClient';
import { ArmClient } from '../api/armClient';
import { TENANT_KEY_TO_LABEL } from './tenantConfig';
import type { PimRole, PimTenantKey } from './types';

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

export interface ArmRoleEligibility {
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

export interface ArmCollection<T> {
  value: T[];
}

const KNOWN_AZURE_ROLE_NAMES: Record<string, string> = {
  'b24988ac-6180-42a0-ab88-20f7382dd24c': 'Contributor',
  '8e3af657-a8ff-443c-a75c-2fe8c4bcb635': 'Owner',
  'acdd72a7-3385-48ef-bd42-f606fba81ae7': 'Reader',
  'f58310d9-a9f6-439a-9e8d-f62e7b41a168': 'User Access Administrator'
};

const azureRolesCache: Partial<Record<PimTenantKey, { expiresAt: number; roles: PimRole[] }>> = {};

export function clearAzureRolesCache(tenantKey: PimTenantKey): void {
  azureRolesCache[tenantKey] = undefined;
}

export async function listEntraEligibleRoles(tenantKey: PimTenantKey, graph: GraphClient): Promise<PimRole[]> {
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

export async function listGroupEligibleRoles(tenantKey: PimTenantKey, graph: GraphClient): Promise<PimRole[]> {
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

export async function listAzureResourceEligibleRoles(tenantKey: PimTenantKey, arm: ArmClient): Promise<PimRole[]> {
  const now = Date.now();
  const cached = azureRolesCache[tenantKey];
  if (cached && cached.expiresAt > now) {
    return cached.roles;
  }

  const subscriptionsResponse = await arm.get<ArmCollection<ArmSubscription>>('/subscriptions?api-version=2020-01-01');
  const roles: PimRole[] = [];
  const bySubscription = new Map(subscriptionsResponse.value.map((sub) => [sub.subscriptionId, sub.displayName]));

  // Track every parent scope (management group / tenant root) seen on any
  // eligibility row across all subscriptions. The ARM `$filter=asTarget()`
  // filter returns inherited eligibilities at parent scopes when queried at a
  // child subscription, BUT it does NOT return active assignments that live at
  // those parent scopes (per
  // https://learn.microsoft.com/rest/api/authorization/role-assignment-schedule-instances/list-for-scope:
  // "asTarget() returns instances created for the current user"). So we need a
  // second pass that queries assignment instances directly at each MG scope.
  const parentScopesSeen = new Set<string>();
  const subscriptionIds = subscriptionsResponse.value.map((sub) => sub.subscriptionId);

  const results = await mapWithConcurrency(subscriptionIds, 10, async (subscriptionId) => {
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
        // Track non-subscription (i.e. management group / tenant root) scopes
        // so the second pass below can also query their assignment instances.
        if (isParentScope(scope)) {
          parentScopesSeen.add(scope);
        }
        const parsedScope = parseAzureScope(scope, subscriptionLabel);
        const roleDefinitionId = entry.properties?.roleDefinitionId ?? '';
        const principalId = entry.properties?.principalId ?? '';
        const normalizedScheduleId = normalizeScheduleId(entry.properties?.roleEligibilityScheduleId || entry.id || entry.name);
        const roleGuid = extractRoleGuid(roleDefinitionId);
        const roleName =
          entry.properties?.expandedProperties?.roleDefinition?.displayName ||
          (roleGuid ? KNOWN_AZURE_ROLE_NAMES[roleGuid] : undefined) ||
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
        // Azure ARM returns TWO kinds of role assignment instances from this endpoint:
        //   1. assignmentType === 'Activated' → a temporary activation born from a PIM
        //      eligibility (this is what "Active Roles" should show).
        //   2. assignmentType === 'Assigned'  → a standing/permanent direct role
        //      assignment that was never PIM-eligible (these must NOT appear in either
        //      the Eligible Roles list or the Active Roles list).
        //
        // We also treat a missing linkedRoleEligibilityScheduleId as a strong signal
        // of a direct (non-PIM) assignment, in case assignmentType is absent on some
        // tenants/regions. Both checks together = belt-and-suspenders filtering.
        const assignmentType = entry.properties?.assignmentType;
        const linkedScheduleId = normalizeScheduleId(entry.properties?.linkedRoleEligibilityScheduleId);
        const isPimActivation = assignmentType === 'Activated' || Boolean(linkedScheduleId);
        if (!isPimActivation) {
          continue;
        }

        const scope = entry.properties?.scope ?? `/subscriptions/${subscriptionId}`;
        const parsedScope = parseAzureScope(scope, subscriptionLabel);
        const roleDefinitionId = entry.properties?.roleDefinitionId ?? '';
        const principalId = entry.properties?.principalId ?? '';
        const roleGuid = extractRoleGuid(roleDefinitionId);
        const roleName =
          entry.properties?.expandedProperties?.roleDefinition?.displayName ||
          (roleGuid ? KNOWN_AZURE_ROLE_NAMES[roleGuid] : undefined) ||
          'Azure Role';
        const semanticWithPrincipal = `${principalId}|${roleDefinitionId}|${scope}`;
        const semanticWithoutPrincipal = `${roleDefinitionId}|${scope}`;

        // Try to pair this activation back to the eligibility row we already indexed
        // (preferred path — keeps a stable id and inherits membership/source fields).
        const existing =
          (linkedScheduleId ? eligibleByScheduleId.get(linkedScheduleId) : undefined) ||
          eligibleBySemanticWithPrincipal.get(semanticWithPrincipal) ||
          eligibleBySemanticWithoutPrincipal.get(semanticWithoutPrincipal);

        if (existing) {
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
          continue;
        }

        // Fallback: this is a real PIM activation but its eligibility row is not
        // visible at this subscription's $filter=asTarget() view. This is the normal
        // case for management-group-scope eligibilities and for group-based PIM,
        // where the eligibility lives at a higher scope but the activation is
        // inherited down to the subscription. Synthesize an active role directly
        // from the activation so it still surfaces under "Active Roles".
        const fallbackScheduleId = linkedScheduleId || normalizeScheduleId(entry.id || entry.name);
        const synthesizedTarget = encodeURIComponent(
          JSON.stringify({
            scope,
            linkedRoleEligibilityScheduleId: fallbackScheduleId,
            principalId,
            roleDefinitionId
          })
        );
        const roleKey = `${principalId}|${roleDefinitionId}|${scope}`;

        roleMap.set(roleKey, {
          id: synthesizedTarget,
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
          state: 'active',
          source: {
            roleEligibilityScheduleId: fallbackScheduleId,
            principalId,
            roleDefinitionId,
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

  // ── Second pass: query active assignments at each parent (MG) scope ──
  //
  // Why: ARM's `$filter=asTarget()` against `roleAssignmentScheduleInstances`
  // only returns assignments whose principal IS the current user at THIS
  // scope. It does not walk up the scope hierarchy for active assignments the
  // way it does for eligibilities. So when a user activates a role that lives
  // at an MG scope, the per-subscription query above will not see it. We have
  // to ask the MG scope directly.
  if (parentScopesSeen.size > 0) {
    // Build a global eligibility index from everything the per-sub pass
    // collected. The MG-scope active rows we are about to fetch may match an
    // eligibility that was originally returned via a DIFFERENT subscription's
    // inheritance view, so the index has to be global, not per-subscription.
    const globalEligibleByScheduleId = new Map<string, PimRole>();
    const globalEligibleBySemanticWithPrincipal = new Map<string, PimRole>();
    const globalEligibleBySemanticWithoutPrincipal = new Map<string, PimRole>();
    for (const role of roles) {
      if (role.state !== 'eligible') {
        continue;
      }
      const scheduleId = normalizeScheduleId(role.source.roleEligibilityScheduleId);
      if (scheduleId) {
        globalEligibleByScheduleId.set(scheduleId, role);
      }
      const principalId = role.source.principalId ?? '';
      const roleDefinitionId = role.source.roleDefinitionId ?? '';
      globalEligibleBySemanticWithPrincipal.set(`${principalId}|${roleDefinitionId}|${role.scope}`, role);
      globalEligibleBySemanticWithoutPrincipal.set(`${roleDefinitionId}|${role.scope}`, role);
    }

    const parentScopes = Array.from(parentScopesSeen);
    const parentResults = await mapWithConcurrency(parentScopes, 10, async (scope) => {
      let assignmentResponse: ArmCollection<ArmRoleEligibility>;
      try {
        assignmentResponse = await arm.get<ArmCollection<ArmRoleEligibility>>(
          `${scope}/providers/Microsoft.Authorization/roleAssignmentScheduleInstances?$filter=asTarget()&api-version=2020-10-01`
        );
      } catch {
        // A 403/404 at this MG scope is not fatal — the user might not have
        // read permission on every parent MG. Just skip and keep going.
        return [] as PimRole[];
      }

      const parsedScope = parseAzureScope(scope, scope);
      const activeRoles: PimRole[] = [];

      for (const entry of assignmentResponse.value) {
        // Same PIM-vs-permanent guard as the per-sub pass: only surface real
        // activations (assignmentType === 'Activated' OR a linked eligibility
        // schedule id is present). Standing/permanent direct assignments must
        // not appear in either Eligible Roles or Active Roles.
        const assignmentType = entry.properties?.assignmentType;
        const linkedScheduleId = normalizeScheduleId(entry.properties?.linkedRoleEligibilityScheduleId);
        const isPimActivation = assignmentType === 'Activated' || Boolean(linkedScheduleId);
        if (!isPimActivation) {
          continue;
        }

        const entryScope = entry.properties?.scope ?? scope;
        const roleDefinitionId = entry.properties?.roleDefinitionId ?? '';
        const principalId = entry.properties?.principalId ?? '';
        const roleGuid = extractRoleGuid(roleDefinitionId);
        const roleName =
          entry.properties?.expandedProperties?.roleDefinition?.displayName ||
          (roleGuid ? KNOWN_AZURE_ROLE_NAMES[roleGuid] : undefined) ||
          'Azure Role';
        const semanticWithPrincipal = `${principalId}|${roleDefinitionId}|${entryScope}`;
        const semanticWithoutPrincipal = `${roleDefinitionId}|${entryScope}`;
        const matchedEligibility =
          (linkedScheduleId ? globalEligibleByScheduleId.get(linkedScheduleId) : undefined) ||
          globalEligibleBySemanticWithPrincipal.get(semanticWithPrincipal) ||
          globalEligibleBySemanticWithoutPrincipal.get(semanticWithoutPrincipal);

        const fallbackScheduleId = linkedScheduleId || normalizeScheduleId(entry.id || entry.name);
        const synthesizedTarget =
          matchedEligibility?.id ??
          encodeURIComponent(
            JSON.stringify({
              scope: entryScope,
              linkedRoleEligibilityScheduleId: fallbackScheduleId,
              principalId,
              roleDefinitionId
            })
          );

        activeRoles.push({
          id: synthesizedTarget,
          tenantKey,
          tenantLabel: TENANT_KEY_TO_LABEL[tenantKey],
          family: 'azureResource',
          displayName: roleName,
          scope: entryScope,
          resource: parsedScope.resource,
          resourceType: parsedScope.resourceType,
          membership: matchedEligibility?.membership ?? 'Group',
          endTime: entry.properties?.endDateTime
            ? new Date(entry.properties.endDateTime).toLocaleString()
            : matchedEligibility?.endTime ?? 'Permanent',
          endTimeIso: entry.properties?.endDateTime ?? matchedEligibility?.endTimeIso,
          state: 'active',
          source: {
            roleEligibilityScheduleId:
              matchedEligibility?.source.roleEligibilityScheduleId || fallbackScheduleId,
            principalId,
            roleDefinitionId
          }
        });
      }

      return activeRoles;
    });

    for (const result of parentResults) {
      if (result.status === 'fulfilled') {
        roles.push(...result.value);
      }
    }
  }

  const deduped = new Map<string, PimRole>();
  for (const role of roles) {
    const baseKey = role.source.roleEligibilityScheduleId || role.id;
    const key = `${baseKey}|${role.scope}`;
    const existing = deduped.get(key);
    if (!existing) {
      deduped.set(key, role);
      continue;
    }
    // Prefer the active record over an eligible one when both surface for the
    // same schedule+scope (e.g. MG-scope assignment returned from multiple
    // subscriptions, or eligible-then-synthesized-active ordering).
    if (existing.state !== 'active' && role.state === 'active') {
      deduped.set(key, role);
    }
  }

  const uniqueRoles = Array.from(deduped.values()).sort((a, b) => {
    if (a.displayName === b.displayName) {
      return a.resource.localeCompare(b.resource);
    }

    return a.displayName.localeCompare(b.displayName);
  });

  azureRolesCache[tenantKey] = {
    expiresAt: now + 45_000,
    roles: uniqueRoles
  };

  return uniqueRoles;
}

function extractRoleGuid(roleDefinitionId: string): string | null {
  const match = /\/roleDefinitions\/([0-9a-fA-F-]{36})$/i.exec(roleDefinitionId);
  return match ? match[1].toLowerCase() : null;
}

function normalizeScheduleId(value?: string): string {
  if (!value) {
    return '';
  }

  const normalized = value.trim().toLowerCase();
  const parts = normalized.split('/').filter(Boolean);
  return parts.length > 0 ? parts[parts.length - 1] : normalized;
}

// Returns true when the scope lives ABOVE the subscription level — i.e. a
// management group or the tenant root. Per-subscription `asTarget()` queries
// don't return active assignments at these scopes, so we have to query them
// separately. Resource-group and individual-resource scopes are NOT parent
// scopes — those are reachable from a per-subscription query.
function isParentScope(scope: string): boolean {
  return /\/providers\/Microsoft\.Management\/managementGroups\//i.test(scope);
}

function parseAzureScope(scope: string, subscriptionDisplayName: string): { resource: string; resourceType: string } {
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

async function mapWithConcurrency<T, R>(items: T[], concurrency: number, worker: (item: T) => Promise<R>): Promise<PromiseSettledResult<R>[]> {
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
