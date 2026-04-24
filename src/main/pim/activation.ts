import { randomUUID } from 'crypto';
import { GraphClient } from '../api/graphClient';
import { ArmClient } from '../api/armClient';
import { DeviceCodeAuth } from '../auth/deviceCodeAuth';
import { TENANT_KEY_TO_ID } from './tenantConfig';
import { clearAzureRolesCache } from './roleListing';
import type { ArmRoleEligibility, ArmCollection } from './roleListing';
import type {
  ActivateRoleRequest,
  ArmActivationResult,
  DeactivateRoleRequest,
  GraphActivationResult,
  PimTenantKey
} from './types';

export async function activateRole(
  request: ActivateRoleRequest,
  graph: GraphClient,
  arm: ArmClient,
  auth: DeviceCodeAuth,
  tenantKey: PimTenantKey
): Promise<{ requestId: string; status: string }> {
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

  const azureTarget = parseAzureActivationTarget(request.roleId);
  const { scope } = azureTarget;

  if (!scope.startsWith('/')) {
    throw new Error('Invalid Azure role scope. Refresh the role list and retry.');
  }

  const resolved = await resolveAzureActivationTarget(arm, azureTarget);
  const requestorPrincipalId = await getArmRequestorPrincipalId(auth, tenantKey);

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

  clearAzureRolesCache(tenantKey);
  return { requestId: result.name ?? requestId, status: result.properties?.status ?? 'submitted' };
}

function parseAzureActivationTarget(roleId: string): {
  scope: string;
  linkedRoleEligibilityScheduleId: string;
  principalId?: string;
  roleDefinitionId?: string;
} {
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

async function resolveAzureActivationTarget(arm: ArmClient, target: {
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

async function getArmRequestorPrincipalId(auth: DeviceCodeAuth, tenantKey: PimTenantKey): Promise<string> {
  const token = await auth.getArmTokenForTenant(TENANT_KEY_TO_ID[tenantKey]);
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

// ── Deactivation ──
//
// Ends an active PIM assignment early. Same family routing as activateRole:
//   - Entra: Graph self-deactivate request.
//   - Group: Graph group self-deactivate request.
//   - Azure resource: ARM self-deactivate request at the role's scope.
//
// On success the per-tenant Azure role cache is cleared so the next list call
// returns fresh state. The renderer also mutates state locally for instant UI.
export async function deactivateRole(
  request: DeactivateRoleRequest,
  graph: GraphClient,
  arm: ArmClient,
  auth: DeviceCodeAuth,
  tenantKey: PimTenantKey
): Promise<{ requestId: string; status: string }> {
  if (request.family === 'entra') {
    const [principalId, roleDefinitionId, directoryScopeId] = request.roleId.split('|');
    const result = await graph.post<GraphActivationResult>('/roleManagement/directory/roleAssignmentScheduleRequests', {
      action: 'selfDeactivate',
      principalId,
      roleDefinitionId,
      directoryScopeId: directoryScopeId || '/'
    });

    return { requestId: result.id, status: result.status };
  }

  if (request.family === 'group') {
    const [principalId, groupId, accessId] = request.roleId.split('|');
    const result = await graph.post<GraphActivationResult>('/identityGovernance/privilegedAccess/group/assignmentScheduleRequests', {
      action: 'selfDeactivate',
      principalId,
      groupId,
      accessId
    });

    return { requestId: result.id, status: result.status };
  }

  // Azure resource role — same scope/principal resolution as activation, but
  // requestType: 'SelfDeactivate'. No duration, no justification.
  const azureTarget = parseAzureActivationTarget(request.roleId);
  const { scope } = azureTarget;

  if (!scope.startsWith('/')) {
    throw new Error('Invalid Azure role scope. Refresh the role list and retry.');
  }

  const resolved = await resolveAzureActivationTarget(arm, azureTarget);
  const requestorPrincipalId = await getArmRequestorPrincipalId(auth, tenantKey);

  if (!resolved.roleDefinitionId) {
    throw new Error('Missing Azure role definition ID. Refresh the role list and retry.');
  }

  const requestId = randomUUID();
  const result = await arm.put<ArmActivationResult>(
    `${scope}/providers/Microsoft.Authorization/roleAssignmentScheduleRequests/${requestId}?api-version=2020-10-01`,
    {
      properties: {
        requestType: 'SelfDeactivate',
        principalId: requestorPrincipalId,
        roleDefinitionId: resolved.roleDefinitionId
      }
    }
  );

  clearAzureRolesCache(tenantKey);
  return { requestId: result.name ?? requestId, status: result.properties?.status ?? 'submitted' };
}
