import {
  state, dom, ROLE_FAMILIES, TENANT_KEYS, ROLES_CACHE_TTL_MS, OPTIMISTIC_GRACE_MS,
  activeFamily, selectedTenantKey, isActivationRawOpen,
  setActiveFamily, setActivationRawOpen, setSelectedTenantKey,
  getPimClient, escapeHtml
} from './state.js';

// ── Activation status ──

function getHumanActivationMessage(rawMessage, options = {}) {
  const isError = Boolean(options.isError);
  const statusHint = String(options.statusHint || '');
  const raw = String(rawMessage || '');
  const haystack = `${statusHint} ${raw}`.toLowerCase();

  if (haystack.includes('roleassignmentexists') || haystack.includes('already exists')) {
    return 'Role is already active';
  }

  if (haystack.includes('pending')) {
    return 'Pending approval';
  }

  if (haystack.includes('provisioned') || haystack.includes('granted') || haystack.includes('active')) {
    return 'Role is active now';
  }

  if (
    haystack.includes('authorization') ||
    haystack.includes('denied') ||
    haystack.includes('forbidden') ||
    haystack.includes('insufficient privileges')
  ) {
    return 'Role activation denied';
  }

  if (
    haystack.includes('invalid') ||
    haystack.includes('missing') ||
    haystack.includes('badrequest') ||
    haystack.includes('scope') ||
    haystack.includes('definition')
  ) {
    return 'Please check the input values';
  }

  if (
    haystack.includes('network') ||
    haystack.includes('timeout') ||
    haystack.includes('timed out') ||
    haystack.includes('econn') ||
    haystack.includes('enotfound') ||
    haystack.includes('fetch') ||
    haystack.includes('arm ') ||
    haystack.includes('graph ')
  ) {
    return 'Unable to reach Azure. Please try again';
  }

  if (!isError) {
    return 'Pending approval';
  }

  return 'Unable to reach Azure. Please try again';
}

export function setActivationStatusMessage(humanMessage, rawMessage = '') {
  if (!dom.activationStatus || !dom.activationHumanMessage || !dom.activationRawMessage || !dom.activationRawToggle) {
    return;
  }

  dom.activationHumanMessage.textContent = humanMessage;
  dom.activationRawMessage.textContent = rawMessage;
  setActivationRawOpen(false);

  const hasRawMessage = Boolean(rawMessage);
  dom.activationRawToggle.hidden = !hasRawMessage;
  dom.activationRawMessage.hidden = true;
  dom.activationRawToggle.setAttribute('aria-expanded', 'false');
  dom.activationRawToggle.innerHTML = '&#9656;';
}

function toggleActivationRawMessage() {
  if (!dom.activationRawMessage || !dom.activationRawToggle || dom.activationRawToggle.hidden) {
    return;
  }

  setActivationRawOpen(!isActivationRawOpen);
  dom.activationRawMessage.hidden = !isActivationRawOpen;
  dom.activationRawToggle.setAttribute('aria-expanded', String(isActivationRawOpen));
  dom.activationRawToggle.innerHTML = isActivationRawOpen ? '&#9662;' : '&#9656;';
}

// ── Roles loading ──

function setRolesLoading(isLoading) {
  state.isRolesLoading = isLoading;
  if (dom.rolesLoadingIndicator) {
    dom.rolesLoadingIndicator.classList.toggle('active', isLoading);
    dom.rolesLoadingIndicator.setAttribute('aria-hidden', String(!isLoading));
  }

  if (dom.refreshBtn) {
    dom.refreshBtn.disabled = isLoading;
  }
}

// Silent refresh button state: disable it and relabel to "Syncing…" so the
// user gets a subtle hint that work is in progress without any skeleton /
// wipe of existing rows. Pairs with the silent branch inside fetchAllRoles.
const REFRESH_BTN_IDLE_LABEL = 'Refresh Roles';
const REFRESH_BTN_SYNCING_LABEL = 'Syncing…';
function setRefreshButtonSyncing(isSyncing) {
  if (!dom.refreshBtn) {
    return;
  }
  dom.refreshBtn.disabled = isSyncing;
  dom.refreshBtn.textContent = isSyncing ? REFRESH_BTN_SYNCING_LABEL : REFRESH_BTN_IDLE_LABEL;
}

function setRolesStatus(message) {
  if (dom.rolesStatus) {
    dom.rolesStatus.textContent = message;
  }
}

function getFamilyLabel(family) {
  if (family === 'azureResource') {
    return 'Azure Resources';
  }

  if (family === 'group') {
    return 'Groups';
  }

  return 'Entra Roles';
}

function updateRoleTabsMeta() {
  const tenantRoles = state.rolesByTenantFamily[selectedTenantKey] || state.rolesByTenantFamily.nuance;
  const tenantSt = state.familyStatusByTenant[selectedTenantKey] || state.familyStatusByTenant.nuance;

  for (const tab of dom.familyTabs) {
    const family = tab.dataset.family;
    if (!family) {
      continue;
    }

    const count = tenantRoles[family].length;
    const status = tenantSt[family];
    const familyLabel = getFamilyLabel(family);

    tab.classList.toggle('loading', status === 'loading');
    tab.classList.toggle('error', status === 'error');
    tab.textContent = `${familyLabel} (${count})`;
  }
}

function hasAnyCachedRoles() {
  return TENANT_KEYS.some((tenantKey) => ROLE_FAMILIES.some((family) => state.rolesByTenantFamily[tenantKey][family].length > 0));
}

function renderLoadingRows() {
  dom.rolesBody.innerHTML = '';
  const skeletonRows = 6;
  for (let i = 0; i < skeletonRows; i += 1) {
    const row = document.createElement('tr');
    row.innerHTML = `
      <td class="skeleton-cell"><div class="skeleton-bar"></div></td>
      <td class="skeleton-cell"><div class="skeleton-bar"></div></td>
      <td class="skeleton-cell"><div class="skeleton-bar"></div></td>
      <td class="skeleton-cell"><div class="skeleton-bar"></div></td>
      <td class="skeleton-cell"><div class="skeleton-bar"></div></td>
    `;
    dom.rolesBody.appendChild(row);
  }
}

// ── Fetch ──

async function fetchRolesForFamily(tenantKey, family, loadToken) {
  try {
    const roles = await getPimClient().listEligibleRoles({ tenantKey, family });
    if (loadToken !== state.activeRolesLoadToken) {
      return;
    }

    state.rolesByTenantFamily[tenantKey][family] = mergeFamilyRoles(
      state.rolesByTenantFamily[tenantKey][family] || [],
      roles
    );
    state.familyStatusByTenant[tenantKey][family] = 'ready';
  } catch (error) {
    if (loadToken !== state.activeRolesLoadToken) {
      return;
    }

    state.rolesByTenantFamily[tenantKey][family] = [];
    state.familyStatusByTenant[tenantKey][family] = 'error';
  }

  updateRoleTabsMeta();
  if (activeFamily === family && selectedTenantKey === tenantKey) {
    renderRoles();
    renderActiveRoles();
  }
}

// Merge a freshly-fetched role list from the backend into the existing local
// list for one (tenant, family) bucket.
//
// Why merge instead of replace: after the user activates a role we flip it to
// `state: 'active'` locally (see applyOptimisticActivation) and stamp it with
// `optimisticActivatedAt`. Azure ARM's listing endpoints are eventually
// consistent — a Refresh within ~30–90s of activation can return data that
// does NOT yet include the just-activated row. A blind replace would drop the
// role from the Active Roles panel and confuse the user.
//
// Rules:
//  1. If backend returned the role (matched by id) → take the backend row as
//     the new source of truth (fresh end time, resource, etc.).
//  2. If backend did NOT return an existing role AND that role is a
//     still-in-grace optimistic activation → keep the local row.
//  3. Otherwise (backend dropped the role AND it's outside the grace window)
//     → drop it locally too. This lets external revokes / expiries take
//     effect once Azure has caught up.
//  4. Any backend row with no local match is appended as-is.
function mergeFamilyRoles(existingRoles, incomingRoles) {
  const now = Date.now();
  const incomingById = new Map();
  for (const incoming of incomingRoles) {
    if (incoming && incoming.id) {
      incomingById.set(incoming.id, incoming);
    }
  }

  const merged = [];
  const consumedIncomingIds = new Set();

  for (const existing of existingRoles) {
    if (!existing || !existing.id) {
      continue;
    }

    const incoming = incomingById.get(existing.id);
    if (incoming) {
      // Backend confirmed this role — carry forward the optimistic stamp so a
      // second Refresh within the grace window is still protected if the
      // listing briefly flickers.
      const carriedStamp = existing.optimisticActivatedAt;
      merged.push(carriedStamp ? { ...incoming, optimisticActivatedAt: carriedStamp } : incoming);
      consumedIncomingIds.add(existing.id);
      continue;
    }

    const isOptimisticActive =
      existing.state === 'active' &&
      typeof existing.optimisticActivatedAt === 'number' &&
      now - existing.optimisticActivatedAt < OPTIMISTIC_GRACE_MS;

    if (isOptimisticActive) {
      // Protect the just-activated row from being wiped by a stale backend
      // snapshot. The backend will catch up within the grace window.
      merged.push(existing);
    }
    // else: drop — backend says it's gone and we have no reason to disbelieve.
  }

  for (const incoming of incomingRoles) {
    if (incoming && incoming.id && !consumedIncomingIds.has(incoming.id)) {
      merged.push(incoming);
    }
  }

  return merged;
}

export async function fetchAllRoles(options = {}) {
  const force = Boolean(options.force);
  // Silent refresh = no skeleton flash, no "loading" wipe. The current rows
  // stay on screen; only the Refresh button is disabled and relabeled so the
  // user knows work is in progress. We default to silent AFTER the first
  // successful post-login load — that first load still uses the skeleton
  // because there is nothing on screen worth preserving.
  const silent = typeof options.silent === 'boolean' ? options.silent : state.isFirstLoadDone;
  const now = Date.now();

  if (!force && now < state.rolesCacheUntil && hasAnyCachedRoles()) {
    renderRoles();
    renderActiveRoles();
    setRolesStatus('Showing recent role data.');
    return;
  }

  const loadToken = state.activeRolesLoadToken + 1;
  state.activeRolesLoadToken = loadToken;
  state.rolesCacheUntil = 0;

  if (silent) {
    setRefreshButtonSyncing(true);
  } else {
    setRolesLoading(true);
  }

  for (const tenantKey of TENANT_KEYS) {
    for (const family of ROLE_FAMILIES) {
      state.familyStatusByTenant[tenantKey][family] = 'loading';
    }
  }

  updateRoleTabsMeta();
  if (!silent) {
    renderHeaders();
    renderLoadingRows();
  }
  setRolesStatus('Pulling roles across Entra, Azure Resources, and Groups...');

  const tasks = [];
  for (const tenantKey of TENANT_KEYS) {
    for (const family of ROLE_FAMILIES) {
      tasks.push(fetchRolesForFamily(tenantKey, family, loadToken));
    }
  }
  await Promise.allSettled(tasks);

  if (loadToken !== state.activeRolesLoadToken) {
    return;
  }

  if (silent) {
    setRefreshButtonSyncing(false);
  } else {
    setRolesLoading(false);
  }

  const loadedFamilies = [];
  const failedFamilies = [];
  for (const tenantKey of TENANT_KEYS) {
    for (const family of ROLE_FAMILIES) {
      const status = state.familyStatusByTenant[tenantKey][family];
      if (status === 'ready') {
        loadedFamilies.push(`${tenantKey}:${family}`);
      }
      if (status === 'error') {
        failedFamilies.push(`${tenantKey}:${family}`);
      }
    }
  }

  if (loadedFamilies.length > 0) {
    state.rolesCacheUntil = Date.now() + ROLES_CACHE_TTL_MS;
    state.isFirstLoadDone = true;
  }

  const totalTenantFamilyCombos = TENANT_KEYS.length * ROLE_FAMILIES.length;

  if (loadedFamilies.length === totalTenantFamilyCombos) {
    setRolesStatus('Roles are up to date.');
  } else if (loadedFamilies.length > 0) {
    setRolesStatus(`Partial load complete (${loadedFamilies.length}/${totalTenantFamilyCombos}).`);
  } else {
    setRolesStatus('Unable to load roles right now. Please try refresh.');
  }

  if (failedFamilies.length > 0) {
    const raw = `Some tenant categories failed to load: ${failedFamilies.join(', ')}.`;
    setActivationStatusMessage('Unable to reach Azure. Please try again', raw);
  }

  renderRoles();
  renderActiveRoles();
}

// ── Duration ──

function normalizeDurationHours(rawValue) {
  const numericValue = Number(rawValue);
  const fallbackValue = 1;
  const safeValue = Number.isFinite(numericValue) ? numericValue : fallbackValue;
  const steppedValue = Math.round(safeValue * 2) / 2;
  return Math.min(8, Math.max(0.5, steppedValue));
}

function getDurationUnit(durationHours) {
  if (durationHours < 1) {
    return 'min';
  }

  if (durationHours === 1) {
    return 'hr';
  }

  return 'hrs';
}

function updateDurationSliderFill(durationHours) {
  if (!dom.durationHoursInput) {
    return;
  }

  const minValue = Number(dom.durationHoursInput.min || '0.5');
  const maxValue = Number(dom.durationHoursInput.max || '8');
  const range = maxValue - minValue;
  const percentage = range > 0 ? ((durationHours - minValue) / range) * 100 : 0;
  dom.durationHoursInput.style.setProperty('--slider-fill', `${percentage}%`);
}

export function syncDurationControls(rawValue) {
  const durationHours = normalizeDurationHours(rawValue);

  if (dom.durationHoursInput) {
    dom.durationHoursInput.value = durationHours.toString();
  }

  if (dom.durationHoursBoxInput) {
    dom.durationHoursBoxInput.value = durationHours.toFixed(1).replace('.0', '');
  }

  if (dom.durationHoursUnit) {
    dom.durationHoursUnit.textContent = getDurationUnit(durationHours);
  }

  updateDurationSliderFill(durationHours);
}

// ── Expiry helpers ──

function getRoleExpiryWarning(role) {
  const iso = role.endTimeIso;
  if (!iso) {
    return null;
  }

  const endTimeMs = Date.parse(iso);
  if (!Number.isFinite(endTimeMs)) {
    return null;
  }

  const diffMs = endTimeMs - Date.now();
  if (diffMs <= 0) {
    return 'This role is expired.';
  }

  const tenMinutesMs = 10 * 60 * 1000;
  if (diffMs > tenMinutesMs) {
    return null;
  }

  const minutesLeft = Math.max(1, Math.ceil(diffMs / 60000));
  return `This role will expire in next ${minutesLeft} min${minutesLeft === 1 ? '' : 's'}.`;
}

function getExpiryBadgeLabel(role) {
  const iso = role.endTimeIso;
  if (!iso) {
    return null;
  }

  const endTimeMs = Date.parse(iso);
  if (!Number.isFinite(endTimeMs)) {
    return null;
  }

  const diffMs = endTimeMs - Date.now();
  if (diffMs <= 0) {
    return 'Expired';
  }

  const tenMinutesMs = 10 * 60 * 1000;
  if (diffMs > tenMinutesMs) {
    return null;
  }

  const minutesLeft = Math.max(1, Math.ceil(diffMs / 60000));
  return `${minutesLeft}m left`;
}

// ── Render ──

function renderHeaders() {
  if (activeFamily === 'azureResource') {
    dom.rolesHeaderRow.innerHTML = `
      <th>Role Name</th>
      <th>Resource</th>
      <th>Resource type</th>
      <th>End time</th>
      <th>Activate</th>
    `;
    return;
  }

  dom.rolesHeaderRow.innerHTML = `
    <th>Role Name</th>
    <th>Scope</th>
    <th>Membership</th>
    <th>End time</th>
    <th>Activate</th>
  `;
}

function getActiveRoleTypeLabel(family) {
  if (family === 'azureResource') {
    return 'Azure Resource Role';
  }

  if (family === 'group') {
    return 'Group Role';
  }

  return 'Entra Role';
}

function getActiveRoleContextLine(role) {
  const typeLabel = getActiveRoleTypeLabel(role.family);

  if (role.family === 'azureResource') {
    const resourceName = role.resource || role.scope || 'Azure scope';
    return `${typeLabel} - ${resourceName}`;
  }

  if (role.family === 'group') {
    const groupScope = role.scope === '/' ? 'Directory' : role.scope;
    return `${typeLabel} - ${groupScope}`;
  }

  const scope = role.scope === '/' ? 'Directory' : role.scope;
  return `${typeLabel} - ${scope}`;
}

export function renderRoles() {
  renderHeaders();
  if (!dom.rolesBody) {
    return;
  }

  dom.rolesBody.innerHTML = '';
  state.roles = (state.rolesByTenantFamily[selectedTenantKey]?.[activeFamily] || []).filter((role) => role.state !== 'active');

  if (state.isRolesLoading && state.roles.length === 0) {
    renderLoadingRows();
    return;
  }

  if (state.roles.length === 0) {
    dom.rolesBody.innerHTML = '<tr><td colspan="5" class="empty-roles">No eligible roles found for this category.</td></tr>';
    return;
  }

  for (const role of state.roles) {
    const row = document.createElement('tr');
    const displayScope = role.scope === '/' ? 'Directory' : role.scope;
    const expiryWarning = getRoleExpiryWarning(role);
    const roleNameMarkup = expiryWarning
      ? `${escapeHtml(role.displayName)}<div class="expiry-soon-note">${escapeHtml(expiryWarning)}</div>`
      : escapeHtml(role.displayName);

    if (activeFamily === 'azureResource') {
      row.innerHTML = `
        <td class="cell-wrap">${roleNameMarkup}</td>
        <td class="cell-wrap" title="${escapeHtml(role.resource)}">${escapeHtml(role.resource)}</td>
        <td>${escapeHtml(role.resourceType)}</td>
        <td>${escapeHtml(role.endTime)}</td>
        <td><button data-role-id="${escapeHtml(role.id)}">Activate</button></td>
      `;
    } else {
      row.innerHTML = `
        <td class="cell-wrap">${roleNameMarkup}</td>
        <td class="cell-wrap" title="${escapeHtml(displayScope)}">${escapeHtml(displayScope)}</td>
        <td>${escapeHtml(role.membership)}</td>
        <td>${escapeHtml(role.endTime)}</td>
        <td><button data-role-id="${escapeHtml(role.id)}">Activate</button></td>
      `;
    }

    const button = row.querySelector('button');
    button.addEventListener('click', () => activateRole(role.id));

    dom.rolesBody.appendChild(row);
  }
}

export function renderActiveRoles() {
  if (!dom.activeRolesList) {
    return;
  }

  dom.activeRolesList.innerHTML = '';
  const groupedActiveRoles = ROLE_FAMILIES.map((family) => ({
    family,
    roles: (state.rolesByTenantFamily[selectedTenantKey]?.[family] || []).filter((role) => role.state === 'active')
  }));
  const totalActiveCount = groupedActiveRoles.reduce((count, entry) => count + entry.roles.length, 0);

  if (state.isRolesLoading && totalActiveCount === 0) {
    dom.activeRolesList.innerHTML = '<div class="empty-roles">Checking active roles...</div>';
    if (dom.activeRolesStatus) {
      dom.activeRolesStatus.textContent = 'Checking active roles...';
    }
    return;
  }

  if (totalActiveCount === 0) {
    dom.activeRolesList.innerHTML = '<div class="empty-roles">No active roles in any category</div>';
    if (dom.activeRolesStatus) {
      dom.activeRolesStatus.textContent = 'No active roles';
    }
    return;
  }

  for (const familyGroup of groupedActiveRoles) {
    if (familyGroup.roles.length === 0) {
      continue;
    }

    const groupElement = document.createElement('section');
    groupElement.className = 'active-family-group';
    groupElement.innerHTML = `<h3 class="active-family-title">${escapeHtml(getFamilyLabel(familyGroup.family))}</h3>`;

    const listElement = document.createElement('div');
    listElement.className = 'active-role-items';

    for (const role of familyGroup.roles) {
      const item = document.createElement('article');
      item.className = 'active-role-item';
      const expiryBadge = getExpiryBadgeLabel(role);
      const badgeMarkup = expiryBadge
        ? `<span class="active-role-timer">${escapeHtml(expiryBadge)}</span>`
        : '';

      // Layout: role name + (optional expiry badge) + Deactivate button all
      // on the SAME top row to keep cards compact. Details below.
      item.innerHTML = `
        <div class="active-role-topline">
          <p class="active-role-name">${escapeHtml(role.displayName)}</p>
          <div class="active-role-topline-right">
            ${badgeMarkup}
            <button type="button" class="active-role-deactivate-btn" data-role-id="${escapeHtml(role.id)}" data-tenant-key="${escapeHtml(role.tenantKey)}" data-family="${escapeHtml(role.family)}">Deactivate</button>
          </div>
        </div>
        <p class="active-role-context">${escapeHtml(getActiveRoleContextLine(role))}</p>
        <p class="active-role-endtime">End time: ${escapeHtml(role.endTime)}</p>
      `;

      // Wire the Deactivate button. Each render rebuilds the DOM so we don't
      // need to worry about stale handlers — the previous element is gone.
      const deactivateBtn = item.querySelector('.active-role-deactivate-btn');
      if (deactivateBtn) {
        deactivateBtn.addEventListener('click', () => deactivateActiveRole(role));
      }

      listElement.appendChild(item);
    }

    groupElement.appendChild(listElement);
    dom.activeRolesList.appendChild(groupElement);
  }

  if (dom.activeRolesStatus) {
    dom.activeRolesStatus.textContent = `${totalActiveCount} active role(s).`;
  }

  // Schedule (or refresh) per-role expiry timers so the UI flips to "Expired"
  // exactly when each role's endTimeIso is reached, with no polling required.
  scheduleActiveRoleExpiryTimers(groupedActiveRoles);
}

// ── Active role expiry timers (zero-poll) ──
//
// Strategy: one setTimeout per active role, scheduled at (endTimeIso - now).
// When it fires:
//   1. Re-render Active Roles (the existing badge logic flips to "Expired").
//   2. Trigger a silent backend refresh so the role disappears entirely once
//      Azure confirms it's gone.
//
// We track timers in state.expiryTimers (Map<roleId, handle>). Every render
// reconciles: keep timers for roles still present, clear timers for roles
// that vanished. setTimeout is paused when the OS sleeps and fires "late" on
// wake — that's fine here, the Expired badge just appears a bit late, and
// the safety net (focus refresh) will catch up immediately.
const MAX_TIMER_DELAY_MS = 24 * 60 * 60 * 1000; // 24 hours — guard against absurd schedules.

function scheduleActiveRoleExpiryTimers(groupedActiveRoles) {
  const currentActiveIds = new Set();
  for (const familyGroup of groupedActiveRoles) {
    for (const role of familyGroup.roles) {
      currentActiveIds.add(role.id);
      const endTimeMs = role.endTimeIso ? Date.parse(role.endTimeIso) : NaN;
      if (!Number.isFinite(endTimeMs)) {
        continue;
      }
      // Already expired by the clock — no timer needed. The next render will
      // either show "Expired" or the silent refresh will drop it.
      if (endTimeMs <= Date.now()) {
        continue;
      }
      // Already have a live timer for this role — leave it alone.
      if (state.expiryTimers.has(role.id)) {
        continue;
      }
      const delay = Math.min(MAX_TIMER_DELAY_MS, endTimeMs - Date.now());
      const handle = setTimeout(() => {
        state.expiryTimers.delete(role.id);
        // Re-render so the badge flips to "Expired" immediately. We do NOT
        // trigger a backend refresh here — this app is user-initiated only,
        // so the row will sit there labeled "Expired" until the user clicks
        // Refresh. That's by design (predictable, no surprise Azure calls).
        renderActiveRoles();
      }, delay);
      state.expiryTimers.set(role.id, handle);
    }
  }

  // Clear timers for roles that are no longer active (e.g. user just
  // deactivated, or a refresh removed them).
  for (const [id, handle] of state.expiryTimers) {
    if (!currentActiveIds.has(id)) {
      clearTimeout(handle);
      state.expiryTimers.delete(id);
    }
  }
}

// ── Deactivation ──
//
// Self-deactivate an active role early. We optimistically remove the role
// from local state so the Active Roles panel updates instantly, then call the
// backend. If the backend call fails, we restore the role and show an error
// — Azure is the source of truth.
async function deactivateActiveRole(role) {
  if (!role || !role.id) {
    return;
  }

  const tenantKey = role.tenantKey;
  const family = role.family;
  const familyBucket = state.rolesByTenantFamily[tenantKey]?.[family];
  if (!Array.isArray(familyBucket)) {
    return;
  }

  const removalIndex = familyBucket.findIndex((entry) => entry.id === role.id);
  if (removalIndex < 0) {
    return;
  }

  // Snapshot the role before removing so we can roll back if the API fails.
  const removedRole = familyBucket[removalIndex];
  familyBucket.splice(removalIndex, 1);
  renderActiveRoles();
  renderRoles();
  setActivationStatusMessage('Deactivating role...');

  try {
    const result = await getPimClient().deactivateRole({
      tenantKey,
      family,
      roleId: role.id
    });
    const raw = `Deactivation submitted. Request: ${result.requestId}, status: ${result.status}`;
    setActivationStatusMessage('Role is being deactivated', raw);
  } catch (error) {
    // Roll back local state — the role is still active in Azure.
    familyBucket.splice(removalIndex, 0, removedRole);
    renderActiveRoles();
    renderRoles();
    const raw = `Deactivation failed: ${String(error.message || error)}`;
    setActivationStatusMessage('Unable to deactivate role. Please try again', raw);
  }
}

// ── Activation ──

async function activateRole(roleId) {
  const role = state.roles.find((item) => item.id === roleId);
  if (!role) {
    setActivationStatusMessage('Please check the input values');
    return;
  }

  const durationHours = normalizeDurationHours(dom.durationHoursInput ? dom.durationHoursInput.value : '1');
  const justification = document.getElementById('justification').value.trim();

  if (!justification) {
    setActivationStatusMessage('Please check the input values');
    return;
  }

  setActivationStatusMessage('Submitting activation request...');

  try {
    const result = await getPimClient().activateRole({
      tenantKey: role.tenantKey || selectedTenantKey,
      family: role.family,
      roleId,
      durationHours,
      justification
    });

    const raw = `Activation submitted. Request: ${result.requestId}, status: ${result.status}`;
    const human = getHumanActivationMessage(raw, { statusHint: result.status, isError: false });
    setActivationStatusMessage(human, raw);

    // Optimistic local update: if Azure accepted/granted/provisioned the
    // request, flip this role to "active" in local state immediately so the
    // Active Roles panel shows it without waiting for a Refresh round-trip.
    // We compute an end time from the requested duration. The next backend
    // refresh (focus, poll, or manual) will replace this with Azure's exact
    // end time. Pending-approval activations are NOT flipped — they only
    // move to active once Azure provisions them.
    if (isImmediateActivationStatus(result.status)) {
      applyOptimisticActivation(role, durationHours);
    }
  } catch (error) {
    const raw = `Activation failed: ${String(error.message || error)}`;
    const human = getHumanActivationMessage(raw, { isError: true });
    setActivationStatusMessage(human, raw);
  }
}

// Returns true when Azure's response status indicates the activation has
// already taken effect (vs awaiting approval). Azure uses several synonyms
// for "granted" depending on the family — we accept any of them.
function isImmediateActivationStatus(status) {
  const normalized = String(status || '').toLowerCase();
  return (
    normalized.includes('granted') ||
    normalized.includes('provisioned') ||
    normalized.includes('accepted') ||
    normalized.includes('schedulecreated')
  );
}

// Mutates local state to flip the eligible role into an active role and
// re-renders both panels. Safe to call multiple times — finds the existing
// entry by id and only flips if it's currently eligible.
function applyOptimisticActivation(eligibleRole, durationHours) {
  const tenantKey = eligibleRole.tenantKey;
  const family = eligibleRole.family;
  const familyBucket = state.rolesByTenantFamily[tenantKey]?.[family];
  if (!Array.isArray(familyBucket)) {
    return;
  }

  const target = familyBucket.find((entry) => entry.id === eligibleRole.id);
  if (!target || target.state === 'active') {
    return;
  }

  const endTimeMs = Date.now() + Math.max(0, Number(durationHours) || 0) * 3600 * 1000;
  const endTimeIso = new Date(endTimeMs).toISOString();
  target.state = 'active';
  target.endTimeIso = endTimeIso;
  target.endTime = new Date(endTimeMs).toLocaleString();
  // Stamp the moment of the optimistic flip. mergeFamilyRoles uses this to
  // preserve the row across any Refresh that hits while Azure's listing
  // endpoint is still catching up (see OPTIMISTIC_GRACE_MS).
  target.optimisticActivatedAt = Date.now();

  renderActiveRoles();
  renderRoles();
}

// ── Login & tenant ──

export async function login() {
  const selectedTenantButton = dom.tenantPresetButtons.find((button) => button.dataset.tenantKey === selectedTenantKey);
  const selectedTenantLabel = selectedTenantButton?.textContent?.trim() || selectedTenantKey;
  dom.authStatus.textContent = 'Signing in... browser will open automatically.';

  try {
    const result = await getPimClient().login(selectedTenantKey);
    dom.authStatus.textContent = `Signed in as ${result.username} (${result.tenantId}). Pulling roles...`;
    await fetchAllRoles({ force: true });
    dom.authStatus.textContent = `Signed in as ${result.username} (${result.tenantId}) for ${selectedTenantLabel}`;
  } catch (error) {
    const message = String(error.message || error);
    dom.authStatus.textContent = `Sign-in failed: ${message}`;
  }
}

export function applyTenantSelection(tenantKey) {
  setSelectedTenantKey(tenantKey);

  for (const button of dom.tenantPresetButtons) {
    const isActive = button.dataset.tenantKey === tenantKey;
    button.classList.toggle('active', isActive);
    button.setAttribute('aria-pressed', String(isActive));
  }

  const active = dom.tenantPresetButtons.find((button) => button.dataset.tenantKey === tenantKey);
  if (dom.tenantStatus && active) {
    dom.tenantStatus.textContent = `Selected tenant: ${active.textContent?.trim() || tenantKey}`;
  }

  updateRoleTabsMeta();
  renderRoles();
  renderActiveRoles();
}

// ── Wiring (called from init) ──

export function wireRolesUI() {
  document.getElementById('loginBtn').addEventListener('click', login);

  dom.refreshBtn.addEventListener('click', () => {
    // Manual refresh is always silent — the user already has rows on screen
    // and shouldn't see a skeleton flash just to pick up any external changes.
    fetchAllRoles({ force: true, silent: true });
  });

  for (const button of dom.tenantPresetButtons) {
    button.addEventListener('click', () => {
      applyTenantSelection(button.dataset.tenantKey || 'nuance');
    });
  }

  if (dom.activationRawToggle) {
    dom.activationRawToggle.addEventListener('click', toggleActivationRawMessage);
    dom.activationRawToggle.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        toggleActivationRawMessage();
      }
    });
  }

  for (const tab of dom.familyTabs) {
    const activateTab = () => {
      setActiveFamily(tab.dataset.family || 'entra');

      for (const tabButton of dom.familyTabs) {
        const isActive = tabButton === tab;
        tabButton.classList.toggle('active', isActive);
        tabButton.setAttribute('aria-selected', String(isActive));
      }

      renderRoles();
      renderActiveRoles();

      if (!state.isRolesLoading && !hasAnyCachedRoles()) {
        fetchAllRoles({ force: false });
      }
    };

    tab.addEventListener('click', activateTab);
    tab.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        activateTab();
      }
    });
  }

  if (dom.durationHoursInput) {
    dom.durationHoursInput.addEventListener('input', () => {
      syncDurationControls(dom.durationHoursInput.value);
    });

    dom.durationHoursInput.addEventListener('pointerdown', () => {
      dom.durationHoursInput.classList.add('is-sliding');
    });

    const stopSliding = () => {
      dom.durationHoursInput.classList.remove('is-sliding');
    };

    dom.durationHoursInput.addEventListener('pointerup', stopSliding);
    dom.durationHoursInput.addEventListener('pointercancel', stopSliding);
    dom.durationHoursInput.addEventListener('blur', stopSliding);
  }

  if (dom.durationHoursBoxInput) {
    dom.durationHoursBoxInput.addEventListener('input', () => {
      syncDurationControls(dom.durationHoursBoxInput.value);
    });

    dom.durationHoursBoxInput.addEventListener('blur', () => {
      syncDurationControls(dom.durationHoursBoxInput.value);
    });
  }

  // Cleanup on window unload — clear all per-role expiry timers so they don't
  // fire after the renderer is torn down. There are no background pollers or
  // intervals to stop: this app is "user-initiated refresh only" by design,
  // so the only thing left to clean up is the per-role expiry setTimeout map.
  window.addEventListener('beforeunload', () => {
    for (const handle of state.expiryTimers.values()) {
      clearTimeout(handle);
    }
    state.expiryTimers.clear();
  });
}
