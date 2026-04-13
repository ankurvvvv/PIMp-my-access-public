import {
  state, dom, ROLE_FAMILIES, TENANT_KEYS, ROLES_CACHE_TTL_MS,
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

    state.rolesByTenantFamily[tenantKey][family] = roles;
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

export async function fetchAllRoles(options = {}) {
  const force = Boolean(options.force);
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
  setRolesLoading(true);

  for (const tenantKey of TENANT_KEYS) {
    for (const family of ROLE_FAMILIES) {
      state.familyStatusByTenant[tenantKey][family] = 'loading';
    }
  }

  updateRoleTabsMeta();
  renderHeaders();
  renderLoadingRows();
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

  setRolesLoading(false);

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
    dom.activeRolesList.innerHTML = '<div class="empty-roles">No active roles in any category.</div>';
    if (dom.activeRolesStatus) {
      dom.activeRolesStatus.textContent = 'No active roles.';
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

      item.innerHTML = `
        <div class="active-role-topline">
          <p class="active-role-name">${escapeHtml(role.displayName)}</p>
          ${badgeMarkup}
        </div>
        <p class="active-role-context">${escapeHtml(getActiveRoleContextLine(role))}</p>
        <p class="active-role-endtime">End time: ${escapeHtml(role.endTime)}</p>
      `;

      listElement.appendChild(item);
    }

    groupElement.appendChild(listElement);
    dom.activeRolesList.appendChild(groupElement);
  }

  if (dom.activeRolesStatus) {
    dom.activeRolesStatus.textContent = `${totalActiveCount} active role(s).`;
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
  } catch (error) {
    const raw = `Activation failed: ${String(error.message || error)}`;
    const human = getHumanActivationMessage(raw, { isError: true });
    setActivationStatusMessage(human, raw);
  }
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
    fetchAllRoles({ force: true });
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
}
