const state = {
  roles: [],
  selectedRoleId: null,
  rolesByTenantFamily: {
    nuance: {
      entra: [],
      azureResource: [],
      group: []
    },
    healthcareCloud: {
      entra: [],
      azureResource: [],
      group: []
    }
  },
  familyStatusByTenant: {
    nuance: {
      entra: 'idle',
      azureResource: 'idle',
      group: 'idle'
    },
    healthcareCloud: {
      entra: 'idle',
      azureResource: 'idle',
      group: 'idle'
    }
  },
  rolesCacheUntil: 0,
  activeRolesLoadToken: 0,
  isRolesLoading: false
};

const ROLE_FAMILIES = ['entra', 'azureResource', 'group'];
const TENANT_KEYS = ['nuance', 'healthcareCloud'];
const ROLES_CACHE_TTL_MS = 90 * 1000;

const authStatus = document.getElementById('authStatus');
const activationStatus = document.getElementById('activationStatus');
const activationHumanMessage = document.getElementById('activationHumanMessage');
const activationRawMessage = document.getElementById('activationRawMessage');
const activationRawToggle = document.getElementById('activationRawToggle');
const rolesBody = document.getElementById('rolesBody');
const activeRolesList = document.getElementById('activeRolesList');
const tenantPresetButtons = Array.from(document.querySelectorAll('.tenant-preset-btn[data-tenant-key]'));
const tenantStatus = document.getElementById('tenantStatus');
const familyTabs = Array.from(document.querySelectorAll('.role-tab[data-family]'));
const rolesStatus = document.getElementById('rolesStatus');
const activeRolesStatus = document.getElementById('activeRolesStatus');
const rolesLoadingIndicator = document.getElementById('rolesLoadingIndicator');
const refreshBtn = document.getElementById('refreshBtn');
const checkUpdatesBtn = document.getElementById('checkUpdatesBtn');
const restartUpdateBtn = document.getElementById('restartUpdateBtn');
const updateStatus = document.getElementById('updateStatus');
const updateMeta = document.getElementById('updateMeta');
const rolesHeaderRow = document.getElementById('rolesHeaderRow');
const settingsDrawer = document.getElementById('settingsDrawer');
const settingsBackdrop = document.getElementById('settingsBackdrop');
const settingsToggleBtn = document.getElementById('settingsToggleBtn');
const updateToast = document.getElementById('updateToast');
const updateToastMessage = document.getElementById('updateToastMessage');
const updateToastAction = document.getElementById('updateToastAction');
const updateToastDismiss = document.getElementById('updateToastDismiss');
const settingsCloseBtn = document.getElementById('settingsCloseBtn');
const windowMinBtn = document.getElementById('windowMinBtn');
const windowMaxBtn = document.getElementById('windowMaxBtn');
const windowCloseBtn = document.getElementById('windowCloseBtn');
const durationHoursInput = document.getElementById('durationHours');
const durationHoursBoxInput = document.getElementById('durationHoursBox');
const durationHoursUnit = document.getElementById('durationHoursUnit');
const developerName = document.getElementById('developerName');
const developerOrg = document.getElementById('developerOrg');
const developerEmail = document.getElementById('developerEmail');
const developerRepo = document.getElementById('developerRepo');
const developerBuildVersion = document.getElementById('developerBuildVersion');
const watermarkVersion = document.getElementById('watermarkVersion');

const APP_DEVELOPER_INFO = {
  name: 'Ankur Vishwakarma',
  organization: 'Microsoft',
  email: 'ankurankur20m@gmail.com',
  repository: 'https://github.com/ankurvvvv'
};

let activeFamily = 'entra';
let isActivationRawOpen = false;
let updaterDiagnosticsWarning = '';
let selectedTenantKey = 'nuance';

// NOTE: Managed/BYO/tenant override sign-in flows are intentionally disabled for now.
// Keep only one-click Azure sign-in UX until public-release auth model is finalized.

function renderUpdateState(updateState) {
  if (!updateState) {
    return;
  }

  if (updateStatus) {
    updateStatus.textContent = updateState.message || 'Update status unavailable.';
  }

  if (checkUpdatesBtn) {
    const isBusy = updateState.phase === 'checking' || updateState.phase === 'downloading';
    checkUpdatesBtn.disabled = isBusy;
  }

  if (restartUpdateBtn) {
    const canRestart = Boolean(updateState.canRestartToUpdate);
    restartUpdateBtn.hidden = !canRestart;
    restartUpdateBtn.disabled = !canRestart;
  }

  if (updateToast) {
    if (updateState.phase === 'downloaded' && updateState.canRestartToUpdate) {
      if (updateToastMessage) {
        updateToastMessage.textContent = `Update ${updateState.version || ''} is ready.`;
      }
      updateToast.hidden = false;
    } else {
      updateToast.hidden = true;
    }
  }

  if (updateMeta) {
    const details = [];
    if (updateState.channel) {
      details.push(`channel: ${updateState.channel}`);
    }
    if (updateState.checkedAt) {
      details.push(`last check: ${new Date(updateState.checkedAt).toLocaleString()}`);
    }
    if (updateState.nextCheckAt) {
      details.push(`next check: ${new Date(updateState.nextCheckAt).toLocaleString()}`);
    }
    const summary = details.join(' | ');
    if (summary && updaterDiagnosticsWarning) {
      updateMeta.textContent = `${summary} | ${updaterDiagnosticsWarning}`;
    } else {
      updateMeta.textContent = summary || updaterDiagnosticsWarning;
    }
  }
}

async function initializeUpdateControls() {
  const client = getPimClient();

  if (checkUpdatesBtn) {
    checkUpdatesBtn.addEventListener('click', async () => {
      checkUpdatesBtn.disabled = true;
      if (updateStatus) {
        updateStatus.textContent = 'Checking for updates...';
      }

      try {
        const nextState = await client.checkForUpdates();
        renderUpdateState(nextState);
      } catch (error) {
        if (updateStatus) {
          updateStatus.textContent = `Update check failed: ${String(error.message || error)}`;
        }
      } finally {
        checkUpdatesBtn.disabled = false;
      }
    });
  }

  if (restartUpdateBtn) {
    restartUpdateBtn.addEventListener('click', async () => {
      restartUpdateBtn.disabled = true;
      if (updateStatus) {
        updateStatus.textContent = 'Restarting to install update...';
      }

      try {
        await client.restartToInstall();
      } catch (error) {
        restartUpdateBtn.disabled = false;
        if (updateStatus) {
          updateStatus.textContent = `Unable to restart for update: ${String(error.message || error)}`;
        }
      }
    });
  }

  try {
    const diagnostics = await client.getUpdateDiagnostics();
    if (diagnostics) {
      const warningParts = [];
      if (!diagnostics.metadataValid && diagnostics.metadataIssue) {
        warningParts.push(diagnostics.metadataIssue);
      }
      if (Array.isArray(diagnostics.warnings) && diagnostics.warnings.length > 0) {
        warningParts.push(...diagnostics.warnings);
      }
      updaterDiagnosticsWarning = warningParts.join(' | ');
    }

    const state = await client.getUpdateState();
    renderUpdateState(state);
  } catch (error) {
    if (updateStatus) {
      updateStatus.textContent = `Unable to read update status: ${String(error.message || error)}`;
    }
  }

  const unsubscribe = client.onUpdateState((state) => {
    renderUpdateState(state);
  });

  window.addEventListener(
    'beforeunload',
    () => {
      unsubscribe();
    },
    { once: true }
  );
}

function renderDeveloperInfo() {
  if (developerName) {
    developerName.textContent = APP_DEVELOPER_INFO.name;
  }

  if (developerOrg) {
    developerOrg.textContent = APP_DEVELOPER_INFO.organization;
  }

  if (developerEmail) {
    developerEmail.textContent = APP_DEVELOPER_INFO.email;
    developerEmail.href = `mailto:${APP_DEVELOPER_INFO.email}`;
  }

  if (developerRepo) {
    developerRepo.textContent = APP_DEVELOPER_INFO.repository;
    developerRepo.href = APP_DEVELOPER_INFO.repository;
  }

  if (developerBuildVersion) {
    developerBuildVersion.textContent = 'Loading...';
    getPimClient()
      .getAppVersion()
      .then((version) => {
        developerBuildVersion.textContent = version;
        if (watermarkVersion) watermarkVersion.textContent = `v${version}`;
      })
      .catch(() => {
        developerBuildVersion.textContent = 'Unknown';
      });
  }
}

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

function setActivationStatusMessage(humanMessage, rawMessage = '') {
  if (!activationStatus || !activationHumanMessage || !activationRawMessage || !activationRawToggle) {
    return;
  }

  activationHumanMessage.textContent = humanMessage;
  activationRawMessage.textContent = rawMessage;
  isActivationRawOpen = false;

  const hasRawMessage = Boolean(rawMessage);
  activationRawToggle.hidden = !hasRawMessage;
  activationRawMessage.hidden = true;
  activationRawToggle.setAttribute('aria-expanded', 'false');
  activationRawToggle.innerHTML = '&#9656;';
}

function toggleActivationRawMessage() {
  if (!activationRawMessage || !activationRawToggle || activationRawToggle.hidden) {
    return;
  }

  isActivationRawOpen = !isActivationRawOpen;
  activationRawMessage.hidden = !isActivationRawOpen;
  activationRawToggle.setAttribute('aria-expanded', String(isActivationRawOpen));
  activationRawToggle.innerHTML = isActivationRawOpen ? '&#9662;' : '&#9656;';
}

function setRolesLoading(isLoading) {
  state.isRolesLoading = isLoading;
  if (rolesLoadingIndicator) {
    rolesLoadingIndicator.classList.toggle('active', isLoading);
    rolesLoadingIndicator.setAttribute('aria-hidden', String(!isLoading));
  }

  if (refreshBtn) {
    refreshBtn.disabled = isLoading;
  }
}

function setRolesStatus(message) {
  if (rolesStatus) {
    rolesStatus.textContent = message;
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
  const tenantStatus = state.familyStatusByTenant[selectedTenantKey] || state.familyStatusByTenant.nuance;

  for (const tab of familyTabs) {
    const family = tab.dataset.family;
    if (!family) {
      continue;
    }

    const count = tenantRoles[family].length;
    const status = tenantStatus[family];
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
  rolesBody.innerHTML = '';
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
    rolesBody.appendChild(row);
  }
}

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
  if (getActiveFamily() === family && selectedTenantKey === tenantKey) {
    renderRoles();
    renderActiveRoles();
  }
}

async function fetchAllRoles(options = {}) {
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
  if (!durationHoursInput) {
    return;
  }

  const minValue = Number(durationHoursInput.min || '0.5');
  const maxValue = Number(durationHoursInput.max || '8');
  const range = maxValue - minValue;
  const percentage = range > 0 ? ((durationHours - minValue) / range) * 100 : 0;
  durationHoursInput.style.setProperty('--slider-fill', `${percentage}%`);
}

function syncDurationControls(rawValue) {
  const durationHours = normalizeDurationHours(rawValue);

  if (durationHoursInput) {
    durationHoursInput.value = durationHours.toString();
  }

  if (durationHoursBoxInput) {
    durationHoursBoxInput.value = durationHours.toFixed(1).replace('.0', '');
  }

  if (durationHoursUnit) {
    durationHoursUnit.textContent = getDurationUnit(durationHours);
  }

  updateDurationSliderFill(durationHours);
}

function getPimClient() {
  const client = window.pimClient;
  if (!client) {
    throw new Error('App bridge is unavailable. Please restart the app and try again.');
  }

  return client;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

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

function getActiveFamily() {
  return activeFamily;
}

function renderHeaders() {
  const activeFamily = getActiveFamily();

  if (activeFamily === 'azureResource') {
    rolesHeaderRow.innerHTML = `
      <th>Role Name</th>
      <th>Resource</th>
      <th>Resource type</th>
      <th>End time</th>
      <th>Activate</th>
    `;
    return;
  }

  rolesHeaderRow.innerHTML = `
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

function setSettingsDrawerOpen(open) {
  settingsDrawer.classList.toggle('open', open);
  settingsBackdrop.classList.toggle('open', open);
  settingsDrawer.setAttribute('aria-hidden', String(!open));
  settingsBackdrop.setAttribute('aria-hidden', String(!open));
}

async function login() {
  const selectedTenantButton = tenantPresetButtons.find((button) => button.dataset.tenantKey === selectedTenantKey);
  const selectedTenantLabel = selectedTenantButton?.textContent?.trim() || selectedTenantKey;
  authStatus.textContent = 'Signing in... browser will open automatically.';

  try {
    const result = await getPimClient().login(selectedTenantKey);
    authStatus.textContent = `Signed in as ${result.username} (${result.tenantId}). Pulling roles...`;
    await fetchAllRoles({ force: true });
    authStatus.textContent = `Signed in as ${result.username} (${result.tenantId}) for ${selectedTenantLabel}`;
  } catch (error) {
    const message = String(error.message || error);
    authStatus.textContent = `Sign-in failed: ${message}`;
  }
}

function applyTenantSelection(tenantKey) {
  selectedTenantKey = tenantKey;

  for (const button of tenantPresetButtons) {
    const isActive = button.dataset.tenantKey === tenantKey;
    button.classList.toggle('active', isActive);
    button.setAttribute('aria-pressed', String(isActive));
  }

  const active = tenantPresetButtons.find((button) => button.dataset.tenantKey === tenantKey);
  if (tenantStatus && active) {
    tenantStatus.textContent = `Selected tenant: ${active.textContent?.trim() || tenantKey}`;
  }

  updateRoleTabsMeta();
  renderRoles();
  renderActiveRoles();
}

function renderRoles() {
  renderHeaders();
  if (!rolesBody) {
    return;
  }

  rolesBody.innerHTML = '';
  const activeFamily = getActiveFamily();
  state.roles = (state.rolesByTenantFamily[selectedTenantKey]?.[activeFamily] || []).filter((role) => role.state !== 'active');

  if (state.isRolesLoading && state.roles.length === 0) {
    renderLoadingRows();
    return;
  }

  if (state.roles.length === 0) {
    rolesBody.innerHTML = '<tr><td colspan="5" class="empty-roles">No eligible roles found for this category.</td></tr>';
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

    rolesBody.appendChild(row);
  }
}

function renderActiveRoles() {
  if (!activeRolesList) {
    return;
  }

  activeRolesList.innerHTML = '';
  const groupedActiveRoles = ROLE_FAMILIES.map((family) => ({
    family,
    roles: (state.rolesByTenantFamily[selectedTenantKey]?.[family] || []).filter((role) => role.state === 'active')
  }));
  const totalActiveCount = groupedActiveRoles.reduce((count, entry) => count + entry.roles.length, 0);

  if (state.isRolesLoading && totalActiveCount === 0) {
    activeRolesList.innerHTML = '<div class="empty-roles">Checking active roles...</div>';
    if (activeRolesStatus) {
      activeRolesStatus.textContent = 'Checking active roles...';
    }
    return;
  }

  if (totalActiveCount === 0) {
    activeRolesList.innerHTML = '<div class="empty-roles">No active roles in any category.</div>';
    if (activeRolesStatus) {
      activeRolesStatus.textContent = 'No active roles.';
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
    activeRolesList.appendChild(groupElement);
  }

  if (activeRolesStatus) {
    activeRolesStatus.textContent = `${totalActiveCount} active role(s).`;
  }
}

async function activateRole(roleId) {
  const role = state.roles.find((item) => item.id === roleId);
  if (!role) {
    setActivationStatusMessage('Please check the input values');
    return;
  }

  const durationHours = normalizeDurationHours(durationHoursInput ? durationHoursInput.value : '1');
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

document.getElementById('loginBtn').addEventListener('click', login);
refreshBtn.addEventListener('click', () => {
  fetchAllRoles({ force: true });
});
for (const button of tenantPresetButtons) {
  button.addEventListener('click', () => {
    applyTenantSelection(button.dataset.tenantKey || 'nuance');
  });
}
if (activationRawToggle) {
  activationRawToggle.addEventListener('click', toggleActivationRawMessage);
  activationRawToggle.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      toggleActivationRawMessage();
    }
  });
}
for (const tab of familyTabs) {
  const activateTab = () => {
    activeFamily = tab.dataset.family || 'entra';

    for (const tabButton of familyTabs) {
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
if (durationHoursInput) {
  durationHoursInput.addEventListener('input', () => {
    syncDurationControls(durationHoursInput.value);
  });

  durationHoursInput.addEventListener('pointerdown', () => {
    durationHoursInput.classList.add('is-sliding');
  });

  const stopSliding = () => {
    durationHoursInput.classList.remove('is-sliding');
  };

  durationHoursInput.addEventListener('pointerup', stopSliding);
  durationHoursInput.addEventListener('pointercancel', stopSliding);
  durationHoursInput.addEventListener('blur', stopSliding);
}

if (durationHoursBoxInput) {
  durationHoursBoxInput.addEventListener('input', () => {
    syncDurationControls(durationHoursBoxInput.value);
  });

  durationHoursBoxInput.addEventListener('blur', () => {
    syncDurationControls(durationHoursBoxInput.value);
  });
}

syncDurationControls(durationHoursInput ? durationHoursInput.value : '1');
updateRoleTabsMeta();
renderRoles();
renderActiveRoles();
renderDeveloperInfo();
initializeUpdateControls();
applyTenantSelection(selectedTenantKey);

if (updateToastAction) {
  updateToastAction.addEventListener('click', async () => {
    try {
      await getPimClient().restartToInstall();
    } catch {}
  });
}

if (updateToastDismiss) {
  updateToastDismiss.addEventListener('click', () => {
    if (updateToast) updateToast.hidden = true;
  });
}

settingsToggleBtn.addEventListener('click', () => {
  const isOpen = settingsDrawer.classList.contains('open');
  setSettingsDrawerOpen(!isOpen);
});

settingsCloseBtn.addEventListener('click', () => {
  setSettingsDrawerOpen(false);
});

settingsBackdrop.addEventListener('click', () => {
  setSettingsDrawerOpen(false);
});

document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') {
    setSettingsDrawerOpen(false);
  }
});

windowMinBtn.addEventListener('click', async () => {
  await getPimClient().windowMinimize();
});

windowMaxBtn.addEventListener('click', async () => {
  await getPimClient().windowToggleMaximize();
});

windowCloseBtn.addEventListener('click', async () => {
  await getPimClient().windowClose();
});
