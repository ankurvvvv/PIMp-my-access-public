export const state = {
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
  isRolesLoading: false,
  // Flipped true after the first successful post-login role fetch completes.
  // Used by fetchAllRoles to decide between skeleton (first load) and silent
  // (subsequent manual refresh) rendering modes.
  isFirstLoadDone: false,
  // Map<roleId, timeoutHandle> — one setTimeout per active role, scheduled to
  // fire at that role's endTimeIso so the UI flips to "Expired" instantly
  // without needing the user to click Refresh.
  expiryTimers: new Map()
};

export const ROLE_FAMILIES = ['entra', 'azureResource', 'group'];
export const TENANT_KEYS = ['nuance', 'healthcareCloud'];
export const ROLES_CACHE_TTL_MS = 90 * 1000;
// Grace window during which a locally-flipped (optimistically activated) role
// is preserved across a Refresh even if Azure's listing endpoint has not yet
// caught up with it. Azure ARM's roleAssignmentScheduleInstances?asTarget()
// view is eventually consistent and can lag 30–90s after a successful
// activation. Without this grace, a user who clicks Refresh immediately after
// activating an Azure role would see the card disappear.
export const OPTIMISTIC_GRACE_MS = 90 * 1000;

export const dom = {
  authStatus: document.getElementById('authStatus'),
  activationStatus: document.getElementById('activationStatus'),
  activationHumanMessage: document.getElementById('activationHumanMessage'),
  activationRawMessage: document.getElementById('activationRawMessage'),
  activationRawToggle: document.getElementById('activationRawToggle'),
  rolesBody: document.getElementById('rolesBody'),
  activeRolesList: document.getElementById('activeRolesList'),
  tenantPresetButtons: Array.from(document.querySelectorAll('.tenant-preset-btn[data-tenant-key]')),
  tenantStatus: document.getElementById('tenantStatus'),
  familyTabs: Array.from(document.querySelectorAll('.role-tab[data-family]')),
  rolesStatus: document.getElementById('rolesStatus'),
  activeRolesStatus: document.getElementById('activeRolesStatus'),
  rolesLoadingIndicator: document.getElementById('rolesLoadingIndicator'),
  refreshBtn: document.getElementById('refreshBtn'),
  checkUpdatesBtn: document.getElementById('checkUpdatesBtn'),
  restartUpdateBtn: document.getElementById('restartUpdateBtn'),
  updateStatus: document.getElementById('updateStatus'),
  updateMeta: document.getElementById('updateMeta'),
  rolesHeaderRow: document.getElementById('rolesHeaderRow'),
  settingsDrawer: document.getElementById('settingsDrawer'),
  settingsBackdrop: document.getElementById('settingsBackdrop'),
  settingsToggleBtn: document.getElementById('settingsToggleBtn'),
  updateToast: document.getElementById('updateToast'),
  updateToastMessage: document.getElementById('updateToastMessage'),
  updateToastAction: document.getElementById('updateToastAction'),
  updateToastDismiss: document.getElementById('updateToastDismiss'),
  settingsCloseBtn: document.getElementById('settingsCloseBtn'),
  windowMinBtn: document.getElementById('windowMinBtn'),
  windowMaxBtn: document.getElementById('windowMaxBtn'),
  windowCloseBtn: document.getElementById('windowCloseBtn'),
  durationHoursInput: document.getElementById('durationHours'),
  durationHoursBoxInput: document.getElementById('durationHoursBox'),
  durationHoursUnit: document.getElementById('durationHoursUnit'),
  developerName: document.getElementById('developerName'),
  developerOrg: document.getElementById('developerOrg'),
  developerEmail: document.getElementById('developerEmail'),
  developerRepo: document.getElementById('developerRepo'),
  developerBuildVersion: document.getElementById('developerBuildVersion'),
  watermarkVersion: document.getElementById('watermarkVersion'),
  // Changelog / "What's new" dialog elements.
  openChangelogBtn: document.getElementById('openChangelogBtn'),
  changelogBackdrop: document.getElementById('changelogBackdrop'),
  changelogModal: document.getElementById('changelogModal'),
  changelogBody: document.getElementById('changelogBody'),
  changelogCloseBtn: document.getElementById('changelogCloseBtn'),
  changelogDismissBtn: document.getElementById('changelogDismissBtn')
};

export let activeFamily = 'entra';
export let isActivationRawOpen = false;
export let updaterDiagnosticsWarning = '';
export let selectedTenantKey = 'nuance';

export function setActiveFamily(value) {
  activeFamily = value;
}

export function setActivationRawOpen(value) {
  isActivationRawOpen = value;
}

export function setUpdaterDiagnosticsWarning(value) {
  updaterDiagnosticsWarning = value;
}

export function setSelectedTenantKey(value) {
  selectedTenantKey = value;
}

export function getPimClient() {
  const client = window.pimClient;
  if (!client) {
    throw new Error('App bridge is unavailable. Please restart the app and try again.');
  }

  return client;
}

export function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
