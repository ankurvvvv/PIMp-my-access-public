import { dom, selectedTenantKey } from './modules/state.js';
import {
  syncDurationControls, renderRoles, renderActiveRoles,
  applyTenantSelection, wireRolesUI
} from './modules/rolesUI.js';
import {
  initializeUpdateControls, renderDeveloperInfo, wireControlsUI
} from './modules/controlsUI.js';
import { wireChangelogUI, checkAndShowChangelogIfNew } from './modules/changelogUI.js';

wireRolesUI();
wireControlsUI();
wireChangelogUI();

syncDurationControls(dom.durationHoursInput ? dom.durationHoursInput.value : '1');
renderRoles();
renderActiveRoles();
renderDeveloperInfo();
initializeUpdateControls();
applyTenantSelection(selectedTenantKey);

// Fire-and-forget: if this is the first launch on a new version, auto-open
// the "What's new" dialog once. Catches its own errors — never blocks boot.
checkAndShowChangelogIfNew();
