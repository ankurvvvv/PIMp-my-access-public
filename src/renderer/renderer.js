import { dom, selectedTenantKey } from './modules/state.js';
import {
  syncDurationControls, renderRoles, renderActiveRoles,
  applyTenantSelection, wireRolesUI
} from './modules/rolesUI.js';
import {
  initializeUpdateControls, renderDeveloperInfo, wireControlsUI
} from './modules/controlsUI.js';

wireRolesUI();
wireControlsUI();

syncDurationControls(dom.durationHoursInput ? dom.durationHoursInput.value : '1');
renderRoles();
renderActiveRoles();
renderDeveloperInfo();
initializeUpdateControls();
applyTenantSelection(selectedTenantKey);
