import {
  dom, updaterDiagnosticsWarning, setUpdaterDiagnosticsWarning,
  getPimClient
} from './state.js';

const APP_DEVELOPER_INFO = {
  name: 'Ankur Vishwakarma',
  organization: 'Microsoft',
  email: 'ankurankur20m@gmail.com',
  repository: 'https://github.com/ankurvvvv'
};

// ── Update controls ──

function renderUpdateState(updateState) {
  if (!updateState) {
    return;
  }

  if (dom.updateStatus) {
    dom.updateStatus.textContent = updateState.message || 'Update status unavailable.';
  }

  if (dom.checkUpdatesBtn) {
    const isBusy = updateState.phase === 'checking' || updateState.phase === 'downloading';
    dom.checkUpdatesBtn.disabled = isBusy;
  }

  if (dom.restartUpdateBtn) {
    const canRestart = Boolean(updateState.canRestartToUpdate);
    dom.restartUpdateBtn.hidden = !canRestart;
    dom.restartUpdateBtn.disabled = !canRestart;
  }

  if (dom.updateToast) {
    if (updateState.phase === 'downloaded' && updateState.canRestartToUpdate) {
      if (dom.updateToastMessage) {
        dom.updateToastMessage.textContent = `Update ${updateState.version || ''} is ready.`;
      }
      dom.updateToast.hidden = false;
    } else {
      dom.updateToast.hidden = true;
    }
  }

  if (dom.updateMeta) {
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
      dom.updateMeta.textContent = `${summary} | ${updaterDiagnosticsWarning}`;
    } else {
      dom.updateMeta.textContent = summary || updaterDiagnosticsWarning;
    }
  }
}

export async function initializeUpdateControls() {
  const client = getPimClient();

  if (dom.checkUpdatesBtn) {
    dom.checkUpdatesBtn.addEventListener('click', async () => {
      dom.checkUpdatesBtn.disabled = true;
      if (dom.updateStatus) {
        dom.updateStatus.textContent = 'Checking for updates...';
      }

      try {
        const nextState = await client.checkForUpdates();
        renderUpdateState(nextState);
      } catch (error) {
        if (dom.updateStatus) {
          dom.updateStatus.textContent = `Update check failed: ${String(error.message || error)}`;
        }
      } finally {
        dom.checkUpdatesBtn.disabled = false;
      }
    });
  }

  if (dom.restartUpdateBtn) {
    dom.restartUpdateBtn.addEventListener('click', async () => {
      dom.restartUpdateBtn.disabled = true;
      if (dom.updateStatus) {
        dom.updateStatus.textContent = 'Restarting to install update...';
      }

      try {
        await client.restartToInstall();
      } catch (error) {
        dom.restartUpdateBtn.disabled = false;
        if (dom.updateStatus) {
          dom.updateStatus.textContent = `Unable to restart for update: ${String(error.message || error)}`;
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
      setUpdaterDiagnosticsWarning(warningParts.join(' | '));
    }

    const state = await client.getUpdateState();
    renderUpdateState(state);
  } catch (error) {
    if (dom.updateStatus) {
      dom.updateStatus.textContent = `Unable to read update status: ${String(error.message || error)}`;
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

// ── Developer info ──

export function renderDeveloperInfo() {
  if (dom.developerName) {
    dom.developerName.textContent = APP_DEVELOPER_INFO.name;
  }

  if (dom.developerOrg) {
    dom.developerOrg.textContent = APP_DEVELOPER_INFO.organization;
  }

  if (dom.developerEmail) {
    dom.developerEmail.textContent = APP_DEVELOPER_INFO.email;
    dom.developerEmail.href = `mailto:${APP_DEVELOPER_INFO.email}`;
  }

  if (dom.developerRepo) {
    dom.developerRepo.textContent = APP_DEVELOPER_INFO.repository;
    dom.developerRepo.href = APP_DEVELOPER_INFO.repository;
  }

  if (dom.developerBuildVersion) {
    dom.developerBuildVersion.textContent = 'Loading...';
    getPimClient()
      .getAppVersion()
      .then((version) => {
        dom.developerBuildVersion.textContent = version;
        if (dom.watermarkVersion) dom.watermarkVersion.textContent = `v${version}`;
      })
      .catch(() => {
        dom.developerBuildVersion.textContent = 'Unknown';
      });
  }
}

// ── Settings drawer ──

function setSettingsDrawerOpen(open) {
  dom.settingsDrawer.classList.toggle('open', open);
  dom.settingsBackdrop.classList.toggle('open', open);
  dom.settingsDrawer.setAttribute('aria-hidden', String(!open));
  dom.settingsBackdrop.setAttribute('aria-hidden', String(!open));
}

// ── Wiring (called from init) ──

export function wireControlsUI() {
  if (dom.updateToastAction) {
    dom.updateToastAction.addEventListener('click', async () => {
      try {
        await getPimClient().restartToInstall();
      } catch {}
    });
  }

  if (dom.updateToastDismiss) {
    dom.updateToastDismiss.addEventListener('click', () => {
      if (dom.updateToast) dom.updateToast.hidden = true;
    });
  }

  dom.settingsToggleBtn.addEventListener('click', () => {
    const isOpen = dom.settingsDrawer.classList.contains('open');
    setSettingsDrawerOpen(!isOpen);
  });

  dom.settingsCloseBtn.addEventListener('click', () => {
    setSettingsDrawerOpen(false);
  });

  dom.settingsBackdrop.addEventListener('click', () => {
    setSettingsDrawerOpen(false);
  });

  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      setSettingsDrawerOpen(false);
    }
  });

  dom.windowMinBtn.addEventListener('click', async () => {
    await getPimClient().windowMinimize();
  });

  dom.windowMaxBtn.addEventListener('click', async () => {
    await getPimClient().windowToggleMaximize();
  });

  dom.windowCloseBtn.addEventListener('click', async () => {
    await getPimClient().windowClose();
  });
}
