import { dom, getPimClient, escapeHtml } from './state.js';

// ── In-app "What's new" dialog ──
//
// Reads CHANGELOG.md (parsed in main process) and renders it in a modal.
// Auto-opens once after every version bump; can also be opened on demand
// from the Settings drawer button.
//
// Why parsing happens in main and not here: keeps the renderer code small and
// avoids shipping a markdown library. We get clean structured data over IPC.

let changelogPayloadCache = null;

function setChangelogOpen(open) {
  if (!dom.changelogModal || !dom.changelogBackdrop) {
    return;
  }
  dom.changelogModal.hidden = !open;
  dom.changelogBackdrop.hidden = !open;
  dom.changelogModal.classList.toggle('open', open);
  dom.changelogBackdrop.classList.toggle('open', open);

  if (open && dom.changelogDismissBtn) {
    // Move keyboard focus into the dialog so Enter/Esc work intuitively and
    // assistive tech announces the dialog title.
    dom.changelogDismissBtn.focus();
  }
}

// Lightweight inline markdown -> safe HTML for bullet text only.
// Supports: **bold**, `code`, [text](https://...) — nothing else.
// Everything is HTML-escaped FIRST, then we re-introduce the four allowed
// tags. This prevents any injection from malformed CHANGELOG entries.
function renderInlineMarkdown(rawText) {
  let html = escapeHtml(rawText);

  // **bold**
  html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');

  // `code`
  html = html.replace(/`([^`]+)`/g, '<code>$1</code>');

  // [text](https://url) — only http/https links, nothing else.
  // We add a data-external-url attribute and intercept clicks to route
  // through preload's openExternal (the renderer is sandboxed and cannot
  // open URLs directly).
  html = html.replace(
    /\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g,
    (_match, text, url) =>
      `<a href="#" class="changelog-link" data-external-url="${escapeHtml(url)}">${text}</a>`
  );

  return html;
}

function renderChangelogPayload(payload) {
  if (!dom.changelogBody) {
    return;
  }

  if (!payload || !payload.found) {
    dom.changelogBody.innerHTML = `
      <p class="changelog-empty">CHANGELOG.md was not found in this build.</p>
    `;
    return;
  }

  if (!Array.isArray(payload.entries) || payload.entries.length === 0) {
    dom.changelogBody.innerHTML = `
      <p class="changelog-empty">No release notes are available yet.</p>
    `;
    return;
  }

  const sections = [];
  for (const entry of payload.entries) {
    const isCurrent = entry.version === payload.currentVersion;
    const headerSuffix = isCurrent ? ' <span class="changelog-current-badge">current</span>' : '';
    const entryHtml = [
      `<section class="changelog-entry">`,
      `  <h3 class="changelog-entry-title">${escapeHtml(entry.version)} <span class="changelog-entry-date">${escapeHtml(entry.date)}</span>${headerSuffix}</h3>`
    ];

    for (const section of entry.sections) {
      entryHtml.push(`  <h4 class="changelog-section-heading">${escapeHtml(section.heading)}</h4>`);
      entryHtml.push('  <ul class="changelog-section-list">');
      for (const item of section.items) {
        entryHtml.push(`    <li>${renderInlineMarkdown(item)}</li>`);
      }
      entryHtml.push('  </ul>');
    }

    entryHtml.push('</section>');
    sections.push(entryHtml.join('\n'));
  }

  dom.changelogBody.innerHTML = sections.join('\n');

  // Wire up external-link clicks now that the DOM exists.
  for (const link of dom.changelogBody.querySelectorAll('.changelog-link')) {
    link.addEventListener('click', async (event) => {
      event.preventDefault();
      const url = link.getAttribute('data-external-url');
      if (!url) {
        return;
      }
      try {
        await getPimClient().openExternal(url);
      } catch {
        // Non-fatal — the link just won't open. No user-facing toast for this
        // edge case to keep the dialog quiet.
      }
    });
  }
}

async function loadChangelogPayload() {
  if (changelogPayloadCache) {
    return changelogPayloadCache;
  }
  try {
    const payload = await getPimClient().getChangelog();
    changelogPayloadCache = payload;
    return payload;
  } catch (error) {
    console.error('Failed to load changelog payload', error);
    return { entries: [], found: false, sourcePath: '', currentVersion: '' };
  }
}

export async function openChangelogDialog() {
  const payload = await loadChangelogPayload();
  renderChangelogPayload(payload);
  setChangelogOpen(true);

  // Mark the current version as "seen" so the auto-popup will not fire again
  // until the next version bump. We do this on OPEN (not on close) so that
  // even if the user force-quits the app without dismissing, we still don't
  // pester them again on next launch.
  if (payload && payload.currentVersion) {
    try {
      await getPimClient().markChangelogSeen(payload.currentVersion);
    } catch {
      // Non-fatal — worst case the dialog opens once more on next launch.
    }
  }
}

function closeChangelogDialog() {
  setChangelogOpen(false);
}

export function wireChangelogUI() {
  if (dom.openChangelogBtn) {
    dom.openChangelogBtn.addEventListener('click', () => {
      openChangelogDialog();
    });
  }

  if (dom.changelogCloseBtn) {
    dom.changelogCloseBtn.addEventListener('click', closeChangelogDialog);
  }

  if (dom.changelogDismissBtn) {
    dom.changelogDismissBtn.addEventListener('click', closeChangelogDialog);
  }

  if (dom.changelogBackdrop) {
    dom.changelogBackdrop.addEventListener('click', closeChangelogDialog);
  }

  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && dom.changelogModal && !dom.changelogModal.hidden) {
      closeChangelogDialog();
    }
  });
}

// Auto-popup on first launch after a version bump.
//
// Logic: ask main for the current app version and the last "seen" version.
// If they differ AND there are real entries to show, open the dialog. The
// dialog itself writes the "seen" marker so we don't re-trigger.
//
// Defensive: never throw out of this function — it runs on app boot and a
// failure here must not block the rest of the app from rendering.
export async function checkAndShowChangelogIfNew() {
  try {
    const [payload, lastSeen] = await Promise.all([
      loadChangelogPayload(),
      getPimClient().getLastSeenChangelogVersion()
    ]);

    if (!payload || !payload.found || !payload.currentVersion) {
      return;
    }
    if (payload.entries.length === 0) {
      return;
    }
    if (lastSeen === payload.currentVersion) {
      return;
    }

    await openChangelogDialog();
  } catch (error) {
    console.error('checkAndShowChangelogIfNew failed', error);
  }
}
