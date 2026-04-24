import { app } from 'electron';
import fs from 'fs';
import path from 'path';

// Persisted state for the in-app "What's new" dialog.
//
// The app auto-shows the changelog dialog ONCE per version: on first launch
// after an update we compare app.getVersion() to lastSeenVersion stored here.
// If they differ, we open the dialog and then write the current version back
// so we don't pester the user again until the next update.
//
// This file is unencrypted plaintext JSON — the only thing it stores is a
// version string, no secrets.
export interface ChangelogState {
  lastSeenVersion: string;
}

const DEFAULT_STATE: ChangelogState = {
  lastSeenVersion: ''
};

export class ChangelogStateStore {
  private readonly storePath = path.join(app.getPath('userData'), 'changelog-state.json');

  load(): ChangelogState {
    if (!fs.existsSync(this.storePath)) {
      return { ...DEFAULT_STATE };
    }

    try {
      const raw = fs.readFileSync(this.storePath, 'utf8');
      const parsed = JSON.parse(raw) as Partial<ChangelogState>;
      return {
        lastSeenVersion: typeof parsed.lastSeenVersion === 'string' ? parsed.lastSeenVersion : ''
      };
    } catch {
      // Corrupt or unreadable — treat as fresh install. Worst case the user
      // sees the changelog dialog one extra time. Not worth surfacing.
      return { ...DEFAULT_STATE };
    }
  }

  save(state: ChangelogState): void {
    try {
      fs.mkdirSync(path.dirname(this.storePath), { recursive: true });
      fs.writeFileSync(this.storePath, JSON.stringify(state, null, 2), 'utf8');
    } catch (error) {
      // Non-fatal — the dialog will just re-open next launch. Log so we can
      // notice if it ever happens repeatedly.
      console.error('ChangelogStateStore: failed to persist state', error);
    }
  }
}
