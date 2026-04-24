import { app } from 'electron';
import fs from 'fs';
import path from 'path';

// Structured representation of the project's CHANGELOG.md, designed to be
// rendered by the renderer's "What's new" dialog WITHOUT shipping a markdown
// parser to the renderer. We do the parsing here in main and ship clean JSON.
//
// CHANGELOG.md must follow Keep-A-Changelog conventions:
//   ## [VERSION] - YYYY-MM-DD
//   ### Added | Changed | Fixed | Removed
//   - bullet line
//
// Anything that doesn't fit this pattern is ignored — we are not building a
// general markdown parser.
export interface ChangelogSection {
  /** "Added" | "Changed" | "Fixed" | "Removed" — kept verbatim from heading. */
  heading: string;
  /** Bullet list items. Inline markdown (**bold**, `code`) is preserved as plain text in this version. */
  items: string[];
}

export interface ChangelogEntry {
  version: string;
  date: string;
  sections: ChangelogSection[];
}

export interface ChangelogPayload {
  entries: ChangelogEntry[];
  /** True if the on-disk CHANGELOG.md was found and parsed; false on miss. */
  found: boolean;
  /** Resolved file path used for the read. Useful for diagnostics. */
  sourcePath: string;
}

function resolveChangelogPath(): string {
  // app.getAppPath() points to the project root in dev and to the asar root
  // in packaged builds. CHANGELOG.md lives at the repo root and is bundled
  // because electron-builder's `files: ["**/*", "!misc/**/*"]` includes it.
  return path.join(app.getAppPath(), 'CHANGELOG.md');
}

function parseChangelog(markdown: string): ChangelogEntry[] {
  const lines = markdown.split(/\r?\n/);
  const entries: ChangelogEntry[] = [];
  let currentEntry: ChangelogEntry | null = null;
  let currentSection: ChangelogSection | null = null;

  // Match: ## [2026.4.3] - 2026-04-24
  // Tolerant of whitespace and an optional `v` prefix on the version.
  const entryRegex = /^##\s+\[v?([^\]]+)\]\s*-\s*(\d{4}-\d{2}-\d{2})\s*$/i;
  // Match: ### Added (and only the four valid headings)
  const sectionRegex = /^###\s+(Added|Changed|Fixed|Removed)\s*$/i;
  // Match: - item   OR   * item
  const bulletRegex = /^\s*[-*]\s+(.*\S)\s*$/;

  for (const rawLine of lines) {
    const entryMatch = entryRegex.exec(rawLine);
    if (entryMatch) {
      currentEntry = {
        version: entryMatch[1].trim(),
        date: entryMatch[2].trim(),
        sections: []
      };
      currentSection = null;
      entries.push(currentEntry);
      continue;
    }

    if (!currentEntry) {
      continue;
    }

    const sectionMatch = sectionRegex.exec(rawLine);
    if (sectionMatch) {
      currentSection = {
        heading: capitalize(sectionMatch[1]),
        items: []
      };
      currentEntry.sections.push(currentSection);
      continue;
    }

    const bulletMatch = bulletRegex.exec(rawLine);
    if (bulletMatch && currentSection) {
      currentSection.items.push(bulletMatch[1]);
    }
  }

  // Drop empty sections and entries that ended up with nothing — stricter
  // than necessary, but keeps the renderer's job trivial.
  for (const entry of entries) {
    entry.sections = entry.sections.filter((section) => section.items.length > 0);
  }

  return entries.filter((entry) => entry.sections.length > 0);
}

function capitalize(value: string): string {
  if (value.length === 0) {
    return value;
  }
  return value[0].toUpperCase() + value.slice(1).toLowerCase();
}

export function readChangelog(): ChangelogPayload {
  const sourcePath = resolveChangelogPath();
  if (!fs.existsSync(sourcePath)) {
    return { entries: [], found: false, sourcePath };
  }

  try {
    const raw = fs.readFileSync(sourcePath, 'utf8');
    const entries = parseChangelog(raw);
    return { entries, found: true, sourcePath };
  } catch (error) {
    console.error('readChangelog: failed to read or parse CHANGELOG.md', { sourcePath, error });
    return { entries: [], found: false, sourcePath };
  }
}
