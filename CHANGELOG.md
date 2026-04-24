# Changelog

All notable user-visible changes to **PIMp my access** are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and this project uses [CalVer](https://calver.org/) (`YYYY.M.counter`).

## [2026.4.3] - 2026-04-24

### Added
- **What's new** dialog inside the app (Settings drawer → "What's new"). Auto-opens once after every update so you can see what changed without leaving the app.
- In-app `CHANGELOG.md` is now bundled with every release and rendered inside the app.

### Changed
- **Refresh is now silent.** Manual refresh no longer flashes skeleton placeholders. Existing rows stay on screen, the Refresh button briefly shows "Syncing…", and only changed rows update. The skeleton is still used on the very first load after sign-in.
- **Active Roles cards** — role name and details are now left-aligned (previously centered when wrapped).
- **Deactivate button** — pressed/raised state now uses red shadow to match the destructive border (previously bled blue from the global button style).
- Product display name normalized to **PIMp my access** across window title, tray, brand label, and installer.

### Fixed
- Activated Azure resource roles no longer disappear from the **Active Roles** panel when you click Refresh within ~90 seconds of activating them. The app now keeps your just-activated roles visible during Azure ARM's eventual-consistency window and only drops them once the backend confirms they are gone.
- Long active role names (e.g. "Identity Governance Administrator") no longer center-align mid-card.

## [2026.4.2] - 2026-04-21

### Added
- Active Azure resource roles now surface correctly in the **Active Roles** panel, including roles whose eligibility lives at a management-group scope and inherits down to subscriptions.
- **Deactivate** button on each active role for self-service early deactivation.
- Per-role expiry timers — active roles flip to "Expired" exactly at their end time without polling Azure.

### Fixed
- Active assignments no longer leak permanent/standing direct role assignments into the Active Roles list — only PIM activations are shown.
- Eligibility-vs-activation matching uses scope-aware lookup so MG-scope activations are paired correctly with their parent eligibility row.

## [2026.4.1] - 2026-04-19

### Added
- CI release pipeline on GitHub Actions — pushing a `v*` tag builds the Windows installer and publishes to GitHub Releases automatically.
- Auto-update check every 30 minutes (was 5 minutes) — less network noise.
- Release procedure documented at [`misc/docs/releasing.md`](misc/docs/releasing.md).
