# IBI Task Manager — v6.14

Cloud-synced PWA for India Business International: the CEO sets self-targets, assigns work to staff, and verifies completion. Live at https://task.indiabusinessinternational.online (GitHub Pages, backend on Google Apps Script + Google Sheet).

## v6.14 — My order syncs across devices
- The CEO's "My order" sequence now syncs through the Google Sheet (Settings row `myOrder`, backend v5.2's `loadDB` returns it; `setSetting` writes always worked).
- Old backend still works: the device-local order (localStorage cache) is used until the new `.gs` is pasted, then auto-migrated to the sheet on first sync.
- `replaceAll` (backup restore) preserves the sheet's `myOrder` when the backup predates v6.14.

## v6.13 — My order for My Targets
- New **Sort: My order (arrange yourself)** option in **My Targets** — the CEO's own priority sequence, independent of date or priority level.
- Reorder by **dragging a card** (desktop) or tapping the **▲ ▼** buttons (works on phones); each open card shows its position number.
- The order and the chosen sort are remembered on the device (localStorage, like Recurring templates — no backend change). New targets join at the bottom; completed tasks stay below the "Completed" line.
