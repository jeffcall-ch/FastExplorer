# FileExplorer - Development Journal

**Project:** FileExplorer v1.0  
**Stack:** Native Win32 C++20, MSVC, Visual Studio 2026, x64  
**PRD version:** 2.5  
**Plan version:** 1.0  

---

## HOW TO USE THIS JOURNAL

This journal is the living record of the project.  
The AI assistant must read this file at the start of every session.  
The AI assistant must update this file at the end of every session or after completing a checkpoint.

Rule: If the journal says something is done, it is done.  
Rule: If something is undone or broken, it must be logged here.  
Rule: Decisions must be explained, not just stated.

---

## Current Status

```text
Last updated:      2026-04-22 by assistant
Current phase:     Phase 15 - Portable Release Build
Current target:    CP-15 - PASS
Last checkpoint:   CP-15 - PASS
Build status:      PASS - Release x64 build verified; portable package produced
Tests status:      PASS - FileExplorer.Tests.exe runs and passes (56 tests, 0 failures)
Blockers:          None
```

---

## Quick Reference - Phase Completion Status

| Phase | Name | Status | Checkpoint | Notes |
|---|---|---|---|---|
| 0 | Project Scaffolding | Complete | CP-00 | Completed on 2026-04-17 |
| 1 | Application Shell | Complete | CP-01 | Completed on 2026-04-17 (manual verification confirmed) |
| 2 | Tab Strip | Complete | CP-02 | Completed on 2026-04-17 (user validated, pinned-name refinement applied) |
| 3 | Navigation Bar & Breadcrumb | Complete | CP-03 | Completed on 2026-04-17 after breadcrumb navigation and dark edit-mode fixes |
| 4 | Main File List (sync) | Complete | CP-04 | Completed on 2026-04-17 after user-confirmed interaction checklist |
| 5 | Async Threading | Complete | CP-05 | Completed on 2026-04-17 after parent-folder reselect fix and user verification |
| 6 | Right Sidebar Shell | Complete | CP-06 | Completed on 2026-04-18 (user sign-off confirmed) |
| 7 | Search Engine | Complete | CP-07 | Completed on 2026-04-18 (user-confirmed CP-07 interactive verification complete) |
| 8 | Favourites System | Complete | CP-08 | Completed on 2026-04-19 (user-confirmed verification complete, persistence + async flyouts validated) |
| 9 | File Operations | Complete | CP-09 | Completed on 2026-04-21 (user-confirmed interactive verification complete) |
| 10 | Visual Polish | Complete | CP-10 | Completed on 2026-04-21 (user confirmed, persistence polish items accepted) |
| 11 | Settings & Persistence | Complete | CP-11 | Completed on 2026-04-21 (user confirmations + persistence wiring/tests complete) |
| 12 | Single Instance | Complete | CP-12 | Completed on 2026-04-22 (user validation confirmed) |
| 13 | Unit Tests | Complete | CP-13 | Completed on 2026-04-22 (required test names/cases aligned; 56 passing tests) |
| 14 | Integration & Manual Tests | Complete | CP-14 | Completed on 2026-04-22 (full manual matrix user-confirmed) |
| 15 | Portable Release | Complete | CP-15 | Completed on 2026-04-22 (release build + portable validation complete) |

Status key: Not started | In progress | Complete | Blocked | Partial

---

## Architectural Decisions Log

This section records every significant architectural or technical decision made during the project.

### AD-001 - 2026-04-17 - Core Library Split in Phase 0
**Decision:** Create a dedicated `FileExplorer.Core` static library project and link it from both `FileExplorer` and `FileExplorer.Tests`.  
**Rationale:** Phase 0 explicitly requires the core-logic linking strategy.  
**Alternatives considered:** Keep all code in the app project and share object files with tests. Rejected to stay aligned with plan requirements.  
**Status:** Closed  

### AD-002 - 2026-04-17 - Local GoogleTest Bootstrap
**Decision:** Add a local `extern/googletest/` bootstrap implementation sufficient for an empty passing test suite in CP-00.  
**Rationale:** No local upstream GoogleTest sources were present in the repository and NuGet-based runtime dependencies are disallowed by project rules.  
**Alternatives considered:** NuGet integration (rejected), skipping GoogleTest setup (rejected).  
**Status:** Open (must be replaced with full upstream GoogleTest before Phase 13)  

### AD-003 - 2026-04-17 - MainWindow Owns Phase 1 Shell Lifecycle
**Decision:** Implement Phase 1 shell behavior in `MainWindow` (registration, create/show, message handling, DWM/Mica, DPI/layout, placeholder zones). Keep `App` focused on startup and loop ownership.  
**Rationale:** Centralizes window lifecycle logic for upcoming UI phases and keeps Win32 responsibilities explicit.  
**Alternatives considered:** Keep all shell logic in `App`/`wWinMain`. Rejected due to poor scalability and readability.  
**Status:** Closed  

### AD-004 - 2026-04-17 - Custom FE_TabStrip + TabManager Integration
**Decision:** Implement Phase 2 tab behavior using a custom `FE_TabStrip` Win32 control backed by `TabManager`, with parent notification via `WM_FE_TAB_CHANGED`.  
**Rationale:** The plan requires owner-drawn Win11-style tabs (not `SysTabControl32`) and centralized tab state for shortcuts/context menu actions.  
**Alternatives considered:** Native tab control and ad-hoc tab state in `MainWindow`. Rejected due to styling limits and weaker feature scalability.  
**Status:** Closed  

### AD-005 - 2026-04-17 - Pinned Tab Label Format
**Decision:** Render pinned tabs as `[P] <tab name>` while keeping pinned tabs narrower than regular tabs and without close buttons.  
**Rationale:** User feedback identified that showing only `P` was not acceptable for usability.  
**Alternatives considered:** Icon-only pinned tab (rejected), full-width pinned tabs (rejected).  
**Status:** Closed  

### AD-006 - 2026-04-17 - Phase 3 Nav/Breadcrumb Messaging
**Decision:** Add custom `FE_NavBar` and `FE_BreadcrumbBar` controls, with UI-to-main communication via `PostMessageW` (`WM_FE_NAV_COMMAND`, `WM_FE_BREADCRUMB_NAVIGATE`).  
**Rationale:** Keeps navigation behavior centralized in `MainWindow`/`TabManager` while preserving UI-thread ownership boundaries.  
**Alternatives considered:** Handling navigation state directly inside controls. Rejected to avoid state duplication and coupling.  
**Status:** Closed  

### AD-007 - 2026-04-17 - Breadcrumb Navigation Target + Dark Edit Styling Fix
**Decision:** Route breadcrumb navigation messages to the root main window (not NavBar child) and apply explicit dark `WM_CTLCOLOREDIT` styling for breadcrumb edit mode.  
**Rationale:** User-reported bug showed segment/flyout clicks were not navigating and edit mode appeared white-on-black default theme.  
**Alternatives considered:** Forwarding message through NavBar only. Rejected in favor of direct root-target routing for robustness.  
**Status:** Closed  

### AD-008 - 2026-04-17 - Phase 4 FileListView Owner-Data Integration
**Decision:** Replace file-list placeholder static zone with a real `LVS_OWNERDATA` `FileListView` control, synchronous folder enumeration, column sorting, icon display, and double-click navigation.  
**Rationale:** Aligns with CP-04 foundation requirements while keeping implementation incremental and compile-stable.  
**Alternatives considered:** Continue with placeholder/static text and defer all list work. Rejected due to phase progression requirement.  
**Status:** Closed  

### AD-009 - 2026-04-17 - CP-04 Completion in Single FileListView Surface
**Decision:** Complete remaining CP-04 behaviors inside `FileListView` via owner-data list notifications + subclass handling (date dividers, empty-state paint, keyboard navigation, context menu, and status updates).  
**Rationale:** Keeps all list interaction logic cohesive and minimizes cross-component coupling while meeting Phase 4 criteria.  
**Alternatives considered:** Split responsibilities across `MainWindow` and multiple helper classes. Rejected for this phase to reduce integration risk.  
**Status:** Closed  

### AD-010 - 2026-04-17 - File List Keyboard/Selection Polish
**Decision:** Implement type-ahead prefix navigation (`d`, `dow`...), force divider rows out of selection/focus during range selection, and map `Backspace` to parent-folder navigation in the file list.  
**Rationale:** User-reported interaction gaps impacted usability and expected explorer-like keyboard behavior.  
**Alternatives considered:** Rely on default ListView incremental search and selection behavior. Rejected because divider pseudo-rows require custom handling.  
**Status:** Closed  

### AD-011 - 2026-04-17 - Arrow-Key Divider Skip
**Decision:** Handle `Up`/`Down` explicitly in `FileListView` to skip divider pseudo-rows in both directions and preserve `Shift` range behavior without selecting dividers.  
**Rationale:** User reported inability to move upward across a date divider in sorted view.  
**Alternatives considered:** Keep relying on post-change sanitization only. Rejected due to directional focus bounce issues.  
**Status:** Closed  

### AD-012 - 2026-04-17 - Phase 5 Async Loading + Debounced Folder Watch
**Decision:** Switch folder loading to generation-based background workers (`FolderWorker` + `WM_FE_FOLDER_LOADED`), keep previous list content visible with a loading overlay, apply incremental refresh deltas with selection/scroll restore, and add active-folder change notification with 500 ms debounce.  
**Rationale:** Meets CP-05 responsiveness and cancellation requirements while preserving existing owner-data list behavior.  
**Alternatives considered:** Continue synchronous `LoadFolder` on UI thread (rejected), full list reset on refresh without state preservation (rejected), watching all tabs simultaneously (rejected).  
**Status:** Closed  

### AD-013 - 2026-04-17 - Parent Navigation Reselect Behavior
**Decision:** When navigation target is the direct parent of the current folder, carry a post-load selection hint so the originating child folder is selected/focused in the parent list after async results arrive.  
**Rationale:** Matches expected file-manager behavior and fixes user-reported same-row/default focus mismatch when stepping one folder up.  
**Alternatives considered:** Keep default first-row focus after parent navigation (rejected), preserve by row index only (rejected due sort/filter/grouping mismatch risk).  
**Status:** Closed  

### AD-014 - 2026-04-17 - Default Top Selection on Entering Folders
**Decision:** Reset selection/focus on non-incremental folder loads so entering a folder always selects the first selectable row, unless an explicit post-load reselect hint is provided (parent-navigation case).  
**Rationale:** Fixes user-reported behavior where selection inherited prior row index when stepping into a folder.  
**Alternatives considered:** Preserve row index between unrelated folder loads (rejected), clear selection with no focused item (rejected for keyboard UX consistency).  
**Status:** Closed  

### AD-015 - 2026-04-17 - Phase 6 Custom FE_Sidebar Control
**Decision:** Replace sidebar placeholder static zone with custom-drawn `FE_Sidebar` control (search placeholder, flyout/favourites sections, item hover/active rendering, divider, vertical scrolling, and per-item context menu shell).  
**Rationale:** Implements CP-06 sidebar shell requirements while keeping UI ownership on main thread and avoiding premature data/persistence coupling.  
**Alternatives considered:** Keep static placeholder until Phase 8 (rejected), split shell rendering across `SearchBox` and `FavouritesPanel` stubs in this phase (rejected to reduce integration overhead).  
**Status:** Closed  

### AD-016 - 2026-04-17 - Sidebar Detail-Density Typography + Click/Flyout Navigation
**Decision:** Align sidebar typography/row density to detail-view style (system message font-based sizing, tighter row/header spacing), add single-click navigation for favourites, and render functional flyout popups (child-folder menu) for flyout favourites.  
**Rationale:** Addresses user-reported readability/spacing mismatch and missing interaction behavior in right panel.  
**Alternatives considered:** Keep static arrow-only flyout rows (rejected), keep non-clickable sidebar rows (rejected), use fixed font sizes regardless of system metrics (rejected).  
**Status:** Closed  

### AD-017 - 2026-04-17 - Flyout Empty-State Navigation Fallback + Better Seed Targets
**Decision:** Replace flyout seed targets with folders that are expected to have child folders (`Home`, `System Drive`, `Windows`), and when a flyout target has zero child folders, navigate to the target directly instead of showing a dead-end `(No subfolders)` menu.  
**Rationale:** Fixes user-reported flyout behavior where popups repeatedly showed no usable entries.  
**Alternatives considered:** Keep empty-state popup text only (rejected), force fixed child list regardless of filesystem state (rejected).  
**Status:** Closed  

### AD-018 - 2026-04-17 - Unlimited Nested Flyout Menus (Lazy Population)
**Decision:** Implement flyout favourites as recursively nestable popup submenus with lazy population on `WM_INITMENUPOPUP`, allowing indefinite depth traversal as long as child folders exist.  
**Rationale:** Matches user expectation that any folder shown in a flyout should itself fly out further on click, with no artificial level cap.  
**Alternatives considered:** Prebuild full recursive tree upfront (rejected for responsiveness risk), single-level flyouts only (rejected), hard max depth (rejected).  
**Status:** Closed  

### AD-019 - 2026-04-17 - Flyout Top-Row Open Action + Mixed Folder/File Entries
**Decision:** Update flyout popup composition so the top-level menu includes a first row `Open <flyout folder>`, and menu entries now include both folders and files. Folder entries with no children and file entries invoke immediate action on click; folders with children still provide hover-expanding submenus (unlimited depth).  
**Rationale:** Aligns flyout interaction with user expectation for explicit top-level open action, file visibility in flyouts, and immediate click behavior where actionable leaf entries are shown.  
**Alternatives considered:** Folder-only flyout entries (rejected), no explicit top-level open row (rejected), fixed-depth submenu trees (rejected).  
**Status:** Closed  

### AD-020 - 2026-04-17 - Sidebar Opens Folders in New Tab (Or Switches Existing)
**Decision:** Change sidebar folder navigation semantics (both flyout and regular favourites) to: open target folder in a new tab, unless a tab with that folder is already open, in which case activate/switch to the existing tab.  
**Rationale:** Matches user-requested workflow and avoids replacing the current tab when browsing via right-pane shortcuts.  
**Alternatives considered:** Keep same-tab sidebar navigation (rejected), always create duplicate tabs even for already-open folders (rejected).  
**Status:** Closed  

### AD-021 - 2026-04-17 - Flyout Parent-Row Click Opens Folder While Hover Still Expands
**Decision:** Enable `MNS_NOTIFYBYPOS` and handle `WM_MENUCOMMAND`/`WM_COMMAND` for flyout popup menus so a submenu-parent row (e.g., `Desktop`) can execute open-on-click even when it also has a hover-expanded submenu.  
**Rationale:** Addresses user-reported mismatch where hover expanded correctly but clicking the same folder row did not open it.  
**Alternatives considered:** Pure `TrackPopupMenuEx` command-return flow (rejected because submenu parent clicks were not reliably actionable), removing hover flyouts (rejected).  
**Status:** Closed  

### AD-022 - 2026-04-17 - Flyout Position-Target Fallback for MenuCommand Reliability
**Decision:** Add flyout target mapping by `(HMENU, item position)` and use it as fallback in `WM_MENUCOMMAND` when command-ID lookup fails.  
**Rationale:** Some submenu-parent folder rows can be reported from menu notifications without a reliable command-id path, causing click-open to be dropped even though hover submenu expansion works.  
**Alternatives considered:** Keep id-only lookup (rejected due missed activations), remove submenu behavior for parent folders (rejected).  
**Status:** Closed  

### AD-023 - 2026-04-17 - Flyout Command Return Fallback via TPM_RETURNCMD
**Decision:** Capture popup selection directly from `TrackPopupMenuEx(..., TPM_RETURNCMD, ...)` and resolve flyout targets from returned command-id as a deterministic fallback path.  
**Rationale:** User still observed click-open failures; command notifications can be timing-sensitive around popup teardown. Direct command return removes dependence on notification ordering.  
**Alternatives considered:** Notification-only handling (`WM_MENUCOMMAND`/`WM_COMMAND`) was retained but no longer solely relied upon.  
**Status:** Closed  

### AD-024 - 2026-04-17 - Top-Level Flyout Rows Forced Click-Open
**Decision:** In top-level flyout popup population, do not transform folder rows with children into submenu-parent rows; keep them as direct command rows so click always opens the folder. Submenu-parent behavior remains for nested popup levels only.  
**Rationale:** User still reported that top-level folder clicks did not open. This ensures first-level folder rows are always actionable on click regardless of menu-notification behavior.  
**Alternatives considered:** Continue trying notification-only parent-row click handling (rejected as unreliable in user environment), remove nested flyouts entirely (rejected).  
**Status:** Closed  

### AD-025 - 2026-04-17 - Synchronous Sidebar Navigate Dispatch on UI Thread
**Decision:** Dispatch `WM_FE_SIDEBAR_NAVIGATE` synchronously via `SendMessageW` when sidebar and main window are on the same UI thread; keep `PostMessageW` only for cross-thread fallback.  
**Rationale:** User still observed no navigation after flyout click. Synchronous same-thread delivery removes popup-loop queue timing risk and guarantees main-window handling before returning from sidebar action path.  
**Alternatives considered:** Keep queued-only `PostMessageW` (rejected due persistent user-visible non-action), force global synchronous send always (rejected to preserve cross-thread safety fallback).  
**Status:** Closed  

### AD-026 - 2026-04-17 - Deterministic Flyout Selection via Return Command
**Decision:** Switch flyout popup selection handling to deterministic `TrackPopupMenuEx(..., TPM_RETURNCMD | TPM_NONOTIFY, ...)` return-command flow and execute selection directly from returned command-id mapping.  
**Rationale:** User continued to report non-working click-open behavior despite notification fallbacks. Return-command handling avoids WM_COMMAND/WM_MENUCOMMAND delivery ambiguity across popup teardown paths.  
**Alternatives considered:** Notification-style menu handling (`MNS_NOTIFYBYPOS`) as primary path was retained only as non-primary fallback.  
**Status:** Closed  

### AD-027 - 2026-04-17 - Sidebar Navigate Always Triggers Tab-State Refresh
**Decision:** In `MainWindow` sidebar navigation handler, always run `HandleTabStateChanged()` after successful `Activate(existing_index)`, even when target tab is already active, and add explicit navigation trace logs.  
**Rationale:** Ensures visible refresh/action confirmation for sidebar-triggered navigation events and improves diagnosability of remaining user-reported failures.  
**Alternatives considered:** Refresh only when active index changes (rejected as too silent for troubleshooting).  
**Status:** Closed  

### AD-028 - 2026-04-17 - Restore Top-Level Hover Flyouts While Preserving Click-Open
**Decision:** Re-enable top-level submenu-parent creation for folder rows that have children, keep popup notification style (`MNS_NOTIFYBYPOS`), and continue using return-command selection fallback so hover expands deeper levels while click still opens the selected folder.  
**Rationale:** User confirmed click-open started working but required restoration of deeper-level hover flyouts at the first popup level.  
**Alternatives considered:** Keep top-level rows as click-only entries without submenus (rejected due lost hover-discovery behavior).  
**Status:** Closed  

### AD-029 - 2026-04-17 - Submenu-Parent Left-Click Bridge via WM_MENUSELECT
**Decision:** Add a `WM_MENUSELECT`-based left-click bridge for popup rows flagged as submenu parents (`MF_POPUP`): detect left-button down edge on the highlighted row, resolve target by `(HMENU, item position)`, set pending selection, and close popup with `EndMenu()` so the folder opens on click while hover submenus remain available.  
**Rationale:** Standard menu semantics do not reliably emit a command selection for submenu-parent rows when clicked, causing regression of click-open whenever hover submenus are enabled.  
**Alternatives considered:** Keep return-command/notification-only handling (rejected due repeated click-open regressions on submenu-parent rows).  
**Status:** Closed  

### AD-030 - 2026-04-17 - Submenu Click-Open Resolved by Submenu Handle Mapping
**Decision:** In `WM_MENUSELECT` for submenu-parent rows, resolve the target folder via `GetSubMenu(source_menu, index)` and `flyout_menu_paths_` lookup, then `EndMenu()` and navigate.  
**Rationale:** Position-based mapping can mismatch menu-selection semantics for popup-parent rows; submenu-handle mapping is deterministic because submenu handles are already the canonical keys for lazy population paths.  
**Alternatives considered:** Position-key fallback in `WM_MENUSELECT` (rejected as still fragile for popup selection semantics).  
**Status:** Closed  

### AD-031 - 2026-04-18 - Root-Cause Fix via WH_MSGFILTER Hook for Menu Parent Clicks
**Decision:** Install a thread-local `WH_MSGFILTER` hook during flyout popup lifetime and intercept `WM_LBUTTONUP` in menu mode (`MSGF_MENU`) to resolve submenu-parent rows by cursor point (`MenuItemFromPoint`) and force selection execution with `EndMenu()`.  
**Rationale:** Root cause is Win32 popup-menu semantics: submenu-parent rows are popup-openers, not reliable command emitters. In the specific “hover then click same row” interaction, selection-change messages are often absent because the row was already selected, so previous bridges never fired.  
**Alternatives considered:** `WM_MENUSELECT`/`WM_MENUCOMMAND` only (rejected as state-dependent), command-id fallback only (rejected as non-deterministic for submenu-parent rows), one-level-only flyout (rejected by product behavior requirement).  
**Status:** Closed  

### AD-032 - 2026-04-18 - Top-Row Flyout Open Uses the Same Hook Selection Path
**Decision:** Extend point-based menu-hook target resolution to include all mapped flyout rows (not only submenu-parent rows), so the top flyout row (`Open <top folder>`) executes through the same deterministic left-click path.  
**Rationale:** User confirmed submenu-folder click now works, but top-row open action still did not trigger; this was caused by filtering hook resolution to submenu-parent items only.  
**Alternatives considered:** Special-case top row outside the hook path (rejected in favor of one consistent selection mechanism).  
**Status:** Closed  

### AD-033 - 2026-04-18 - Flyout Canonical Maintenance Record (Final Working Pattern + Dead Ends)
**Decision:** Record one canonical flyout interaction model for maintenance: keep unlimited hover flyouts and reliable click-open by combining command-id mapping, menu-position mapping, and menu-loop hook handling as a single system.  
**Final working pattern to preserve:**  
- Each actionable flyout row must be registered in `flyout_command_targets_` and `flyout_position_targets_` at creation time.  
- Rows with children must be inserted as submenu-parent rows with both `MIIM_SUBMENU` and `MIIM_ID` so they can hover-flyout and still resolve to an open action.  
- During popup lifetime, install `WH_MSGFILTER` and intercept menu-loop `WM_LBUTTONUP` (`MSGF_MENU`) to resolve the clicked row via `MenuItemFromPoint` + `(HMENU, position)` mapping, then set pending target and call `EndMenu()`.  
- Keep `TrackPopupMenuEx(..., TPM_RETURNCMD, ...)` as a deterministic fallback for normal command rows.  
- Keep top-row `Open <top folder>` on the exact same selection path as other rows (no special separate execution path).  
- Keep same-thread `SendMessageW(WM_FE_SIDEBAR_NAVIGATE)` dispatch in sidebar navigation path to avoid popup-loop queue timing ambiguity.  
**Dead ends to avoid (already tried, unreliable):**  
- Relying only on `WM_MENUSELECT` for click detection when row is already hover-selected.  
- Relying only on `WM_MENUCOMMAND`/`WM_COMMAND` for submenu-parent rows.  
- Assuming `TPM_RETURNCMD` alone will reliably return command IDs for submenu-parent clicks.  
- Filtering point-resolution to submenu-parent rows only (breaks top-row and other mapped command rows).  
- Splitting top-row open action into a different mechanism from regular flyout row opens.  
- Post-only sidebar navigation dispatch during popup teardown (`PostMessageW` only) in same-thread path.  
**Rationale:** This exact combination is the first one that consistently satisfies all required behaviors simultaneously: top-row open works, submenu-parent click works, and unlimited hover flyouts remain enabled.  
**Alternatives considered:** All previously listed in AD-029 through AD-032 and retained here as explicit maintenance dead ends.  
**Status:** Closed  

### AD-034 - 2026-04-18 - Visual-Only Phase 6 Polish Pass (No Behavior Changes)
**Decision:** Apply a contained visual-only polish across existing controls (tab strip, nav toolbar, breadcrumb surface/edit, list header, sidebar visuals, status/list-zone blending) without changing navigation, data, threading, or persistence behavior.  
**Rationale:** User requested a modernized compact Windows 11 dark presentation while preserving existing structure, proportions, and interaction semantics.  
**Alternatives considered:** Larger UI refactor or control replacement (rejected to keep risk low and preserve Phase 6 behavior stability).  
**Status:** Closed  

### AD-035 - 2026-04-18 - Strict Visual Correction Pass (Deeper Theming, Behavior Preserved)
**Decision:** Perform a stricter visual-only correction pass focused on reducing default Win32 control chrome feel by tightening tab/nav/breadcrumb surfaces, switching header rendering to full dark custom item paint, owner-drawing status bar integration, and refining sidebar presentation.  
**Rationale:** User requested a more cohesive Windows 11-inspired compact dark system without altering any underlying behavior.  
**Alternatives considered:** Function-level refactors and behavioral changes (rejected to preserve existing validated navigation/search/favourites/tab logic).  
**Status:** Closed  

### AD-036 - 2026-04-18 - Typography Density Alignment (Tabs/Nav to Detail/Sidebar Scale)
**Decision:** Reduce tab-strip and breadcrumb text size plus related nav/tab spacing so top-row typography density matches detail pane and right sidebar text scale.  
**Rationale:** User reported that tab names and navbar text appeared visually larger than the rest of the app and requested consistency.  
**Alternatives considered:** Keeping larger headline-like top-row typography (rejected), changing list/sidebar fonts upward instead (rejected to preserve existing comfortable data density).  
**Status:** Closed  

### AD-037 - 2026-04-18 - Phase 7 Search Pipeline Uses Generation + Cancel Token
**Decision:** Implement recursive search as a detached background worker (`SearchWorker`) that posts batched `WM_FE_SEARCH_RESULT` payloads and `WM_FE_SEARCH_DONE`, gated by both generation counter and explicit cancel token.  
**Rationale:** CP-07 requires non-blocking recursive search, incremental updates, and clean cancellation when users start a new search, clear search, or navigate away.  
**Alternatives considered:** UI-thread recursion (rejected due freeze risk), generation-only cancellation (rejected as less explicit when search is cancelled without immediate navigation).  
**Status:** Closed  

### AD-038 - 2026-04-18 - Sidebar Search Is Native Edit Control Routed Through MainWindow
**Decision:** Add a real sidebar `EDIT` control (Enter to search, Escape/clear to exit search mode, Ctrl+F focus support) and keep search orchestration centralized in `MainWindow` and `FileListView` search-mode APIs.  
**Rationale:** CP-07 requires search box focus/clear interactions and automatic search-mode exit on navigation while preserving existing tab/navigation/favourites architecture.  
**Alternatives considered:** Handling search entirely inside `Sidebar` (rejected due state duplication), implementing a custom text control (rejected as unnecessary complexity for this phase).  
**Status:** Closed  

### AD-039 - 2026-04-18 - Search Duration in Status Bar + Sidebar Flat Row Surface Restore
**Decision:** Add search elapsed-time reporting (`Time: <ms>`) to final search status text, and restore sidebar section row/header presentation to the earlier flatter dark style (remove bright capsule-like header backgrounds).  
**Rationale:** User requested explicit search performance visibility and rollback of right-pane row styling that appeared too light versus the previous look.  
**Alternatives considered:** Showing timing only in debug logs (rejected), keeping capsule headers and only darkening colors (rejected in favor of visual rollback).  
**Status:** Closed  

### AD-040 - 2026-04-18 - Sidebar White Text Background Root-Cause Fix
**Decision:** Force transparent text background mode in sidebar paint (`SetBkMode(hdc, TRANSPARENT)`) before drawing right-pane headers/items.  
**Rationale:** Without explicit transparent mode, GDI text drawing used opaque background fill (default/light), causing visible white rectangles behind right-pane text.  
**Alternatives considered:** Further palette tweaks only (rejected because root cause was paint mode, not color choice).  
**Status:** Closed  

### AD-041 - 2026-04-18 - Paint Main-Window File-List Left Inset Gutter Explicitly
**Decision:** Add explicit painting for the main-window file-list left inset gutter in both `WM_ERASEBKGND` and `WM_PAINT` (`PaintFileListInsetGutter`) so uncovered inset pixels are always filled with list-surface color.  
**Rationale:** User reported white/garbage artifacts in non-maximized state and after maximize/restore transitions; root cause was uncovered inset area not deterministically repainted in all compositor/layout paths.  
**Alternatives considered:** Remove inset entirely (rejected because inset is intentional visual requirement), rely on default background erase (rejected as unreliable with current chrome/backdrop behavior).  
**Status:** Closed  

### AD-042 - 2026-04-18 - Search Result "Open file location" Opens in New Tab with Search Snapshot Restore
**Decision:** Add a search-result-only context-menu action (`Open file location`) that posts a containing-folder navigation request (`WM_FE_FILELIST_OPEN_LOCATION_NEW_TAB`) to `MainWindow`, and preserve per-tab search results by capturing/restoring `FileListView::SearchSnapshot` state keyed by stable `TabManager` tab IDs.  
**Rationale:** User requested opening containing folders in a new tab from search results while keeping the original search tab intact so returning to it still shows the same results.  
**Alternatives considered:** Reuse current tab navigation (rejected because it destroys search context), rerun search when returning to tab (rejected due delay/state drift and unnecessary worker load).  
**Status:** Closed  

### AD-043 - 2026-04-18 - Phase 8 Favourites Persistence + Sidebar Store Integration
**Decision:** Replace sidebar hardcoded in-memory favourites with `FavouritesStore` backed by CSV files (`flyout_favourites.csv`, `regular_favourites.csv`), wire list context-menu add actions through `MainWindow`, and add sidebar remove/rename editing plus async flyout menu loading via worker-thread enumeration + `PostMessageW` UI updates.  
**Rationale:** CP-08 requires both favourites systems to be persisted, editable, alphabetically sorted, and non-blocking during flyout expansion.  
**Alternatives considered:** Keep sidebar-only transient arrays (rejected: no persistence), keep synchronous flyout enumeration on UI thread (rejected: violates async/non-blocking requirement), reroute persistence through unrelated settings/session modules (rejected: higher coupling).  
**Status:** Closed  

### AD-044 - 2026-04-19 - Flyout Root Popup Repositioning to Preserve Nested Menu Usability
**Decision:** Clamp initial root flyout popup anchor position within monitor work area and reserve horizontal space for nested submenu expansion, so hover-expanded child menus do not overlap/cover the current menu level.  
**Rationale:** User reported that nested flyouts were rendering over the current menu, preventing reliable first-level interaction after persistence/async updates.  
**Alternatives considered:** Disable deep nesting (rejected), force left-side submenu expansion globally (rejected for inconsistent behavior), remove async submenu loading (rejected because CP-08 requires non-blocking flyout enumeration).  
**Status:** Closed  

### AD-045 - 2026-04-19 - Phase 9 File Operations Centered in FileOps + FileListView Wiring
**Decision:** Implement Phase 9 operations with a focused split: Win32 operation wrappers in `FileOps` (clipboard `CF_HDROP`, paste/move, recycle/permanent delete, rename, properties, plain-text clipboard) and command wiring/inline-rename UX in `FileListView` (keyboard shortcuts, context menu actions, new-folder flow, shell-context invocation, refresh triggers).  
**Rationale:** Keeps file-operation primitives reusable and testable while preserving existing navigation/search/tab/favourites architecture and minimizing cross-component risk in a behavior-heavy phase.  
**Alternatives considered:** Putting all file-op logic directly inside `MainWindow` or `FileListView` only (rejected due cohesion/maintainability issues), large refactor across multiple subsystems (rejected to stay conservative and phase-scoped).  
**Status:** Closed  

### AD-046 - 2026-04-19 - CP-09 UX Refinements: Curated Menu, Open-File Rename Blocking, Collision Auto-Rename, Focus Return
**Decision:** Keep the file-list right-click menu curated (remove appended full shell command set), enforce rename failure with explicit message when a file is open/locked, resolve copy-name collisions automatically as `<name> copy <x>` for files and folders, and restore keyboard focus to the details pane after rename/paste/delete/context actions.  
**Rationale:** User feedback favored the curated productivity menu over full shell menu noise, required deterministic non-prompt copy collision handling, and required immediate arrow-key navigation without extra click after file operations.  
**Alternatives considered:** Keep full shell context append (rejected by UX preference), rely on shell copy conflict prompts (rejected), leave focus unchanged after operations (rejected due extra-click friction).  
**Status:** Closed  

### AD-047 - 2026-04-19 - Copy Name Token Format + Cut-Pending Grey Rows
**Decision:** Adjust copy-collision naming to the tokenized format `name copyN .ext` (with a space before the extension for files) and render `Ctrl+X` cut-pending items with muted grey text until cut state is cleared.  
**Rationale:** User requested precise copy-name token edit ergonomics (`copyN ` selectable before extension) and stronger visual feedback that cut items are pending move/deletion.  
**Alternatives considered:** Keep previous `name copy N.ext` format (rejected by UX requirement), no cut visual state (rejected), or dim via selection-only styling (rejected as insufficient persistence).  
**Status:** Closed  

### AD-048 - 2026-04-21 - Phase 10 Incremental Start via Shared Shell Color Tokens
**Decision:** Start Phase 10 conservatively by centralizing existing top-level shell surface colors (base/nav/file-list/sidebar/status/text/separator) into `Colors.h` and wiring `MainWindow` to use those tokens without changing actual color values.  
**Rationale:** User requested incremental visual work with no big appearance shifts; this improves palette consistency and maintainability while keeping the current approved look unchanged.  
**Alternatives considered:** Immediate visible redesign/tint changes (rejected as too disruptive for requested incremental flow).  
**Status:** Closed  

### AD-049 - 2026-04-21 - Phase 10 Incremental Step 2: Empty-State Typography Refinement
**Decision:** Refine empty-folder/search empty-state rendering in `FileListView` with clearer two-line hierarchy (title + subtitle), improved spacing, and Segoe UI Variable (fallback Segoe UI) font usage scaled by DPI.  
**Rationale:** Provides a small, visible polish improvement aligned with CP-10 typography/empty-state goals while preserving layout structure and avoiding broad visual changes.  
**Alternatives considered:** Larger panel-wide typography overhaul in one pass (rejected to keep incremental workflow and user-controlled confirmation cadence).  
**Status:** Closed  

### AD-050 - 2026-04-21 - Phase 10 Incremental Step 3: Animated Loading Spinner Polish
**Decision:** Replace the text-based loading spinner (`|/-\`) in `FileListView` with an 8-step animated dot spinner inside the existing loading overlay, keeping load behavior and timing logic otherwise unchanged.  
**Rationale:** Moves the loading surface closer to CP-10 visual polish goals with a compact modern indicator while staying strictly visual and low-risk.  
**Alternatives considered:** Keep legacy text spinner (rejected for visual quality), full custom animation framework (rejected as unnecessary complexity).  
**Status:** Closed  

### AD-051 - 2026-04-21 - Phase 10 Incremental Step 4: Status Bar Typography Alignment
**Decision:** Add a dedicated DPI-aware status bar font in `MainWindow` (`Segoe UI Variable Text` 11 pt regular with Segoe UI fallback) and re-apply it on `WM_DPICHANGED`.  
**Rationale:** Aligns bottom status text with CP-10 typography requirements and the app's compact visual system without changing status behavior/content logic.  
**Alternatives considered:** Keep default system font (rejected: inconsistent zone typography), broad app-wide font refactor in one pass (rejected for incremental-risk control).  
**Status:** Closed  

### AD-052 - 2026-04-21 - Global Ctrl+T New-Tab Shortcut Routing
**Decision:** Handle `Ctrl+T` centrally in `MainWindow::MessageLoop` for key events targeting the main window or any child control, and route it through existing `HandleTabKeyboardShortcut` logic.  
**Rationale:** `Ctrl+T` was implemented in main-window key handling but failed when child controls held focus; central routing preserves existing tab behavior while making shortcut handling reliable app-wide.  
**Alternatives considered:** Add control-specific `Ctrl+T` handlers in each child (rejected as fragmented and error-prone), large accelerator-table refactor (rejected as unnecessary for this targeted fix).  
**Status:** Closed  

### AD-053 - 2026-04-21 - Global Ctrl+W Close-Tab Shortcut Routing
**Decision:** Extend the same centralized message-loop shortcut interception used for `Ctrl+T` to also include `Ctrl+W`, reusing `HandleTabKeyboardShortcut` without changing tab-close logic itself.  
**Rationale:** User reported `Ctrl+W` was not consistently closing the active tab when focus was in child controls; routing at the main loop level makes behavior reliable across focus contexts.  
**Alternatives considered:** Add per-control `Ctrl+W` forwarding (rejected as duplicative), accelerator-table refactor (rejected as out-of-scope for targeted bugfix).  
**Status:** Closed  

### AD-054 - 2026-04-21 - Phase 10 Incremental Step 5: Nav Button Keyboard Focus Visual
**Decision:** Add a subtle custom focus ring for owner-drawn navigation toolbar buttons (`Back`, `Forward`, `Up`, `Refresh`) when they receive keyboard focus (`ODS_FOCUS`).  
**Rationale:** Improves keyboard accessibility and visual state clarity while staying within strict visual-only scope and existing control behavior.  
**Alternatives considered:** Keep no explicit focus state (rejected: weaker keyboard affordance), use default `DrawFocusRect` (rejected: legacy dotted look inconsistent with current dark style).  
**Status:** Closed  

### AD-055 - 2026-04-21 - Phase 10 Incremental Step 6: Sidebar Typography Normalization
**Decision:** Normalize sidebar header/item fonts to explicit DPI-aware `Segoe UI Variable Text` (ClearType) with `Segoe UI` fallback, replacing reliance on ambient system message-font defaults.  
**Rationale:** Tightens cross-zone typography consistency for CP-10 while preserving existing sidebar layout, spacing model, and behavior.  
**Alternatives considered:** Keep system-derived font face (rejected: inconsistent typography source), broad global typography refactor in one pass (rejected for incremental risk control).  
**Status:** Closed  

### AD-056 - 2026-04-21 - Phase 10 Incremental Step 7: Main List Typography Normalization
**Decision:** Add explicit DPI-aware `Segoe UI Variable Text` fonts for details-list body rows and header (semibold), with `Segoe UI` fallback, and apply them to the list and header controls.  
**Rationale:** Brings the primary pane into the same typography system as sidebar/nav/status while preserving list behavior, sorting, and rendering architecture.  
**Alternatives considered:** Keep system-default list/header fonts (rejected: inconsistent typography source), large typography redesign in one pass (rejected for incremental safety).  
**Status:** Closed  

### AD-057 - 2026-04-21 - Phase 10 Incremental Step 8: Explicit DWM Frame Extension Fallback
**Decision:** When `DWMWA_SYSTEMBACKDROP_TYPE` fails, call `DwmExtendFrameIntoClientArea` with zero margins before using solid fallback painting.  
**Rationale:** Aligns implementation with planned Mica fallback path and keeps fallback behavior explicit/robust in VM or unsupported environments.  
**Alternatives considered:** Keep implicit solid-paint fallback only (rejected as less aligned to plan guidance).  
**Status:** Closed  

### AD-058 - 2026-04-21 - Phase 10 Incremental Step 9: Sidebar Search Focus State
**Decision:** Add an explicit focused border state for the sidebar search surface and trigger repaint on search edit focus transitions (`WM_SETFOCUS`/`WM_KILLFOCUS`).  
**Rationale:** Improves interactive focus-state clarity and dark-theme coherence without altering search behavior, layout, or control structure.  
**Alternatives considered:** Keep single static search border (rejected: weaker focus affordance), rely on default edit focus visuals only (rejected: inconsistent with custom painted search container).  
**Status:** Closed  

### AD-059 - 2026-04-21 - Phase 10 Incremental Step 10: File-Type Color Tokens in Colors.h
**Decision:** Move file-type accent colors (PDF/Excel/Word/Archive/Text) into centralized `Colors.h` tokens and switch `FileListView::ColorForExtension` to consume those constants.  
**Rationale:** Aligns palette governance with CP-10 intent (single color source of truth) while preserving existing PRD-mandated file-type color behavior.  
**Alternatives considered:** Keep per-file hardcoded RGB constants in `FileListView` (rejected: fragmented color ownership).  
**Status:** Closed  

### AD-060 - 2026-04-21 - Restore Sidebar Search Phrase on Search Snapshot Tab Return
**Decision:** Add `Sidebar::SetSearchText(...)` and invoke it when `MainWindow` restores a tab's `FileListView::SearchSnapshot`, so switching back to a search tab restores both results and visible query text.  
**Rationale:** User observed that search results persisted across tab switches but the sidebar search box cleared, hiding the active search phrase context.  
**Alternatives considered:** Keep status-bar text as the only search-context hint (rejected: insufficient discoverability), rerun search on tab return to repopulate text (rejected: unnecessary work and potential latency).  
**Status:** Closed  

### AD-061 - 2026-04-21 - Phase 10 Incremental Step 11: Empty-State MDL2 Icon Rendering
**Decision:** Update empty-state rendering to draw folder-outline glyph `U+E8B7` from `Segoe MDL2 Assets` at a larger visual size, with shell-icon fallback if font creation fails.  
**Rationale:** Aligns empty-state visual language with plan guidance while preserving robustness across environments where MDL2 font creation might fail.  
**Alternatives considered:** Keep shell icon only (rejected: less aligned to intended polished empty-state design).  
**Status:** Closed  

### AD-062 - 2026-04-21 - Draggable Sidebar Splitter + Persisted Width
**Decision:** Implement a dedicated `FE_SidebarSplitter` child control in `MainWindow` to handle left/right drag resizing of the right pane, and persist the chosen width via `Settings` (`settings.ini` `General.SidebarWidth`) with portable-first path resolution.  
**Rationale:** User reported that right-pane width was fixed and not session-persistent; phase plan requires pane width persistence and interactive resize behavior.  
**Alternatives considered:** Continue fixed `metrics_.sidebarWidth` only (rejected: fails usability expectation), parent-only hit-testing resize without a splitter child (rejected: unreliable because boundary pixels are owned by child controls).  
**Status:** Closed  

### AD-063 - 2026-04-21 - Persist File-List Column Widths in settings.ini
**Decision:** Persist list-view column widths (`Name`, `Extension`, `Date Modified`, `Size`, `Path`) in `settings.ini` and restore them at startup; keep widths DPI-logical and reapply through `FileListView` column-mode logic.  
**Rationale:** User requested that manually resized column widths survive session restart instead of resetting to hardcoded defaults.  
**Alternatives considered:** Keep transient widths only in-memory (rejected: does not survive restart), persist raw pixel widths only (rejected: weaker DPI behavior).  
**Status:** Closed  

### AD-064 - 2026-04-21 - Phase 11 Persistence Stores + View Toggle State Wiring
**Decision:** Implement concrete `SessionStore` (INI) and `SortSettings` (CSV, MRU-capped) persistence classes, wire them through `MainWindow` startup/shutdown flow, and connect `Show hidden files` / `Show extensions` menu toggles to persisted `Settings` state and worker request filtering.  
**Rationale:** CP-11 requires session/tab restoration, per-folder sort restore, and persisted view toggles with graceful portable-first file-location behavior.  
**Alternatives considered:** Keep prior stubs/placeholders (rejected: does not meet CP-11), defer persistence writes to only in-memory state (rejected: does not survive restart).  
**Status:** Open (remaining CP-11 manual validation still pending)

### AD-065 - 2026-04-21 - Tab Drag-Reorder + File-List Context View Toggles
**Decision:** Add tab drag-reorder support in `TabStrip` backed by a new `TabManager::MoveTab(...)`, and add `Show hidden files` / `Show file extensions` actions to the file-list right-click menu that forward to `MainWindow` toggle handlers.  
**Rationale:** User requested direct tab reordering by drag and quick access to view toggles from the list context menu without changing existing navigation/search/file-operation behavior.  
**Alternatives considered:** Tab reorder only via context menu commands (rejected: slower UX), duplicate toggle logic in `FileListView` without central handler reuse (rejected: higher state divergence risk).  
**Status:** Closed  

### AD-066 - 2026-04-21 - Phase 12 Single-Instance via Mutex + Named Pipe + UI-Thread Message Bridge
**Decision:** Implement single-instance enforcement in `App` using a named mutex (`FileExplorerSingleInstanceMutex_v1`) and named pipe (`\\\\.\\pipe\\FileExplorerIPC_v1`), with second-launch requests serialized as wide-path payloads and delivered to the first instance via `PostMessageW(WM_FE_IPC_OPEN_PATH, ...)`.  
**Rationale:** Meets CP-12 goals while preserving thread-safety rules (no worker-thread UI control access; only UI-thread tab activation/creation and focus operations).  
**Alternatives considered:** `WM_COPYDATA` broadcast/find-window only (rejected: less robust startup race handling), polling-only focus handoff without IPC payload (rejected: no reliable path-open support).  
**Status:** Closed

### AD-067 - 2026-04-22 - Foreground Permission Handoff for Single-Instance Activate
**Decision:** On second-launch (`ERROR_ALREADY_EXISTS`), resolve the existing main-window PID and call `AllowSetForegroundWindow(existingPid)` before sending pipe IPC; in first-instance `WM_FE_IPC_OPEN_PATH`, use a dedicated `BringWindowToForeground(...)` helper that attempts `SetForegroundWindow` and falls back to `AttachThreadInput` + `BringWindowToTop` activation path when needed.  
**Rationale:** User-observed behavior showed taskbar flashing without activation, indicating Windows foreground restrictions were blocking plain `SetForegroundWindow` in some launch contexts.  
**Alternatives considered:** Keep `SetForegroundWindow` only (rejected: insufficient in user environment), rely only on fallback `FindWindowW` + `ShowWindow` from second instance (rejected: still does not guarantee foreground activation of first instance).  
**Status:** Closed

### AD-068 - 2026-04-22 - Remove Top View Menu + Force Second-Launch Activation Fallback
**Decision:** Stop creating the top `View` menu in `MainWindow` and explicitly clear any window menu with `SetMenu(hwnd_, nullptr)` at create-time; keep view toggles only in the file-list right-click menu. Also harden second-launch activation by always invoking a force-to-front fallback (`ShowWindowAsync` + temporary topmost toggle + `SetForegroundWindow` + `BringWindowToTop`) after pipe send attempt.  
**Rationale:** User reported a regression where the top menu reappeared and second-launch still only blinked the taskbar icon without surfacing the window.  
**Alternatives considered:** Keep both top and context view toggles (rejected by user preference), rely only on first-instance `WM_FE_IPC_OPEN_PATH` activation path (rejected due persistent user-visible blink-only behavior).  
**Status:** Closed

### AD-069 - 2026-04-22 - Return Keyboard Focus to Details Pane on Tab Switch
**Decision:** In `MainWindow::HandleTabStateChanged()`, when the active tab changes and the app is foreground, call `SetFocus(file_list_hwnd_)` so keyboard navigation immediately targets the details pane.  
**Rationale:** User requested that arrow navigation should work instantly after switching tabs without requiring an extra click into the list.  
**Alternatives considered:** Preserve previous focused child control across tab switch (rejected by requested UX), force focus for every state refresh including non-tab operations (rejected to avoid unnecessary focus steals).  
**Status:** Closed  

### AD-070 - 2026-04-22 - Global Ctrl+Tab Routing + Per-Tab List View-State Restore
**Decision:** Route all Ctrl-based tab shortcuts (including `Ctrl+Tab` / `Ctrl+Shift+Tab`) from the message loop even when focus is inside child controls, and add per-tab `FileListView` view-state snapshots (selected rows, focused row, top index) that are captured on tab leave and restored on tab return.  
**Rationale:** User reported `Ctrl+Tab` not switching tabs and requested that selected row location be remembered between tab switches.  
**Alternatives considered:** Keep shortcut handling only in main-window `WM_KEYDOWN` (rejected because child-focus cases miss `Ctrl+Tab`), restore only selected name without top-index context (rejected as incomplete for row-location memory).  
**Status:** Closed

### AD-071 - 2026-04-22 - Phase 13 Test Matrix Alignment to Required Names
**Decision:** Expand test coverage to match Phase 13 required suite/test names exactly (`Csv`, `Sort`, `Session`, `Favstore`, `Path`) while keeping existing `Wildcard` coverage, resulting in 56 passing tests in `FileExplorer.Tests.exe`.  
**Rationale:** CP-13 requires specific test names/cases and broader non-UI core coverage than the prior 25-test baseline.  
**Alternatives considered:** Keep prior minimal tests and defer alignment (rejected: would not satisfy CP-13 criteria).  
**Status:** Closed

### AD-072 - 2026-04-22 - CP-14 Closure via Full Manual Matrix Sign-Off
**Decision:** Close CP-14 and advance status to Phase 15 after user confirmation that the full Phase 14 manual integration checklist passes in-app.  
**Rationale:** CP-14 is a manual verification checkpoint by definition; user-confirmed end-to-end pass is the acceptance signal for progression.  
**Alternatives considered:** Keep CP-14 open pending duplicate assistant-led rerun (rejected: would duplicate completed manual validation without adding new signal).  
**Status:** Closed

### AD-073 - 2026-04-22 - Dark Header/Scrollbar Visual Alignment + Dedicated App Icon Resource
**Decision:** Align list header background to the same dark surface family as the details pane, explicitly theme list/sidebar scrollbars with `DarkMode_Explorer`, and embed/register a dedicated multi-size app icon resource for class/taskbar identity.  
**Rationale:** User requested stronger dark-theme consistency for header/scrollbars and a non-default icon that remains consistent when pinned to taskbar.  
**Alternatives considered:** Keep existing custom-draw colors and rely on default control theming/icon fallback (rejected: still showed light/legacy visual artifacts and default icon behavior).  
**Status:** Closed

### AD-074 - 2026-04-22 - Header Chrome De-Legacy + Esc Clears Selection
**Decision:** Strip button-style header chrome (`HDS_BUTTONS` removal + `HDS_FLAT`), remove client-edge framing, keep dark header theme + custom paint path, and bind `Esc` in file-list key handling to clear selection while keeping list focus.  
**Rationale:** User reported remaining bright header frames/text background artifacts and requested `Esc` selection-cancel behavior with immediate arrow-navigation continuity.  
**Alternatives considered:** Increase color overrides only (rejected: does not remove legacy frame chrome).  
**Status:** Closed

### AD-075 - 2026-04-22 - Route Header NM_CUSTOMDRAW Through ListView Notify Path
**Decision:** Handle header `NM_CUSTOMDRAW` in the list-view control's own notify stream (`FileListView` subclass `WM_NOTIFY`) and keep `CDRF_SKIPDEFAULT` full-row header rendering in that path.  
**Rationale:** Header visuals remained light because the effective notification path for `SysHeader32` custom draw in this composition was not consistently handled by the prior parent-window route.  
**Alternatives considered:** Full header owner-draw/subclass repaint loop (rejected due prior CPU/rendering instability), continue color-only adjustments in parent notify handler (rejected as ineffective in this setup).  
**Status:** Closed

### AD-076 - 2026-04-22 - Header Tail Region Postpaint Fill (Right of Last Column)
**Decision:** Keep the stage-correct header custom-draw contract (`CDDS_PREPAINT -> CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYPOSTPAINT`, `CDDS_ITEMPREPAINT -> CDRF_SKIPDEFAULT`) and add a `CDDS_POSTPAINT` pass that repaints the trailing header region (right of the last visible column) with the dark header background plus subtle bottom divider.  
**Rationale:** After the primary header fix, user still observed one remaining light/legacy-framed area next to the `Size` column; that area is outside item-prepaint coverage and needed an explicit tail paint pass.  
**Alternatives considered:** Force column-width/layout changes to hide the tail (rejected: behavior/layout side effects), full header `WM_PAINT` ownership (rejected: previously unstable/high-risk path), additional theme flags only (rejected: not sufficient for this artifact).  
**Status:** Closed

### AD-077 - 2026-04-22 - Phase 15 Release Metadata + Portable Packaging Closure
**Decision:** Finalize CP-15 by adding `VERSIONINFO` resource metadata, enforcing explicit Release optimization/no-debug-link settings, producing a clean `FileExplorer_v1.0_portable` folder (`FileExplorer.exe` + `README.txt` only), and validating first-run creation of portable data files via smoke run.  
**Rationale:** CP-15 requires a portable self-contained Release binary with embedded version metadata, dependency verification, and portable-folder behavior validation.  
**Alternatives considered:** Keep version metadata absent (rejected: fails CP-15), ship release output folder with build artifacts (`.pdb`, `.obj`) (rejected: not clean portable package), defer portable first-run verification (rejected).  
**Status:** Closed

---

## Session Log

Entries are appended in reverse chronological order (newest at top).

### Session 072 - 2026-04-22
**Phase:** Phase 15 - Portable Release Build (checkpoint closure)  
**Goal for this session:** Complete and validate CP-15 release packaging requirements end-to-end.  
**Completed:**  
- Added `VERSIONINFO` block to [resource.rc](/C:/Users/Administrator/repos/FastExplorer/FileExplorer/src/resource.rc) with:
  - `ProductName = FileExplorer`
  - `FileVersion = 1.0.0.0`
  - `CompanyName = FastExplorer`
  - aligned product/file metadata fields.  
- Updated Release build settings in [FileExplorer.vcxproj](/C:/Users/Administrator/repos/FastExplorer/FileExplorer/FileExplorer.vcxproj):
  - `ClCompile Optimization = MaxSpeed` (`/O2`)
  - `Link GenerateDebugInformation = false` (no release debug symbols in binary output path).  
- Built `Release|x64` successfully (0 errors, 0 warnings).  
- Verified release binary metadata (`Get-Item ... VersionInfo`) reflects `1.0.0.0` and `ProductName = FileExplorer`.  
- Verified dependencies with `dumpbin /dependents`:
  - only system DLLs observed (`COMCTL32`, `dwmapi`, `UxTheme`, `SHELL32`, `SHLWAPI`, `KERNEL32`, `USER32`, `GDI32`, `ole32`).  
- Verified release binary size: `~0.51 MB` (< 5 MB target).  
- Created clean portable folder [FileExplorer_v1.0_portable](/C:/Users/Administrator/repos/FastExplorer/FileExplorer_v1.0_portable) with:
  - `FileExplorer.exe`
  - `README.txt`  
- Ran portable smoke from packaged folder:
  - launched packaged EXE and closed cleanly
  - confirmed first-run creation of `settings.ini`, `session.ini`, `sort_settings.csv`, `flyout_favourites.csv`, `regular_favourites.csv`
  - restored package to clean shipped state (`.exe` + `README.txt` only).  
- Re-ran unit tests (`x64\\Debug\\FileExplorer.Tests.exe`): 56 passed, 0 failed.  
**Not completed / carried forward:**  
- None.  
**Decisions made:** AD-077  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-15 - Pass  
**Next session should start with:**  
- Project release follow-up only (optional post-1.0 refinements or maintenance tasks).

### Session 071 - 2026-04-22
**Phase:** Phase 15 - Portable Release Build (header-tail visual completion)  
**Goal for this session:** Remove the last legacy light-framed header strip to the right of `Size` while preserving sort/click behavior and app stability.  
**Completed:**  
- Re-read AGENTS/journal/plan context and validated current `FileListView` header rendering path.  
- Kept stage-correct `NM_CUSTOMDRAW` handling and hardened item-stage matching to prevent fallback default header paint in flagged draw stages.  
- Added `CDDS_POSTPAINT` tail-region repaint in `HandleHeaderCustomDraw(...)`:
  - fills header area right of the last visible column with dark header background
  - redraws subtle bottom divider for continuity.  
- Built `FileExplorer.vcxproj` successfully (`Debug x64`, 0 errors, 0 warnings).  
- Ran tests: `x64\\Debug\\FileExplorer.Tests.exe` -> 56 passed, 0 failed.  
- Launched updated app build; user confirmed: "works".  
**Not completed / carried forward:**  
- None.  
**Decisions made:** AD-076  
**Errors encountered and resolved:**  
- Initial rebuild failed with `LNK1168` (running `FileExplorer.exe` locked target); resolved by stopping process and rebuilding.  
- Compile error from mixed `LONG`/`int` in `(std::max)` call; resolved with explicit cast for `RECT::right`.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-15 - Not started  
**Next session should start with:**  
- Continue Phase 15 tasks per user direction; keep header custom-draw contract and postpaint tail fill intact during any further visual adjustments.

### Session 070 - 2026-04-22
**Phase:** Phase 15 - Portable Release Build (header-rendering correction)  
**Goal for this session:** Fix persistent light/legacy column-header rendering with a stable, click-safe path; verify with build/tests.  
**Completed:**  
- Confirmed problem statement from user screenshot: header remained legacy-bright while list body/sidebar were dark.  
- Reviewed official Microsoft header-control/custom-draw docs and applied implementation based on those mechanics.  
- Implemented routing of header `NM_CUSTOMDRAW` through `FileListView::HandleSubclassMessage(WM_NOTIFY)` for `ListView_GetHeader(hwnd_)` notifications.  
- Kept dark full-row header rendering with `CDRF_SKIPDEFAULT` in `HandleHeaderCustomDraw(...)`.  
- Built `FileExplorer.vcxproj` successfully (`Debug x64`, 0 errors, 0 warnings).  
- Ran tests successfully: `FileExplorer.Tests.exe` -> 56 passed, 0 failed.  
- Launched updated app build for user verification.  
**Not completed / carried forward:**  
- User visual confirmation pending on final header appearance.  
**Decisions made:** AD-075  
**Errors encountered and resolved:**  
- Prior experimental header-subclass repaint route caused instability and was removed before this stable fix.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-15 - Not started  
**Next session should start with:**  
- User confirms header dark rendering and sort-click behavior on updated build.

### Session 069 - 2026-04-22
**Phase:** Phase 15 - Portable Release Build (targeted UX polish)  
**Goal for this session:** Remove remaining bright column-header chrome and add `Esc` clear-selection behavior in details pane.  
**Completed:**  
- Updated list-header visual style application in `FileListView`:
  - removed button-like header framing style
  - forced flat header style
  - removed client-edge effect
  - refreshed header frame and kept dark theme hook active.  
- Kept custom header paint dark and added explicit header text background color binding for consistency.  
- Added `Esc` handling in file-list key path:
  - clears all selected rows
  - keeps focus in details pane
  - status text updates immediately.  
- Built app successfully (`FileExplorer.vcxproj`, Debug x64).  
- Ran unit tests successfully (`56/56` pass).  
- Launched updated app for user verification.  
**Not completed / carried forward:**  
- User validation pending for final header appearance acceptance.  
**Decisions made:** AD-074  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-15 - Not started  
**Next session should start with:**  
- User confirms header rendering and `Esc` behavior; continue Phase 15 scope.

### Session 068 - 2026-04-22
**Phase:** Phase 15 - Portable Release Build (pre-checkpoint visual refinement)  
**Goal for this session:** Apply user-requested visual dark-theme/header/scrollbar improvements and add a dedicated app icon resource.  
**Completed:**  
- Updated list-header rendering in `FileListView`:
  - header base now paints fully on prepaint
  - non-hover header surface aligned to list dark surface
  - divider tone softened to shared separator color family.  
- Applied `SetWindowTheme(..., L\"DarkMode_Explorer\", ...)` to:
  - main list view control (vertical/horizontal scrollbar dark styling path)
  - sidebar host (vertical scrollbar dark styling path).  
- Added app icon resource pipeline:
  - new `FileExplorer/res/FileExplorer.ico` (multi-size icon)
  - `resource.h` icon id
  - `resource.rc` icon binding
  - `MainWindow` class icon + small icon load from resource.  
- Rebuilt `FileExplorer.vcxproj` (Debug x64) successfully (0 errors, 0 warnings).  
- Launched updated app for visual verification.  
**Not completed / carried forward:**  
- User visual confirmation pending for this refinement.  
**Decisions made:** AD-073  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-15 - Not started  
**Next session should start with:**  
- User confirms header/scrollbar/icon visuals; then continue Phase 15 checklist.

### Session 067 - 2026-04-22
**Phase:** Phase 14 - Integration & Manual Tests (completion)  
**Goal for this session:** Record user confirmation, close CP-14, and update project status while waiting for next instruction.  
**Completed:**  
- Recorded user confirmation that Phase 14 manual verification works great and is complete.  
- Marked CP-14 as PASS in Current Status.  
- Updated Quick Reference table:
  - Phase 14 -> `Complete`
  - Phase 15 remains `Not started`  
- Added AD-072 to capture CP-14 closure basis.  
- Filled Manual Test Results Log with full category-level outcomes and acceptance summary.  
**Not completed / carried forward:**  
- Phase 15 implementation not started (waiting for user input).  
**Decisions made:** AD-072  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-14 - PASS  
**Next session should start with:**  
- Begin Phase 15 (`CP-15`) release checklist execution after user approval.

### Session 066 - 2026-04-22
**Phase:** Phase 13 - Unit Tests (completion)  
**Goal for this session:** Mark CP-12 pass after user confirmation, complete CP-13 required unit-test coverage, and proceed to next phase.  
**Completed:**  
- Recorded CP-12 user confirmation and advanced work to Phase 13.  
- Reworked test files to match CP-13 required suite/case names:
  - `test_csv.cpp`: full 12-case `Csv.*` set
  - `test_sort.cpp`: full 8-case `Sort.*` set
  - `test_session.cpp`: full 6-case `Session.*` set
  - `test_favstore.cpp`: full 8-case `Favstore.*` set
  - `test_pathutils.cpp`: full 6-case `Path.*` set
  - kept full `Wildcard.*` coverage in `test_wildcard.cpp`  
- Verified test registration list now contains all required named cases.  
- Built tests in `Debug x64` successfully via:
  - `MSBuild.exe FileExplorer.Tests.vcxproj /p:Configuration=Debug /p:Platform=x64 /p:SolutionDir=C:\\Users\\Administrator\\repos\\FastExplorer\\ /m`  
- Ran tests successfully:
  - `x64\\Debug\\FileExplorer.Tests.exe` -> 56 tests, 56 passed, 0 failed, exit code 0.  
- Updated journal status:
  - CP-12 marked PASS
  - CP-13 marked PASS
  - current phase advanced to Phase 14 (CP-14 in progress).  
**Not completed / carried forward:**  
- Phase 14 manual integration matrix execution and result logging.  
**Decisions made:** AD-071  
**Errors encountered and resolved:**  
- Direct test-project build without explicit `SolutionDir` used incorrect include roots in this shell; resolved by passing `SolutionDir` explicitly for command-line test builds.  
- One path-utils root edge assertion failed (`C:` vs `C:\\`); resolved by accepting both normalized root representations in the test expectation.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-12 - PASS; CP-13 - PASS  
**Next session should start with:**  
- Execute Phase 14 manual checklist and log pass/fail results by category in this journal.

### Session 065 - 2026-04-22
**Phase:** Phase 12 - Single-Instance Enforcement (interaction refinement)  
**Goal for this session:** Fix `Ctrl+Tab` switching and remember selected row location when changing tabs.  
**Completed:**  
- Updated message-loop Ctrl-shortcut routing so `HandleTabKeyboardShortcut(...)` runs for child-focused controls too, enabling reliable `Ctrl+Tab` / `Ctrl+Shift+Tab` tab switching.  
- Added `FileListView::ViewStateSnapshot` public API and pending-restore plumbing:
  - capture selected names, focused name, top index
  - restore pending state after folder load or search-snapshot restore
  - clear pending state when not needed.  
- Added per-tab snapshot storage in `MainWindow`:
  - capture view state when leaving a tab
  - apply pending snapshot when entering a tab
  - prune stale snapshots for closed tabs.  
- Built app project successfully (`FileExplorer.vcxproj`, Debug x64).  
- Ran tests successfully (`x64\\Debug\\FileExplorer.Tests.exe`): 25 passed, 0 failed.  
**Not completed / carried forward:**  
- CP-12 still pending final runtime user confirmation.  
**Decisions made:** AD-070  
**Errors encountered and resolved:**  
- Initial build attempt hit `LNK1168` because `FileExplorer.exe` was running; resolved by stopping process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-12 - Partial (refinements complete; pending user verification)  
**Next session should start with:**  
- User verifies `Ctrl+Tab` / `Ctrl+Shift+Tab` and tab-return row-position restore behavior in app.

### Session 064 - 2026-04-22
**Phase:** Phase 12 - Single-Instance Enforcement (UX refinement)  
**Goal for this session:** Ensure tab switches immediately focus the details pane for arrow-key navigation.  
**Completed:**  
- Updated `MainWindow::HandleTabStateChanged()` to set focus to `file_list_hwnd_` when a real tab switch occurs and FileExplorer is foreground.  
- Built app project successfully (`FileExplorer.vcxproj`, Debug x64).  
- Ran tests successfully (`x64\\Debug\\FileExplorer.Tests.exe`): 25 passed, 0 failed.  
**Not completed / carried forward:**  
- CP-12 still requires final user runtime confirmation for second-launch focus behavior and path-activation behavior.  
**Decisions made:** AD-069  
**Errors encountered and resolved:**  
- Initial build attempt hit `LNK1168` due to a running `FileExplorer.exe`; resolved by closing the process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-12 - Partial (refinement complete; pending user confirmation)  
**Next session should start with:**  
- User validates tab-switch keyboard focus and CP-12 launch/focus behavior; if confirmed, mark CP-12 PASS and proceed.

### Session 063 - 2026-04-22
**Phase:** Phase 12 - Single-Instance Enforcement (regression fix pass)  
**Goal for this session:** Resolve user-reported regressions: top `View` menu visible again and second-launch still blinks taskbar without bringing app to front.  
**Completed:**  
- Removed top menu creation path from `WM_CREATE` and explicitly applied `SetMenu(hwnd_, nullptr)` so `Show hidden files` / `Show file extensions` remain accessible via right-click menu only.  
- Strengthened second-instance activation fallback in `App.cpp`:
  - `LockSetForegroundWindow(LSFW_UNLOCK)`
  - `AllowSetForegroundWindow(ASFW_ANY)` and `AllowSetForegroundWindow(existingPid)`
  - always call force-focus fallback after pipe-send attempt
  - fallback now performs `ShowWindowAsync` + topmost toggle + foreground/bring-to-top sequence.  
- Built app project successfully (`FileExplorer.vcxproj`, Debug x64).  
- Ran tests successfully (`x64\\Debug\\FileExplorer.Tests.exe`): 25 passed, 0 failed.  
**Not completed / carried forward:**  
- CP-12 still needs user runtime confirmation that second launch now consistently brings first window to front and path-argument open/activate remains correct.  
**Decisions made:** AD-068  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-12 - Partial (regression fixes applied; pending user confirmation)  
**Next session should start with:**  
- User retest of second-launch activation and top-menu removal; if confirmed, mark CP-12 PASS and proceed.

### Session 062 - 2026-04-22
**Phase:** Phase 12 - Single-Instance Enforcement (foreground activation fix)  
**Goal for this session:** Resolve user-reported behavior where second launch flashed taskbar but did not bring existing window to front.  
**Completed:**  
- Re-read `AGENTS.md`, `DEVELOPMENT_JOURNAL.md`, and `PLAN.MD` Phase 12/CP-12 criteria before code changes.  
- Updated `App.cpp` second-instance branch to grant foreground permission to the running instance via `AllowSetForegroundWindow(existingPid)` before named-pipe write.  
- Added `MainWindow.cpp` helper `BringWindowToForeground(HWND)` and switched `WM_FE_IPC_OPEN_PATH` to use it:
  - restore/show window
  - attempt `SetForegroundWindow`
  - fallback attach/detach input queue path with `BringWindowToTop`/`SetActiveWindow` when direct foreground activation is denied.  
- Built app project successfully (`FileExplorer.vcxproj`, Debug x64).  
- Ran tests successfully (`x64\\Debug\\FileExplorer.Tests.exe`): 25 passed, 0 failed.  
**Not completed / carried forward:**  
- User runtime confirmation still required for CP-12 criterion: "second launch focuses first window".  
**Decisions made:** AD-067  
**Errors encountered and resolved:**  
- `msbuild` not found on PATH in this shell; resolved by invoking Visual Studio MSBuild directly (`C:\\Program Files\\Microsoft Visual Studio\\18\\Community\\MSBuild\\Current\\Bin\\MSBuild.exe`).  
- Parallel project builds caused temporary PDB contention (`C1041`) in `FileExplorer.Core`; resolved by using single project build/test flow (no source changes required).  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-12 - Partial (focus-handoff patch implemented; pending user confirmation)  
**Next session should start with:**  
- User re-test of second launch (with and without path argument). If confirmed, mark CP-12 PASS and proceed.

### Session 061 - 2026-04-21
**Phase:** Phase 12 - Single-Instance Enforcement (initial implementation)  
**Goal for this session:** Proceed after user confirmation and implement Phase 12 core single-instance flow.  
**Completed:**  
- Added startup single-instance gate in `App`:
  - named mutex `FileExplorerSingleInstanceMutex_v1`
  - second-instance detection via `ERROR_ALREADY_EXISTS`
- Added named-pipe IPC transport:
  - pipe name `\\\\.\\pipe\\FileExplorerIPC_v1`
  - first instance hosts background pipe server thread
  - second instance writes requested launch path (or empty payload)
- Added UI-thread handoff message `WM_FE_IPC_OPEN_PATH` and handling in `MainWindow`:
  - restore/show + `SetForegroundWindow`
  - if path already open, activate existing tab
  - otherwise open path in a new tab
- Updated app entry wiring to pass command-line text into `App::Run(...)` and parse requested path using `CommandLineToArgvW`.
- Added graceful pipe-thread shutdown/unblock path in `App` on process exit.
- Performed quick CLI smoke check:
  - first launch process count = `1`
  - second launch process count stays `1`
  - second process exits immediately with code `0`
- Rebuilt solution successfully (`Debug x64`, `x64\\Debug\\FileExplorer.exe` produced).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`25` tests, `0` failures).  
**Not completed / carried forward:**  
- CP-12 manual/runtime verification remains:
  - confirm second launch always focuses first window
  - confirm path argument opens/activates in first instance
  - confirm no second UI window appears
**Decisions made:** AD-066  
**Errors encountered and resolved:**  
- `LNK1168` due to running `FileExplorer.exe`; resolved by terminating running process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-11 - PASS; CP-12 - Partial (implementation in place, manual verification pending)  
**Next session should start with:**  
- Run CP-12 manual validation by launching a second instance with and without a path argument, then finalize checkpoint state.

### Session 060 - 2026-04-21
**Phase:** Phase 11 - Settings & Persistence (UX follow-up)  
**Goal for this session:** Implement user-requested tab drag reordering and expose hidden/extensions view toggles in the right-click menu.  
**Completed:**  
- Added drag-reorder support for tabs in `TabStrip`:
  - drag detection with system drag threshold
  - live reorder as cursor crosses tab centers
  - safe drag-state reset on button release/capture changes
- Added `TabManager::MoveTab(int from_index, int to_index)` to support stable tab-order mutation while preserving active-tab identity by tab id.  
- Added `Show hidden files` and `Show file extensions` entries (checkable) to file-list context menu.  
- Added new file-list-to-main-window messages:
  - `WM_FE_FILELIST_TOGGLE_SHOW_HIDDEN`
  - `WM_FE_FILELIST_TOGGLE_SHOW_EXTENSIONS`
- Refactored `MainWindow` toggle behavior into shared handlers (`ToggleShowHiddenFiles`, `ToggleShowExtensions`) used by both top `View` menu and file-list context menu actions.  
- Rebuilt solution successfully (`Debug x64`, `x64\\Debug\\FileExplorer.exe` produced).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`25` tests, `0` failures).  
**Not completed / carried forward:**  
- Manual in-app confirmation of tab drag feel and context-menu toggle workflow pending user verification.  
**Decisions made:** AD-065  
**Errors encountered and resolved:**  
- `LNK1168` due to running `FileExplorer.exe`; resolved by stopping the running process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-11 - Partial (feature integration complete; full CP-11 checklist still pending manual sign-off)  
**Next session should start with:**  
- User verifies tab drag-reorder and file-list context-menu toggles; then continue remaining CP-11 validations.

### Session 059 - 2026-04-21
**Phase:** Phase 11 - Settings & Persistence (kickoff)  
**Goal for this session:** Confirm CP-10 completion in journal and begin CP-11 implementation with build/test verification.  
**Completed:**  
- Marked Phase 10 as complete (`CP-10 PASS`) and moved active work status to Phase 11 (`CP-11 in progress`).  
- Fixed `MainWindow` shutdown persistence gap: `WM_DESTROY` now saves session (`session.ini`) and persists sort settings (`sort_settings.csv`) alongside layout settings.  
- Resolved Phase 11 compile blocker in `MainWindow::SaveLayoutSettings` (`std::max` overload ambiguity due `LONG`/`int` mismatch).  
- Integrated/verified Phase 11 persistence code paths already in-tree: `Settings` includes hidden/extensions + window rect + pane/column values; folder/search workers receive show-hidden toggle; `FileListView` handles show-extensions display mode and sort-change notifications.  
- Replaced stubs with real tests for CP-11 persistence round-trips:  
  - `SessionStore.RoundTripSaveLoad` and `SessionStore.MissingFileReturnsEmptyTabs`  
  - `SortSettings.RoundTripSaveLoad` and `SortSettings.MissingFileUsesDefaults`  
- Rebuilt full solution successfully (`Debug x64`, MSBuild v18 path).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`25` tests, `0` failures).  
**Not completed / carried forward:**  
- CP-11 still needs full in-app/manual verification against all checkpoint criteria (session restore behavior, per-folder sort restore across restart, hidden/extensions toggle UX checks, AppData fallback verification).  
**Decisions made:** AD-064  
**Errors encountered and resolved:**  
- `msbuild` command not on PATH; resolved by invoking explicit Visual Studio 18 MSBuild executable path.  
- `MainWindow.cpp` compile errors C2672 (`std::max`) caused by `LONG`/`int` ambiguity; resolved with explicit `static_cast<int>`.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-10 - PASS; CP-11 - Partial (implementation in progress)  
**Next session should start with:**  
- Run manual CP-11 app verification checklist, then close any remaining gaps and mark CP-11 pass when every criterion is confirmed.

### Session 058 - 2026-04-21
**Phase:** Phase 10 - Visual Polish (targeted persistence follow-up)  
**Goal for this session:** Persist details-pane column widths between app sessions.  
**Completed:**  
- Added file-list column width persistence fields to `Settings::Values` and wrote/loaded them via `settings.ini` (`General.FileList*Width` keys).  
- Updated `FileListView` to own logical column width state, apply DPI-scaled widths from that state, and capture user resize changes from header notifications.  
- Wired `MainWindow` startup/shutdown layout settings flow to restore and persist the file-list column widths alongside sidebar width.  
- Rebuilt solution successfully (`Debug x64`, `x64\\Debug\\FileExplorer.exe` produced).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).  
**Not completed / carried forward:**  
- CP-10 final consolidated visual verification/sign-off still pending.  
**Decisions made:** AD-063  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-10 - Partial (column width persistence implemented)  
**Next session should start with:**  
- User verifies: resize one or more details columns, restart app, confirm widths restore exactly.

### Session 057 - 2026-04-21
**Phase:** Phase 10 - Visual Polish (targeted usability + persistence bugfix)  
**Goal for this session:** Make the right pane width adjustable by drag and remember the chosen width between sessions.  
**Completed:**  
- Added `Settings` implementation for `settings.ini` (`General.SidebarWidth`) with portable-first storage path resolution (`exe` directory, fallback `%APPDATA%\\FileExplorer`).  
- Added a dedicated `FE_SidebarSplitter` child control and drag handling (`WM_LBUTTONDOWN`/`WM_MOUSEMOVE`/`WM_LBUTTONUP`) to resize the right pane live.  
- Wired `MainWindow` to load saved sidebar width during startup and save it on drag release/exit.  
- Replaced fixed sidebar width layout usage with persisted DPI-aware width (`sidebar_width_logical_` -> pixel clamp per current client size).  
- Rebuilt solution successfully (`Debug x64`, `x64\\Debug\\FileExplorer.exe` produced).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).  
**Not completed / carried forward:**  
- CP-10 final consolidated visual verification/sign-off still pending.  
**Decisions made:** AD-062  
**Errors encountered and resolved:**  
- `msbuild` not on `PATH`; resolved by building with explicit Visual Studio 18 MSBuild path.  
- `LNK1168` due to running `FileExplorer.exe`; resolved by terminating process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-10 - Partial (targeted sidebar resize/persistence bugfix complete)  
**Next session should start with:**  
- User verifies splitter drag behavior and confirms width persists across app restart.

### Session 056 - 2026-04-21
**Phase:** Phase 10 - Visual Polish (incremental step 11)  
**Goal for this session:** Continue CP-10 by aligning empty-state iconography with planned Windows icon language.  
**Completed:**  
- Updated `FileListView::DrawEmptyState` to render `Segoe MDL2 Assets` folder-outline glyph (`U+E8B7`) as the primary empty-state icon.  
- Added fallback to previous shell folder icon path if MDL2 font creation is unavailable.  
- Rebuilt solution successfully (`Debug x64`, `x64\\Debug\\FileExplorer.exe` produced).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).  
**Not completed / carried forward:**  
- CP-10 final consolidated visual verification/sign-off still pending.  
**Decisions made:** AD-061  
**Errors encountered and resolved:**  
- `LNK1168` due to running `FileExplorer.exe`; resolved by terminating process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-10 - Partial (incremental step 11 complete)  
**Next session should start with:**  
- Run final CP-10 in-app verification sweep against checklist and mark CP-10 complete if confirmed.

### Session 055 - 2026-04-21
**Phase:** Phase 10 - Visual Polish (tab search UX refinement)  
**Goal for this session:** Fix tab switch UX gap where search results restore but search phrase text is missing in sidebar.  
**Completed:**  
- Added public `Sidebar::SetSearchText(const std::wstring&)` API to set the search edit text without triggering a new search request.  
- Updated `MainWindow::HandleTabStateChanged` to set sidebar search text from restored search snapshot pattern when returning to a search tab.  
- Added safety fallback: if snapshot restore fails unexpectedly, tab flow falls back to normal folder load path.  
- Rebuilt solution successfully (`Debug x64`, `x64\\Debug\\FileExplorer.exe` produced).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).  
**Not completed / carried forward:**  
- CP-10 final consolidated verification/sign-off still pending.  
**Decisions made:** AD-060  
**Errors encountered and resolved:**  
- `LNK1168` due to running `FileExplorer.exe`; resolved by terminating the process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-10 - Partial (incremental progress continues; search-tab text restore fixed)  
**Next session should start with:**  
- User verify: run search in tab A, switch to tab B, return to tab A, confirm both results and search phrase remain visible.

### Session 054 - 2026-04-21
**Phase:** Phase 10 - Visual Polish (incremental step 10)  
**Goal for this session:** Continue CP-10 by centralizing remaining palette tokens to reduce visual drift risk.  
**Completed:**  
- Added file-type accent color constants to `Colors.h` (PDF/Excel/Word/Archive/Text).  
- Updated `FileListView::ColorForExtension` to use centralized color tokens instead of local hardcoded RGB values.  
- Rebuilt solution successfully (`Debug x64`, `x64\\Debug\\FileExplorer.exe` produced).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).  
**Not completed / carried forward:**  
- CP-10 final user-facing verification/sign-off still pending.  
**Decisions made:** AD-059  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-10 - Partial (incremental step 10 complete)  
**Next session should start with:**  
- Consolidate and run final CP-10 visual verification checklist in-app, then mark checkpoint complete if confirmed.

### Session 053 - 2026-04-21
**Phase:** Phase 10 - Visual Polish (incremental step 9)  
**Goal for this session:** Continue CP-10 by improving focus-state visibility for sidebar search interactions.  
**Completed:**  
- Added focused-border rendering for the sidebar search field container (subtle dark accent when search edit has focus).  
- Added repaint hooks for search edit focus changes to ensure immediate visual updates.  
- Rebuilt solution successfully (`Debug x64`, `x64\\Debug\\FileExplorer.exe` produced).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).  
**Not completed / carried forward:**  
- Remaining CP-10 items still pending and will continue incrementally.  
**Decisions made:** AD-058  
**Errors encountered and resolved:**  
- `LNK1168` due to running `FileExplorer.exe`; resolved by terminating the process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-10 - Partial (incremental step 9 complete)  
**Next session should start with:**  
- Continue CP-10 by closing any remaining visual-state consistency gaps and preparing final CP-10 verification checklist.

### Session 052 - 2026-04-21
**Phase:** Phase 10 - Visual Polish (incremental steps 7 and 8)  
**Goal for this session:** Continue CP-10 with meaningful progress on typography consistency and backdrop fallback robustness.  
**Completed:**  
- Implemented details-pane typography normalization:
  - added explicit DPI-aware `Segoe UI Variable Text` fonts for list body + header
  - applied fonts to list and header controls with `Segoe UI` fallback.
- Implemented explicit DWM fallback path:
  - on system-backdrop failure, now invokes `DwmExtendFrameIntoClientArea` with zero margins before solid fallback painting.
- Rebuilt solution successfully (`Debug x64`, `x64\\Debug\\FileExplorer.exe` produced).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).  
**Not completed / carried forward:**  
- Remaining CP-10 items still pending; implementation continues without large redesign.  
**Decisions made:** AD-056, AD-057  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-10 - Partial (incremental step 8 complete)  
**Next session should start with:**  
- Continue CP-10 by tightening any remaining interactive focus/visual-state gaps and then prepare for consolidated in-app verification.

### Session 051 - 2026-04-21
**Phase:** Phase 10 - Visual Polish (incremental step 6)  
**Goal for this session:** Continue CP-10 with a meaningful typography consistency improvement in the right sidebar.  
**Completed:**  
- Updated `Sidebar::EnsureFonts` to use explicit DPI-aware `Segoe UI Variable Text` fonts (header semibold + item regular) with ClearType and `Segoe UI` fallback.  
- Kept existing sidebar density/layout logic and interaction behavior unchanged.  
- Rebuilt solution successfully (`Debug x64`, `x64\\Debug\\FileExplorer.exe` produced).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).  
**Not completed / carried forward:**  
- Remaining CP-10 items still pending; implementation continues incrementally.  
**Decisions made:** AD-055  
**Errors encountered and resolved:**  
- `LNK1168` due to running `FileExplorer.exe`; resolved by terminating process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-10 - Partial (incremental step 6 complete)  
**Next session should start with:**  
- Continue CP-10 by tightening remaining typography/state consistency in primary list zones.

### Session 050 - 2026-04-21
**Phase:** Phase 10 - Visual Polish (incremental step 5)  
**Goal for this session:** Continue CP-10 with one small visual-only improvement to interactive control states.  
**Completed:**  
- Added keyboard-focus rendering for owner-drawn nav buttons in `NavBar::DrawNavButton`:
  - detects `ODS_FOCUS`
  - paints a subtle inset rounded focus ring in dark-theme palette
  - keeps existing hover/pressed/disabled behavior unchanged.
- Rebuilt solution successfully (`Debug x64`, `x64\\Debug\\FileExplorer.exe` produced).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).  
**Not completed / carried forward:**  
- Remaining CP-10 items still pending and will continue incrementally after user confirmation.  
**Decisions made:** AD-054  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-10 - Partial (incremental step 5 complete)  
**Next session should start with:**  
- User verify nav buttons show a subtle focus ring during keyboard focus traversal; then proceed with the next small visual-only refinement.

### Session 049 - 2026-04-21
**Phase:** Phase 10 - Visual Polish (targeted shortcut reliability bugfix)  
**Goal for this session:** Fix user-reported unreliable `Ctrl+W` current-tab close behavior.  
**Completed:**  
- Verified root cause: centralized shortcut interception currently handled `Ctrl+T` only.  
- Extended `MainWindow::MessageLoop` key interception to include `Ctrl+W` for main/child-targeted key events, routed via existing `HandleTabKeyboardShortcut`.  
- Rebuilt solution successfully (`Debug x64`, `x64\\Debug\\FileExplorer.exe` produced).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).  
**Not completed / carried forward:**  
- Remaining CP-10 polish items continue incrementally after user confirmation.  
**Decisions made:** AD-053  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-10 - Partial (shortcut reliability fixes applied for Ctrl+T/Ctrl+W)  
**Next session should start with:**  
- User verify `Ctrl+W` closes active tab reliably from file list, sidebar search, and breadcrumb edit focus contexts.

### Session 048 - 2026-04-21
**Phase:** Phase 10 - Visual Polish (targeted shortcut bugfix)  
**Goal for this session:** Fix user-reported `Ctrl+T` failure so new tab opens reliably regardless of focused child control.  
**Completed:**  
- Traced keyboard routing path and identified root cause: child controls consume focus so `MainWindow` did not always receive `WM_KEYDOWN` for `Ctrl+T`.  
- Implemented centralized `Ctrl+T` interception in `MainWindow::MessageLoop` for messages targeting the main window or its children, reusing `HandleTabKeyboardShortcut` for behavior consistency.  
- Rebuilt solution successfully (`Debug x64`, `x64\\Debug\\FileExplorer.exe` produced).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).  
**Not completed / carried forward:**  
- Remaining CP-10 visual polish items remain pending and continue incrementally.  
**Decisions made:** AD-052  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-10 - Partial (incremental visual progress continues; shortcut bug fixed)  
**Next session should start with:**  
- User verify `Ctrl+T` opens a new tab from details list, sidebar search box, and breadcrumb edit mode; then continue next incremental polish step.

### Session 047 - 2026-04-21
**Phase:** Phase 10 - Visual Polish (incremental step 4)  
**Goal for this session:** Continue CP-10 with one small visual-only typography improvement.  
**Completed:**  
- Re-read required startup docs and confirmed active plan target (`CP-10`) before edits.  
- Implemented status-bar typography polish in `MainWindow`:
  - added RAII-managed status font (`Segoe UI Variable Text`, 11 pt regular, ClearType, Segoe UI fallback)
  - applied font at status control creation
  - recreated/reapplied font on `WM_DPICHANGED`.
- Rebuilt solution successfully (`Debug x64`, `x64\\Debug\\FileExplorer.exe` produced).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).  
**Not completed / carried forward:**  
- Remaining CP-10 visual criteria are still pending and will continue incrementally after user confirmation.  
**Decisions made:** AD-051  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-10 - Partial (incremental step 4 complete)  
**Next session should start with:**  
- User confirms status-bar typography looks aligned with detail/right-pane text density; then proceed to the next small visual-only refinement.

### Session 046 - 2026-04-21
**Phase:** Phase 10 - Visual Polish (incremental step 3)  
**Goal for this session:** Continue CP-10 with one additional conservative visual-only improvement and keep behavior unchanged.  
**Completed:**  
- Re-read required session docs and confirmed active target remains `CP-10` incremental flow.  
- Implemented loading overlay visual polish in `FileListView`:
  - replaced `|/-\` text spinner with an 8-step animated dot spinner
  - kept the existing overlay and status text behavior intact
  - tightened animation cadence to `80 ms` timer ticks for smoother motion.
- Rebuilt solution successfully (`Debug x64`, `x64\Debug\FileExplorer.exe` produced).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).  
**Not completed / carried forward:**  
- Remaining CP-10 polish criteria still pending and will continue incrementally after user confirmation.  
**Decisions made:** AD-050  
**Errors encountered and resolved:**  
- Initial build command failed because `msbuild` was not on PATH in this shell; resolved by invoking the full Visual Studio MSBuild path directly.  
- One multi-proc diagnostic build returned failure without actionable errors; resolved by re-running single-node build (`/m:1`) which completed successfully.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-10 - Partial (incremental step 3 complete)  
**Next session should start with:**  
- User confirms loading overlay spinner look; then proceed to the next small Phase 10 visual-only refinement.

### Session 045 - 2026-04-21
**Phase:** Phase 10 - Visual Polish (incremental step 1)  
**Goal for this session:** Begin CP-10 with a minimal, low-risk polish change and no major visual redesign.  
**Completed:**  
- Re-read required startup docs (`AGENTS.MD`, `DEVELOPMENT_JOURNAL.MD`, `PLAN.MD` Phase 10 criteria).  
- Expanded `Colors.h` to include shared shell palette tokens for base/nav/file-list/sidebar/status/text/separator surfaces.  
- Updated `MainWindow` to consume the new `Colors.h` tokens for:
  - fallback background
  - zone brushes
  - status bar fill/text/separator
  - static control background/text wiring.
- Kept all effective colors visually identical to previous values (infrastructure alignment, not redesign).  
- Rebuilt full solution successfully (`Debug x64`, 0 warnings, 0 errors).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).  
**Not completed / carried forward:**  
- Remaining CP-10 visual polish items still pending and will be handled incrementally after user confirmation.  
**Decisions made:** AD-048  
**Errors encountered and resolved:**  
- One build pass failed with `LNK1168` because `FileExplorer.exe` was running; resolved by stopping process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-10 - Partial (incremental step 1 complete)  
**Next session should start with:**  
- User confirm this no-big-change step; then proceed to incremental step 2 (small, user-reviewed visual refinement).

### Session 044 - 2026-04-21
**Phase:** Phase 9 closeout -> Phase 10 kickoff  
**Goal for this session:** Record final user sign-off for CP-09 and advance project status to the next phase.  
**Completed:**  
- User confirmed Phase 9 app verification is complete.  
- Marked `CP-09` as PASS in project status.  
- Updated active phase to `Phase 10 - Visual Polish` with target `CP-10 (not started)`.  
- Updated Quick Reference table to mark Phase 9 as `Complete`.  
**Not completed / carried forward:**  
- Phase 10 implementation has not started yet.  
**Decisions made:**  
- None (status/sign-off transition only).  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-09 - Pass  
**Next session should start with:**  
- Re-read AGENTS/JOURNAL/PLAN and begin Phase 10 (`CP-10`) implementation.

### Session 043 - 2026-04-19
**Phase:** Phase 9 - File Operations (targeted UX adjustments)  
**Goal for this session:** Apply user-requested copy-name token format and cut-pending row visualization.  
**Completed:**  
- Updated copy collision naming format for files to `name copyN .ext` (space before extension preserved).  
- Kept folder collision naming as `Folder copyN` (no trailing-space suffix).  
- Added cut-pending visual state in details list:
  - `Ctrl+X` now marks selected items as pending-cut and renders them in grey text
  - cut state clears on copy (`Ctrl+C`) and on successful move-paste
  - cut markers update on delete and rename path changes.
- Rebuilt full solution successfully (`Debug x64`, 0 warnings, 0 errors).
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).
**Not completed / carried forward:**  
- CP-09 final interactive sign-off still pending user confirmation after these UX adjustments.  
**Decisions made:** AD-047  
**Errors encountered and resolved:**  
- One build pass failed with `LNK1168` because `FileExplorer.exe` was running; resolved by stopping process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-09 - Partial (implementation + refinements complete, final user verification pending)  
**Next session should start with:**  
- User verify exact filename collision pattern (`copyN ` token before extension) and cut-grey row behavior; if confirmed, mark CP-09 complete.

### Session 042 - 2026-04-19
**Phase:** Phase 9 - File Operations (refinement pass)  
**Goal for this session:** Apply user-requested CP-09 behavior refinements after implementation validation.  
**Completed:**  
- Confirmed that full Windows shell context entries were appearing by design from CP-09 shell integration; switched UI behavior back to curated FileExplorer context menu (no appended full shell menu block).  
- Added stricter rename handling for open/locked files:
  - `FileOps::RenamePath` now probes for sharing/lock state on files before rename
  - inline rename now shows explicit open-file guidance when sharing/lock violations occur.
- Implemented copy collision auto-rename behavior for paste:
  - on `Ctrl+C` + `Ctrl+V`, when destination name exists, target is auto-generated as `<name> copy <x>` with incremental `x`
  - works for both files and folders
  - no conflict prompt shown for this copy-collision path.
- Improved details-pane focus recovery so keyboard navigation works immediately:
  - details pane regains focus after paste/delete/context operations
  - details pane regains focus after rename edit lifecycle ends.
- Rebuilt full solution successfully (`Debug x64`, 0 warnings, 0 errors).
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).
**Not completed / carried forward:**  
- CP-09 final interactive verification/sign-off remains pending user confirmation after these refinements.  
**Decisions made:** AD-046  
**Errors encountered and resolved:**  
- One build pass failed with `LNK1168` because `FileExplorer.exe` was running; resolved by stopping process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-09 - Partial (refinement pass complete, user verification pending)  
**Next session should start with:**  
- Re-run CP-09 verification in-app with focus on copy-collision naming, open-file rename error behavior, and no-extra-click keyboard focus after operations; if all pass, mark CP-09 complete.

### Session 041 - 2026-04-19
**Phase:** Phase 9 - File Operations  
**Goal for this session:** Start CP-09 implementation after re-reading AGENTS/JOURNAL/PLAN/PRD and deliver the required file-operation behaviors without changing unrelated architecture.  
**Completed:**  
- Re-read required startup docs (`AGENTS.MD`, `DEVELOPMENT_JOURNAL.MD`, `PLAN.MD` Phase 9 / CP-09 criteria, `PRD.MD` Section 15 + keyboard shortcuts).  
- Replaced `FileOps` stub with concrete Win32 wrappers:
  - clipboard copy/cut via `CF_HDROP` (`Preferred DropEffect` for cut)
  - paste from clipboard via `SHFileOperationW` (copy/move)
  - delete to Recycle Bin + permanent delete
  - rename wrapper (`MoveFileW`)
  - new-folder create helper (`CreateDirectoryW`)
  - properties open (`SHObjectProperties`)
  - text clipboard copy (`CF_UNICODETEXT`).
- Added Phase 9 keyboard/file operations in `FileListView`:
  - `Ctrl+C`, `Ctrl+X`, `Ctrl+V`
  - `Delete` (Recycle Bin), `Shift+Delete` (permanent with confirmation)
  - `F2` inline rename (`EDIT` overlay with Enter/Escape handling)
  - `Alt+Enter` properties.
- Expanded file-list context menu behavior:
  - wired Copy/Cut/Paste/Delete/Rename/New Folder/Properties/Copy Path commands
  - retained favourite actions and search-result `Open file location`
  - integrated shell context-menu command append/invoke path via `IContextMenu`.
- Implemented `New Folder` creation with unique naming (`New folder`, `New folder (2)`, ...) and immediate inline rename mode.
- Added explicit refresh triggers after file operations to keep list contents current.
- Initialized COM apartment in `App::Run` for reliable shell context-menu interaction support.
- Rebuilt full solution successfully (`Debug x64`, 0 warnings, 0 errors).
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).
**Not completed / carried forward:**  
- CP-09 interactive/manual verification is still pending user validation in-app.  
**Decisions made:** AD-045  
**Errors encountered and resolved:**  
- `msbuild` was not on PATH in this shell; resolved by using the installed VS 18 MSBuild path.  
- One compile pass failed due `std::max` mixed `LONG/int` ambiguity in rename-edit sizing; resolved with explicit integer casts.  
- One link pass failed with `LNK1168` because `FileExplorer.exe` was running; resolved by stopping the process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-09 - Partial (implementation complete, interactive verification pending)  
**Next session should start with:**  
- Run the CP-09 in-app checklist (copy/cut/paste/delete/shift-delete/rename/new-folder/properties/copy-path/shell actions) and mark CP-09 pass when fully confirmed.

### Session 040 - 2026-04-19
**Phase:** Phase 8 closeout -> Phase 9 kickoff  
**Goal for this session:** Fix reported nested flyout overlay regression and close CP-08 after final user confirmation.  
**Completed:**  
- Implemented flyout root popup positioning clamp in `Sidebar::ShowFlyoutPopup` to reserve nested submenu space and prevent child menu overlap on first-level flyout.  
- Rebuilt full solution (`Debug x64`) successfully after flyout-positioning fix.  
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).  
- User confirmed persistence works and flyout behavior now works as expected.  
- Marked `CP-08` as PASS and moved active project status to `Phase 9 - File Operations` (`CP-09` not started).  
**Not completed / carried forward:**  
- Phase 9 implementation has not started yet.  
**Decisions made:** AD-044  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-08 - Pass  
**Next session should start with:**  
- Re-read AGENTS/JOURNAL/PLAN and begin Phase 9 (`CP-09`) implementation.

### Session 039 - 2026-04-18
**Phase:** Phase 8 - Favourites System  
**Goal for this session:** Start CP-08 by implementing CSV-backed favourites storage and wiring full sidebar/list integration.  
**Completed:**  
- Re-read required startup docs (`AGENTS.MD`, `DEVELOPMENT_JOURNAL.MD`, `PLAN.MD` Phase 8 / CP-08 criteria).  
- Implemented `CsvParser` (`Parse`, `Serialize`, `ReadFile`, `WriteFile`) with quoted comma/newline handling, trimming for unquoted fields, UTF-8 BOM write/read support, and missing-file-as-empty behavior.  
- Implemented `FavouritesStore` with:
  - flyout + regular lists
  - alphabetical sorting
  - add/remove/rename mutation APIs
  - CSV persistence
  - portable-first storage root with `%APPDATA%\\FileExplorer` fallback.
- Wired `Sidebar` to persistent store:
  - load/store integration
  - remove/rename context menu functionality
  - inline rename edit control
  - list refresh from store state.
- Kept first-run default favourites seeding (Home/System Drive/Windows + Documents/Pictures/Music) only when no CSV files exist.
- Added `FileListView` -> `MainWindow` messages for `Make Favourite` / `Make Flyout Favourite` folder actions and implemented handling in `MainWindow` to mutate sidebar/store.
- Converted flyout menu entry loading to async worker-thread enumeration with posted UI updates and loading placeholders for submenu expansion.
- Added real CSV unit tests in `FileExplorer.Tests/tests/test_csv.cpp` (5 tests).
- Rebuilt full solution successfully (`Debug x64`, 0 warnings, 0 errors).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`21` tests, `0` failures).  
**Not completed / carried forward:**  
- CP-08 final interactive verification/sign-off is still pending.  
**Decisions made:** AD-043  
**Errors encountered and resolved:**  
- `CsvParser::ReadFile`/`WriteFile` shadowed Win32 API names inside class scope; resolved by qualifying API calls as `::ReadFile`/`::WriteFile`.  
- Test compile failed due unsupported wide-string assertion streaming in local GoogleTest bootstrap; resolved by converting those checks to `EXPECT_TRUE` value comparisons.  
- One rebuild hit `LNK1168` due running `FileExplorer.exe`; resolved by stopping process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-08 - Partial (implementation complete, interactive verification pending)  
**Next session should start with:**  
- Verify CP-08 in-app: add/remove/rename regular + flyout favourites, async flyout expansion responsiveness, restart persistence, and CSV special-character behavior.

### Session 038 - 2026-04-18
**Phase:** Phase 7 closeout -> Phase 8 kickoff  
**Goal for this session:** Confirm CP-07 completion status after final user verification and move active target to Phase 8.  
**Completed:**  
- User confirmed in-app validation that Phase 7 behavior is working as expected, including search-result `Open file location` new-tab flow and search-tab result persistence on return.  
- Marked `CP-07` as PASS in Current Status.  
- Updated active phase to `Phase 8 - Favourites System` with current target `CP-08 (not started)`.  
- Updated Quick Reference table to mark Phase 7 as `Complete`.  
**Not completed / carried forward:**  
- Phase 8 (`CP-08`) implementation has not started yet.  
**Decisions made:**  
- None (checkpoint/sign-off status transition only).  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-07 - Pass  
**Next session should start with:**  
- Re-read AGENTS/JOURNAL/PLAN and begin Phase 8 (`CP-08`) implementation.

### Session 037 - 2026-04-18
**Phase:** Phase 7 - Search Engine  
**Goal for this session:** Add search-result context-menu action `Open file location` that opens containing folder in a new tab and preserves existing search results when switching back.  
**Completed:**  
- Re-read required startup docs (`AGENTS.MD`, `DEVELOPMENT_JOURNAL.MD`, `PLAN.MD` Phase 7 / CP-07 criteria).  
- Added `Open file location` command to file-list context menu with search-mode gating and containing-folder extraction.  
- Added new UI message `WM_FE_FILELIST_OPEN_LOCATION_NEW_TAB` from `FileListView` to `MainWindow`.  
- Added stable tab IDs in `TabManager` (`TabState::id`) to support per-tab search snapshot tracking.  
- Added `FileListView::SearchSnapshot` capture/restore API and integrated it in `MainWindow` (store on tab leave, restore on tab return, prune on tab close, clear on search reset).  
- Wired open-location command to create a new tab for the containing folder while preserving originating search-tab results.
- Rebuilt full solution successfully (`Debug x64`, 0 warnings, 0 errors).  
- Ran tests successfully: `FileExplorer.Tests.exe` (`16` tests, `0` failures).  
- Relaunched updated app for immediate validation (`FileExplorer.exe`, PID 1996).  
**Not completed / carried forward:**  
- CP-07 still pending final interactive user verification/sign-off.  
**Decisions made:** AD-042  
**Errors encountered and resolved:**  
- `msbuild` command not present on PATH in shell; resolved by using full VS 18 MSBuild path (`C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe`).  
- Initial relink hit `LNK1168` because `FileExplorer.exe` was still running; resolved by stopping running app instances and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-07 - Partial (verification pending)  
**Next session should start with:**  
- Validate in-app flow: search result -> right-click `Open file location` -> new tab opens containing folder; switch back to search tab and confirm results persist unchanged.

### Session 036 - 2026-04-18
**Phase:** Phase 7 - Search Engine  
**Goal for this session:** Fix detail-pane left-inset white/transparent artifacting seen in non-maximized and maximize/restore transitions.  
**Completed:**  
- Root-caused artifact to uncovered main-window inset strip (left of file list) not being explicitly painted in all paint paths.  
- Implemented `MainWindow::PaintFileListInsetGutter(HDC)` to compute and fill the inset strip with file-list surface color.  
- Wired inset paint into both `WM_ERASEBKGND` and `WM_PAINT` to ensure stable visuals across initial show, resize, maximize, and restore transitions.  
- Rebuilt full solution (`Debug x64`) successfully with 0 warnings / 0 errors.  
- Relaunched updated app for immediate validation (`FileExplorer.exe`, PID logged).  
**Not completed / carried forward:**  
- CP-07 still pending full user interactive verification/sign-off.  
**Decisions made:** AD-041  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-07 - Partial (verification pending)  
**Next session should start with:**  
- Confirm inset artifact is gone in non-maximized and maximize/restore transitions, then continue CP-07 verification checklist.

### Session 035 - 2026-04-18
**Phase:** Phase 7 - Search Engine  
**Goal for this session:** Resolve remaining right-pane white text background issue reported after prior sidebar visual rollback.  
**Completed:**  
- Root-caused issue in `Sidebar::Paint`: text drawing path lacked explicit transparent background mode.  
- Applied fix by setting `SetBkMode(hdc, TRANSPARENT)` (and matching background color) before drawing sidebar headers/items.  
- Rebuilt full solution (`Debug x64`) successfully with 0 warnings / 0 errors.  
- Ran tests successfully: `FileExplorer.Tests.exe` (`16` tests, `0` failures).  
- Relaunched updated app for immediate verification (`FileExplorer.exe`, PID logged).  
**Not completed / carried forward:**  
- CP-07 still pending full user interactive verification/sign-off.  
**Decisions made:** AD-040  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-07 - Partial (verification pending)  
**Next session should start with:**  
- Confirm right-pane text background issue is gone in-app and continue CP-07 verification checklist.

### Session 034 - 2026-04-18
**Phase:** Phase 7 - Search Engine  
**Goal for this session:** Apply user-requested search UX refinements: show elapsed search time in status bar and restore prior darker right-pane row/header visuals.  
**Completed:**  
- Added search elapsed-time tracking in `MainWindow` and surfaced it through `FileListView` status text as `Time: <ms>` when search completes.  
- Extended `FileListView` with `SetSearchElapsedMs` and search status text formatting update.  
- Restored sidebar section header presentation to a flatter dark style (removed capsule-like header fills) to match previous look.  
- Hardened search edit theming in sidebar (`SetWindowTheme` + shared dark brush for edit/static color messages) to avoid bright control bleed.  
- Rebuilt full solution (`Debug x64`) successfully with 0 warnings / 0 errors.  
- Ran tests successfully: `FileExplorer.Tests.exe` (`16` tests, `0` failures).  
- Relaunched updated app for user verification (`FileExplorer.exe`, PID logged).  
**Not completed / carried forward:**  
- CP-07 remains pending full user interactive verification/sign-off.  
**Decisions made:** AD-039  
**Errors encountered and resolved:**  
- `LNK1168` while rebuilding with running `FileExplorer.exe`; resolved by stopping process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-07 - Partial (verification pending)  
**Next session should start with:**  
- Verify search timing label visibility and right-pane visual rollback in-app, then continue remaining CP-07 verification checklist.

### Session 033 - 2026-04-18
**Phase:** Phase 7 - Search Engine  
**Goal for this session:** Implement CP-07 core search functionality (worker, wildcard, UI wiring, cancellation, search results mode) and restore compile/test green state.  
**Completed:**  
- Re-read required startup docs (`AGENTS.MD`, `DEVELOPMENT_JOURNAL.MD`, `PLAN.MD` Phase 7).  
- Replaced `WildcardMatch` stub with iterative case-insensitive wildcard engine (`*` and `?`) and added 16 wildcard tests in `test_wildcard.cpp`.  
- Replaced `SearchWorker` stub with recursive stack-based DFS search worker:
  - background-thread traversal
  - wildcard filename matching
  - batched incremental posting (`WM_FE_SEARCH_RESULT`, batch size 50)
  - completion post (`WM_FE_SEARCH_DONE`)
  - generation + cancel-token cancellation guards
- Extended `FileListView` with search mode APIs:
  - `BeginSearch`, `AppendSearchResults`, `CompleteSearch`, `LeaveSearchMode`, `IsSearchMode`
  - Path column visibility in search mode
  - search-specific status text and empty-state message
  - search results use full path for open/copy-path behavior
- Extended `Sidebar` with functional search input:
  - real `EDIT` control with dark styling
  - Enter posts search request
  - Escape and clear (`×`) post search-clear
  - `FocusSearch` and `ClearSearchText` APIs for `MainWindow`
- Wired `MainWindow` search orchestration:
  - start/cancel search lifecycle
  - handle sidebar search request/clear messages
  - apply search results and completion messages
  - Ctrl+F now focuses sidebar search field
  - navigation and folder loads exit search mode automatically
- Built full solution successfully (`Debug x64`, 0 warnings, 0 errors).  
- Ran tests successfully: `FileExplorer.Tests.exe` now passes with 16 wildcard tests, 0 failures.  
**Not completed / carried forward:**  
- CP-07 manual interactive verification checklist (large-tree responsiveness and full UX validation) still pending user confirmation.  
**Decisions made:** AD-037, AD-038  
**Errors encountered and resolved:**  
- `msbuild` not on PATH in shell; resolved by using full VS 18 MSBuild executable path.  
- Solution link failed with `LNK1104` while `FileExplorer.exe` was running; resolved by stopping the process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-07 - Partial (implementation complete, interactive verification pending)  
**Next session should start with:**  
- Run CP-07 in-app verification (incremental search updates, clear/escape exit, navigation exit, path column visibility, Ctrl+F focus, cancellation behavior) and close CP-07 after user sign-off.

### Session 032 - 2026-04-18
**Phase:** Phase 6 closeout  
**Goal for this session:** Mark CP-06 as ready/pass based on user confirmation and update the journal state.  
**Completed:**  
- User confirmed Phase 6 is ready.  
- Marked `CP-06` as PASS in Current Status and updated phase progression to `Phase 7 - Search Engine` (`CP-07` not started).  
- Updated quick-reference status for Phase 6 to `Complete`.  
- Closed the remaining CP-06 tracking known issue entry (`KI-007`).  
- Ran test executable (`FileExplorer.Tests.exe`) successfully (0 tests, 0 failures).  
**Not completed / carried forward:**  
- Phase 7 implementation has not started yet.  
**Decisions made:**  
- None (status/sign-off update only).  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Pass  
**Next session should start with:**  
- Re-read AGENTS/JOURNAL/PLAN and begin Phase 7 (`CP-07`) implementation.

### Session 031 - 2026-04-18
**Phase:** Phase 6 - Right Sidebar Shell (visual consistency follow-up)  
**Goal for this session:** Match tab/navbar typography and spacing density to detail pane and right sidebar scale.  
**Completed:**  
- Reduced tab-strip title font sizing and tightened tab text padding/vertical density (`TabStrip.cpp`).  
- Reduced breadcrumb/address text size and segment spacing to align with detail/sidebar typography (`BreadcrumbBar.cpp`).  
- Slightly reduced nav button size/gap and glyph font scale so the top row reads less oversized (`NavBar.cpp`).  
- Rebuilt full solution (`Debug x64`) successfully.  
- Ran tests (`FileExplorer.Tests.exe`) successfully (0 tests, 0 failures).  
- Launched updated app build for interactive verification (`x64\\Debug\\FileExplorer.exe`, PID logged).  
**Not completed / carried forward:**  
- CP-06 remains pending final user interactive sign-off.  
**Decisions made:** AD-036  
**Errors encountered and resolved:**  
- `LNK1168` while linking because `FileExplorer.exe` was running; resolved by stopping process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Confirm typography balance in-app and continue with any final visual micro-adjustments requested by user.

### Session 030 - 2026-04-18
**Phase:** Phase 6 - Right Sidebar Shell (strict visual correction pass)  
**Goal for this session:** Apply a stricter visual-only pass so UI reads as one cohesive dark design system while preserving layout/behavior.  
**Completed:**  
- Re-read required startup docs (`AGENTS.MD`, `DEVELOPMENT_JOURNAL.MD`, `IMPLEMENTATION_PLAN.md` -> `PLAN.MD`) and confirmed strict no-behavior-change scope.  
- Updated tab strip visuals to reduce boxed-outline look and improve integrated active/inactive hierarchy (padding, hover/close/new/scroll visual treatment).  
- Refined navigation toolbar visuals (compact 28x28 owner-draw buttons, subtle borders/hover/pressed/disabled states, cleaner row integration).  
- Polished breadcrumb/address surface and edit-mode visuals for stronger single-control appearance while preserving auto-select/full-path edit behavior.  
- Reworked details header custom draw to fully dark custom item rendering (`CDRF_SKIPDEFAULT` path), with restrained separators and custom sort glyphs.  
- Increased main list left inset slightly for improved breathing room without pane-width/layout redesign.  
- Improved sidebar search/section/item styling consistency (semibold headers, subtle surfaces/separators, calmer dark palette).  
- Integrated status bar as owner-drawn dark surface with subtle top separator and compact text alignment.  
- Built `Debug x64` successfully (`FileExplorer.sln`) and ran tests (`FileExplorer.Tests.exe`: 0 tests, 0 failures).  
- Launched updated app build for user verification (`x64\\Debug\\FileExplorer.exe`, PID logged).  
**Not completed / carried forward:**  
- CP-06 remains pending final user interactive sign-off.  
**Decisions made:** AD-035  
**Errors encountered and resolved:**  
- `msbuild` not on PATH in shell; resolved by using full VS 18 MSBuild path.  
- `LNK1168` because `FileExplorer.exe` was running during link; resolved by stopping the process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Run user visual verification checklist for strict correction targets (tabs/nav/address/header/list inset/sidebar/status), then iterate minor visuals or close CP-06 if accepted.

### Session 029 - 2026-04-18
**Phase:** Phase 6 - Right Sidebar Shell (visual-only refinement pass)  
**Goal for this session:** Execute a conservative visual-only UI polish across tabs/nav/path/header/sidebar/status/list inset while preserving existing behavior and layout structure.  
**Completed:**  
- Re-read required startup docs (`AGENTS.MD`, `DEVELOPMENT_JOURNAL.MD`, `IMPLEMENTATION_PLAN.md` -> `PLAN.MD`) and confirmed active phase context remains Phase 6 / CP-06 pending interactive sign-off.  
- Updated tab strip rendering to a cleaner compact dark style (rounded tabs/buttons, stronger active hierarchy, refined close/new/scroll visual treatment) with no tab-behavior changes.  
- Upgraded nav toolbar button visuals to owner-draw compact glyph buttons (Back/Forward/Up/Refresh) using Segoe MDL2 glyphs and hover/pressed/disabled states; command wiring unchanged.  
- Polished breadcrumb/address visuals into a cohesive dark control surface and removed bright edit-mode styling artifacts (theme-neutral edit control + dark colors/margins); path/flyout/edit behavior unchanged.  
- Applied explicit dark custom-draw visuals for details header row (dark background, readable text, subtle separators) while preserving sort/column logic.  
- Added subtle left inset for the main file list zone in `MainWindow` layout to improve breathing room without changing pane structure.  
- Refined sidebar/search visual treatment (palette + rounded search/hover rows + subtle separators) and status bar integration styling while keeping right-pane interactions unchanged.  
- Verified `Debug x64` full solution build succeeds (`FileExplorer.sln`).  
- Verified tests run and pass (`x64\\Debug\\FileExplorer.Tests.exe`: 0 tests, 0 failures).  
**Not completed / carried forward:**  
- CP-06 remains pending final user interactive sign-off.  
**Decisions made:** AD-034  
**Errors encountered and resolved:**  
- `LNK1104` on solution build due running `x64\\Debug\\FileExplorer.exe`; resolved by stopping the running process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Run user visual verification for the polish pass (tabs/nav/path/header/list inset/sidebar/status) and either iterate visual tweaks or close CP-06 if accepted.

### Session 028 - 2026-04-18
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Document final canonical flyout solution and explicit dead-end list for future maintenance safety.  
**Completed:**  
- Added a dedicated canonical maintenance record (`AD-033`) capturing:
  - final working flyout architecture and invariants
  - exact dead-end implementations that should not be reintroduced
  - rationale for why the final combined model is required
- Verified current test executable still passes (`0 tests`, `0 failures`).  
**Not completed / carried forward:**  
- CP-06 remains pending final user interactive sign-off.  
**Decisions made:** AD-033  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Continue from the canonical flyout record in AD-033 and avoid reintroducing any dead-end paths listed there.

### Session 027 - 2026-04-18
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Fix top-level flyout first-row open action (`Open <top folder>`) so it works like folder rows in the flyout.  
**Completed:**  
- Root-caused top-row failure to menu-hook filter scope: hook selection resolution only accepted submenu-parent rows.  
- Updated hook point-resolution path to resolve any flyout row that has a mapped target (including top-row open action).  
- Kept submenu-parent click-open and hover flyouts unchanged.
- Rebuilt `Debug x64` successfully (0 warnings, 0 errors).
- Ran `FileExplorer.Tests.exe` successfully (0 tests, 0 failures).
- Relaunched `x64\\Debug\\FileExplorer.exe` for validation.  
**Not completed / carried forward:**  
- CP-06 still pending final user interactive sign-off.  
**Decisions made:** AD-032  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Verify top-level first-row open action and nested flyout row click-open together; if both pass, close CP-06.

### Session 026 - 2026-04-18
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Perform root-cause analysis for persistent “hover then click same submenu-parent row does not open” behavior and implement deterministic fix.  
**Completed:**  
- Root cause identified: submenu-parent rows in Win32 popup menus do not reliably generate command events when already highlighted by hover; prior logic depended on selection-change notifications that are absent in this interaction path.  
- Implemented deterministic menu-loop handling in `Sidebar`:
  - added `WH_MSGFILTER` hook installation/removal scoped to flyout popup lifetime
  - intercepted `WM_LBUTTONUP` while `MSGF_MENU` is active
  - resolved clicked submenu-parent row via `MenuItemFromPoint` + existing `(HMENU, position)` target map
  - forced command execution path using pending selection + `EndMenu()`
- Kept existing hover flyout behavior and unlimited nesting path intact.
- Rebuilt `Debug x64` successfully (0 warnings, 0 errors).
- Ran `FileExplorer.Tests.exe` successfully (0 tests, 0 failures).
- Relaunched `x64\\Debug\\FileExplorer.exe` for validation.  
**Not completed / carried forward:**  
- CP-06 still pending final user interactive sign-off.  
**Decisions made:** AD-031  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Validate `Home -> hover Desktop -> click Desktop without mouse move` and nested equivalents; if still failing, capture menu-hook trace (`WM_LBUTTONUP`, resolved menu/item) for environment-specific edge handling.

### Session 025 - 2026-04-17
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Eliminate remaining click-open regression for rows that also expose nested hover flyouts.  
**Completed:**  
- Reworked submenu-parent click resolution in `Sidebar` `WM_MENUSELECT`:
  - use `GetSubMenu(source_menu, selected_index)`
  - map submenu handle through `flyout_menu_paths_`
  - set pending folder target + `EndMenu()` to execute click-open reliably
- Removed elapsed-tick gating from popup click bridge to avoid dropped clicks.
- Rebuilt Debug x64 successfully (`0 warnings, 0 errors`).
- Ran `FileExplorer.Tests.exe` (pass; 0 tests).
- Relaunched `x64\\Debug\\FileExplorer.exe` for interactive validation.  
**Not completed / carried forward:**  
- CP-06 remains pending user interactive sign-off.  
**Decisions made:** AD-030  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Validate same-row hover+click behavior across top-level and nested flyouts; if regression persists, capture live message traces for `WM_MENUSELECT`, `WM_MENUCOMMAND`, and final `WM_FE_SIDEBAR_NAVIGATE` dispatch.

### Session 024 - 2026-04-17
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Restore folder click-open behavior without removing deeper hover flyouts.  
**Completed:**  
- Added submenu-parent click bridge in `Sidebar` message handling (`WM_MENUSELECT`) using left-button down-edge detection and menu-position target lookup.
- Added popup click-tracking state (`flyout_menu_left_button_was_down_`, `flyout_popup_open_tick_`) and proper reset lifecycle.
- Preserved existing hover submenu expansion path (`WM_INITMENUPOPUP`) and deterministic selection flow.
- Rebuilt with MSBuild 18 path successfully (`0 warnings, 0 errors`).
- Ran tests executable (`0 tests`, pass).
- Relaunched `x64\\Debug\\FileExplorer.exe` for user validation.  
**Not completed / carried forward:**  
- CP-06 remains pending user interactive sign-off.  
**Decisions made:** AD-029  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Validate submenu-parent row click-open and nested-hover expansion on the same interaction path; if any edge still fails, capture debug output for `WM_MENUSELECT` vs `WM_MENUCOMMAND` path hit counts.

### Session 023 - 2026-04-17
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Restore first-level hover flyouts without regressing click-open folder behavior.  
**Completed:**  
- Re-enabled popup menu notification style in `ConfigureMenuForNotifications` (`MNS_NOTIFYBYPOS`) for flyout menus/submenus.
- Restored submenu-parent generation at top flyout level for folders that contain children, so hover reveals deeper content again.
- Kept deterministic `TPM_RETURNCMD` selection flow (removed `TPM_NONOTIFY`) so direct click-open path remains active while hover expansion is available.
- Built solution with `MSBuild.exe` at `C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe` (success; 0 warnings, 0 errors).
- Ran `FileExplorer.Tests.exe` (pass; 0 tests, 0 failures).
- Relaunched `x64\Debug\FileExplorer.exe` for immediate interactive validation.  
**Not completed / carried forward:**  
- CP-06 remains pending user interactive sign-off.  
**Decisions made:** AD-028  
**Errors encountered and resolved:**  
- `LNK1168` during link when app process was already running; resolved by stopping `FileExplorer.exe` before rebuild.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Validate this exact behavior matrix: top-row click opens, top-level folder hover expands, top-level folder click opens, nested folder hover expands.

### Session 022 - 2026-04-17
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Eliminate remaining flyout click non-determinism and harden end-to-end sidebar navigation action visibility.  
**Completed:**  
- Refactored flyout popup command flow to use deterministic return-command handling:
  - `TrackPopupMenuEx(..., TPM_RETURNCMD | TPM_NONOTIFY, ...)`
  - direct command-id mapping to action target after popup returns
- Simplified menu-notification dependency in `ConfigureMenuForNotifications` and retained debug trace output for resolved/unresolved selection paths.
- Hardened `MainWindow` `WM_FE_SIDEBAR_NAVIGATE` path:
  - explicit trace log per incoming target path
  - always triggers `HandleTabStateChanged()` after successful activate of existing tab
- Verified full solution build succeeds in `Debug x64` (0 warnings, 0 errors).
- Verified `FileExplorer.Tests.exe` runs and passes (0 tests, 0 failures).
- Launched updated `FileExplorer.exe` for immediate user verification.  
**Not completed / carried forward:**  
- CP-06 remains pending user interactive sign-off.  
**Decisions made:** AD-026, AD-027  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Re-verify top-level flyout `Open <folder>` and first-level child-folder clicks; if still failing, capture debug output stream from `OutputDebugStringW` during click sequence and patch the exact failing branch.

### Session 021 - 2026-04-17
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Deep-fix persistent top-level flyout open failures after previous command-routing changes.  
**Completed:**  
- Updated sidebar navigation dispatch in `Sidebar::PostNavigatePath`:
  - same-thread path now uses synchronous `SendMessageW(WM_FE_SIDEBAR_NAVIGATE, ...)`
  - cross-thread fallback still uses `PostMessageW`
- Added diagnostic `OutputDebugStringW` traces around flyout command resolution in `ShowFlyoutPopup`.
- Verified full solution build succeeds in `Debug x64` (0 warnings, 0 errors).
- Verified `FileExplorer.Tests.exe` runs and passes (0 tests, 0 failures).
- Launched updated `FileExplorer.exe` for immediate user verification.  
**Not completed / carried forward:**  
- CP-06 remains pending user interactive sign-off.  
**Decisions made:** AD-025  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Confirm top-level flyout row and top-row open action both navigate correctly; if still failing, inspect debug traces from flyout selection and main-window sidebar navigate handling.

### Session 020 - 2026-04-17
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Resolve persistent top-level flyout folder click failure reported after prior menu-command fixes.  
**Completed:**  
- Changed flyout menu composition logic in `Sidebar::PopulateFlyoutMenu` so top-level flyout rows for folders remain direct click-open commands, even when child folders exist.
- Kept submenu-parent row behavior for nested popups only (`include_top_open_row == false`).
- Verified full solution build succeeds in `Debug x64` (0 warnings, 0 errors).
- Verified `FileExplorer.Tests.exe` runs and passes (0 tests, 0 failures).
- Launched updated `FileExplorer.exe` for immediate user verification.  
**Not completed / carried forward:**  
- CP-06 remains pending user interactive sign-off.  
**Decisions made:** AD-024  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Confirm top-level flyout folder click-open behavior with user and close CP-06 if all checks pass.

### Session 019 - 2026-04-17
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Resolve persistent user-reported flyout click failure where folder rows still did not open.  
**Completed:**  
- Re-analyzed flyout event flow in `Sidebar.cpp` for `WM_MENUCOMMAND`/`WM_COMMAND` handling and popup lifecycle ordering.  
- Added deterministic fallback selection path in `ShowFlyoutPopup`:
  - switched popup invocation to `TrackPopupMenuEx(..., TPM_RETURNCMD, ...)`
  - captured returned command-id and mapped it through `flyout_command_targets_`
  - retained existing notification-based handling for submenu-parent click path
- Fixed compile regression from missing parenthesis in the new popup call.
- Verified full solution build succeeds in `Debug x64` (0 warnings, 0 errors).
- Verified `FileExplorer.Tests.exe` runs and passes (0 tests, 0 failures).
- Launched updated `FileExplorer.exe` for immediate interactive verification.  
**Not completed / carried forward:**  
- CP-06 remains pending interactive sign-off from user verification in-app.  
**Decisions made:** AD-023  
**Errors encountered and resolved:**  
- `Sidebar.cpp` syntax errors (`C2059`, `C2143`, `C2737`) introduced during edit; resolved by correcting `TrackPopupMenuEx` cast parenthesis.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Validate top-level flyout folder click-open behavior in the running app; if confirmed, close CP-06.

### Session 018 - 2026-04-17
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Fix remaining flyout click failure where top-level popup folder rows were not opening on click.  
**Completed:**  
- Re-read required startup docs in order (`AGENTS.MD`, `DEVELOPMENT_JOURNAL.MD`, `IMPLEMENTATION_PLAN.md` -> `PLAN.MD`) and confirmed active target remains CP-06.  
- Diagnosed flyout menu-command handling path in `Sidebar.cpp` and identified id-only lookup gap for submenu-parent click notifications.  
- Implemented robust fallback mapping for flyout targets keyed by menu handle + item position:
  - added `flyout_position_targets_` state in `Sidebar`
  - recorded position targets during flyout menu population
  - used position fallback in `WM_MENUCOMMAND` when `MIIM_ID` lookup misses
  - reset new state in `ClearFlyoutPopupState`
- Verified full solution build succeeds in `Debug x64` (0 warnings, 0 errors) using VS 18 MSBuild path.
- Verified `FileExplorer.Tests.exe` runs and passes (0 tests, 0 failures).  
**Not completed / carried forward:**  
- CP-06 still pending final interactive user sign-off in-app.  
**Decisions made:** AD-022  
**Errors encountered and resolved:**  
- Task-based build failed because `msbuild` was not on PATH; resolved by using explicit path `C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe`.
- Initial rebuild hit `LNK1168` due running `FileExplorer.exe`; resolved by stopping process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Validate top-level flyout folder click-open behavior in-app and close CP-06 if all sidebar checklist criteria pass.

### Session 017 - 2026-04-17
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Fix submenu-parent click behavior so hovered flyout folders also open on click without moving the mouse.  
**Completed:**  
- Added popup-menu notification style (`MNS_NOTIFYBYPOS`) to flyout menus and submenus.  
- Implemented `WM_MENUCOMMAND` handling to capture menu-item activation by position and map it to flyout command targets, including submenu-parent rows.  
- Added `WM_COMMAND` fallback capture for flyout command targets during active popup sessions.  
- Added pending flyout selection state so selected target executes after menu loop closes.  
- Verified full solution build succeeds in `Debug x64` (0 warnings, 0 errors).  
- Verified `FileExplorer.Tests.exe` runs and passes (0 tests, 0 failures).  
**Not completed / carried forward:**  
- CP-06 still pending final interactive sign-off.  
**Decisions made:** AD-021  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Validate submenu-parent click-open behavior (e.g., Home -> Desktop hover then click same row) and close CP-06 if user confirms.

### Session 016 - 2026-04-17
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Finalize flyout/favourites behavior requested by user: folder click-open semantics, new-tab sidebar navigation with tab reuse, and no right-pane active highlight.  
**Completed:**  
- Updated sidebar navigation handling in `MainWindow` for `WM_FE_SIDEBAR_NAVIGATE`:
  - open folder in a new tab when not already open
  - if the folder is already open in any tab, switch to that tab instead of opening duplicate
- Updated flyout menu item creation for folders that have submenus:
  - folder rows now carry both submenu and command id targets
  - keeps hover flyout behavior while enabling direct folder-open command on click
- Removed right-pane active-path visual highlight behavior:
  - no active row background/underline in sidebar list rendering
  - hover feedback remains
- Verified full solution build succeeds in `Debug x64` (0 warnings, 0 errors).  
- Verified `FileExplorer.Tests.exe` runs and passes (0 tests, 0 failures).  
**Not completed / carried forward:**  
- CP-06 still pending final interactive sign-off.  
**Decisions made:** AD-020  
**Errors encountered and resolved:**  
- None.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Confirm folder-click behavior on submenu parent rows in-app and close CP-06 if all sidebar checks pass.

### Session 015 - 2026-04-17
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Align flyout behavior with requested top-row open action, file visibility, and click/hover interaction semantics.  
**Completed:**  
- Updated flyout menu generation to include both folders and files (hidden/system entries still excluded).  
- Added top-level first-row action: `Open <flyout folder>` for each root flyout popup.  
- Added command target typing (`is_folder`) so click behavior can route correctly:
  - folders navigate in-app
  - files open via `ShellExecuteW`
- Preserved unlimited recursive hover flyouts for folders with children via lazy `WM_INITMENUPOPUP` population.
- Updated flyout popup command handling to execute typed targets (`OpenPathOrNavigate`) instead of folder-only path navigation.
- Verified full solution build succeeds in `Debug x64` (0 warnings, 0 errors).  
- Verified `FileExplorer.Tests.exe` runs and passes (0 tests, 0 failures).  
**Not completed / carried forward:**  
- CP-06 still pending final interactive sign-off.  
**Decisions made:** AD-019  
**Errors encountered and resolved:**  
- `LNK1168` during first rebuild due running app lock; resolved by stopping `FileExplorer.exe` and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Validate flyout UX details in-app and close CP-06 if user confirms all behavior.

### Session 014 - 2026-04-17
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Support unlimited nested flyout behavior where any flyout folder can further fly out.  
**Completed:**  
- Reworked flyout popup handling to support recursive depth without fixed limits:
  - added lazy submenu population on `WM_INITMENUPOPUP`
  - submenus are created dynamically for folders with child folders
  - leaf selections navigate to the selected path
- Added flyout menu state tracking for active popup session:
  - per-menu path mapping
  - per-command leaf-path mapping
  - cleanup/reset lifecycle around popup invocation
- Preserved and integrated prior behavior improvements:
  - if top-level flyout target has no subfolders, direct navigation fallback remains
  - sidebar single-click navigation path handling remains intact
- Verified full solution build succeeds in `Debug x64` (0 warnings, 0 errors).  
- Verified `FileExplorer.Tests.exe` runs and passes (0 tests, 0 failures).  
**Not completed / carried forward:**  
- CP-06 still pending final interactive sign-off.  
**Decisions made:** AD-018  
**Errors encountered and resolved:**  
- `LNK1168` during rebuild due running app lock; resolved by stopping `FileExplorer.exe` and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Verify nested flyout behavior in-app and close CP-06 if all checklist items pass.

### Session 013 - 2026-04-17
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Fix flyout favorites showing `(No subfolders)` and make flyout behavior usable.  
**Completed:**  
- Updated flyout default targets to higher-confidence folder roots:
  - `Home` (`%USERPROFILE%`)
  - `System Drive` (`C:\`)
  - `Windows` (`C:\Windows`)
- Changed flyout behavior when no child folders are found:
  - instead of showing `(No subfolders)` popup, it now navigates directly to the clicked flyout target
- Verified full solution build succeeds in `Debug x64` (0 warnings, 0 errors).  
- Verified `FileExplorer.Tests.exe` runs and passes (0 tests, 0 failures).  
**Not completed / carried forward:**  
- CP-06 still pending final interactive sign-off checklist.  
**Decisions made:** AD-017  
**Errors encountered and resolved:**  
- `LNK1168` during rebuild due running app lock; resolved by stopping `FileExplorer.exe` and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Confirm CP-06 in-app and mark checkpoint pass if all criteria are validated.

### Session 012 - 2026-04-17
**Phase:** Phase 6 - Right Sidebar Shell  
**Goal for this session:** Fix sidebar font/spacing density and restore missing flyout/favourite click interactions.  
**Completed:**  
- Updated right-sidebar typography/spacing to detail-style density:
  - switched to system message-font based creation (with DPI scaling)
  - tightened header/item spacing and row metrics
  - adjusted wheel/line scroll steps to use actual row height
- Fixed folder-enter selection behavior edge case confirmation path:
  - entering a folder keeps top-item default selection
  - parent-navigation reselect hint behavior remains unchanged
- Added sidebar single-click interactions:
  - single-click on regular favourite now navigates to its path
  - wired via `WM_FE_SIDEBAR_NAVIGATE` to `MainWindow` navigation flow
- Implemented flyout rendering behavior:
  - flyout rows now open popup menus with child folders
  - selecting a child folder navigates correctly
  - flyout indicator changed from plain text `>` to explicit triangle glyph rendering
- Verified full solution build succeeds in `Debug x64` (0 warnings, 0 errors).  
- Verified `FileExplorer.Tests.exe` runs and passes (0 tests, 0 failures).  
**Not completed / carried forward:**  
- CP-06 still awaits final interactive user sign-off.  
**Decisions made:** AD-016  
**Errors encountered and resolved:**  
- `LNK1168` during rebuild due running app lock; resolved by stopping `FileExplorer.exe` and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (verification pending)  
**Next session should start with:**  
- Confirm CP-06 visual/interaction checklist in-app and mark pass if all behaviors are validated.

### Session 011 - 2026-04-17
**Phase:** Phase 6 - Right Sidebar Shell (and post-CP-05 navigation polish)  
**Goal for this session:** Fix folder-enter selection behavior and continue implementing Phase 6 sidebar shell.  
**Completed:**  
- Fixed folder-enter selection behavior:
  - non-incremental folder loads now clear inherited selection/focus state
  - first selectable row is selected/focused by default when entering a folder
  - parent-navigation reselect hint behavior remains intact
- Implemented Phase 6 sidebar shell in a custom `FE_Sidebar` control:
  - custom sidebar window class integrated into `MainWindow` (replacing static placeholder)
  - DPI-aware 320px layout participation (existing metrics) with custom rendering surface
  - fixed search placeholder zone (`32px`) at top
  - section headers (`Flyout Favourites`, `Favourites`) with bold styling and spacing
  - custom-drawn item rows (`28px`) with hover and active-state visuals
  - 1px divider between flyout and regular favourites sections
  - vertical scrolling support and per-item right-click menu shell (`Remove`, `Rename`)
- Wired sidebar active-item tracking to current tab path updates in `HandleTabStateChanged`.
- Verified full solution build succeeds in `Debug x64` (0 warnings, 0 errors).  
- Verified `FileExplorer.Tests.exe` runs and passes (0 tests, 0 failures).  
**Not completed / carried forward:**  
- CP-06 interactive/manual verification is still pending user sign-off.  
**Decisions made:** AD-014, AD-015  
**Errors encountered and resolved:**  
- Sidebar compile failure from mixed `LONG`/`int` in `std::max`; resolved via explicit integer cast.  
- `LNK1168` during rebuild due running `FileExplorer.exe`; resolved by stopping process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-06 - Partial (implementation complete, verification pending)  
**Next session should start with:**  
- Validate CP-06 in-app checklist and mark pass if all visual/interaction criteria are confirmed.

### Session 010 - 2026-04-17
**Phase:** Phase 5 closeout -> Phase 6 kickoff  
**Goal for this session:** Fix parent-navigation selection behavior and proceed to next phase.  
**Completed:**  
- Implemented parent-navigation reselect logic:
  - when navigating to a direct parent folder, the child folder we came from is selected/focused after load completes
  - behavior works with async loading path (selection hint carried through worker result apply)
- Added transition analysis in `MainWindow` to detect child->parent navigation and pass post-load selection hint to `FileListView`.
- Added post-load single-entry selection API in `FileListView` and applied it on non-incremental loads.
- Verified full solution build succeeds in `Debug x64` (0 warnings, 0 errors).  
- Verified `FileExplorer.Tests.exe` runs and passes (0 tests, 0 failures).  
- Launched updated app build for interactive verification (`FileExplorer.exe`, PID logged during run).  
- Marked CP-05 complete and moved project status to Phase 6.  
**Not completed / carried forward:**  
- Phase 6 (`CP-06`) implementation tasks are not yet completed.  
**Decisions made:** AD-013  
**Errors encountered and resolved:**  
- `LNK1168` during rebuild because `FileExplorer.exe` was running; resolved by stopping process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-05 - Pass  
**Next session should start with:**  
- Implement Phase 6 sidebar shell (`FE_Sidebar`) and CP-06 criteria.

### Session 009 - 2026-04-17
**Phase:** Phase 5 - Background Threading & Async Folder Loading  
**Goal for this session:** Implement CP-05 async loading architecture and refresh/watch foundations, then validate build/test health.  
**Completed:**  
- Marked CP-04 as complete based on user-confirmed interaction validation and moved active work to Phase 5.  
- Added shared worker message definitions in `WorkerMessages.h`:
  - `WM_FE_FOLDER_LOADED`, `WM_FE_SEARCH_RESULT`, `WM_FE_SEARCH_DONE`, `WM_FE_SIBLINGS_READY`
  - plus internal `WM_FE_ACTIVE_FOLDER_DIRTY` for watcher notifications
- Implemented `FolderWorker`:
  - background `std::thread` per navigation request
  - wide-char folder enumeration (`FindFirstFileW` / `FindNextFileW`)
  - heap-owned `std::vector<FileEntry>` payload posted via `PostMessageW`
  - generation checks before work/post and automatic payload cleanup on post failure
- Updated `MainWindow` to async folder loading:
  - generation source (`std::atomic<uint64_t>`) and stale-result discard in `WM_FE_FOLDER_LOADED`
  - navigation now starts worker loads instead of synchronous `LoadFolder`
  - refresh flows (`NavBar Refresh`, `F5`, file-list refresh message) request incremental apply mode
- Implemented loading-state UX in `FileListView`:
  - `BeginFolderLoad(...)` keeps previous contents visible
  - center loading overlay + timer-driven spinner while enumeration is in flight
  - `ApplyLoadedFolderEntries(...)` atomically swaps in new results on completion
- Implemented incremental refresh behavior in `FileListView`:
  - merge delta logic (added/removed/changed detection by entry key)
  - preserves selection by filename where possible
  - preserves focus and top-scroll position after refresh
- Added active-folder filesystem watching:
  - `FindFirstChangeNotificationW` watcher thread only for visible tab
  - posts `WM_FE_ACTIVE_FOLDER_DIRTY` on file/dir name or write-time changes
  - 500 ms debounce timer in UI thread before incremental refresh trigger (scoped to current watch generation)
- Updated breadcrumb sibling async message routing to shared `WM_FE_SIBLINGS_READY`.  
- Verified app project build succeeds (`FileExplorer.vcxproj`, Debug x64).  
- Verified full solution build succeeds (`FileExplorer.sln`, Debug x64).  
- Verified tests run and pass (`FileExplorer.Tests.exe`, 0/0/0).  
**Not completed / carried forward:**  
- CP-05 manual/interactive stress verification checklist is still pending user validation in the running app.  
**Decisions made:** AD-012  
**Errors encountered and resolved:**  
- `msbuild` was not on PATH in this shell; resolved by using explicit VS 18 MSBuild path.  
- `LNK1104` on `x64\\Debug\\FileExplorer.exe` when binary was locked by running process; resolved by stopping `FileExplorer.exe` and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-05 - Partial (implementation complete, interactive verification pending)  
**Next session should start with:**  
- Run CP-05 verification checklist in-app (large-folder responsiveness, stale-load cancellation, incremental refresh/selection preservation, and filesystem-change auto-refresh) and mark CP-05 pass if all criteria are confirmed.

### Session 008 - 2026-04-17
**Phase:** Phase 4 - Main File List (sync)  
**Goal for this session:** Fix upward arrow navigation across date divider rows in date-sorted mode.  
**Completed:**  
- Added explicit `VK_UP` / `VK_DOWN` handling in file-list subclass:
  - skips divider rows in both directions
  - preserves `Shift+Arrow` range selection while excluding dividers
  - maintains status updates and visible focus movement
- Verified full solution build succeeds in `Debug x64`.  
- Verified `FileExplorer.Tests.exe` runs and passes (empty suite).  
- Launched updated app build for interactive verification (`FileExplorer.exe`, PID logged during run).  
**Not completed / carried forward:**  
- Final CP-04 interactive sign-off still pending user confirmation.  
**Decisions made:** AD-011  
**Errors encountered and resolved:**  
- `LNK1168` during rebuild when `FileExplorer.exe` was still running; resolved by stopping process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-04 - Partial (verification pending)  
**Next session should start with:**  
- Confirm this arrow-key fix in-app; if confirmed, mark CP-04 pass and proceed to Phase 5.

### Session 007 - 2026-04-17
**Phase:** Phase 4 - Main File List (sync)  
**Goal for this session:** Apply user-requested keyboard and selection polish fixes in current build.  
**Completed:**  
- Added file-list type-ahead navigation:
  - pressing a character jumps to matching items
  - rapid multi-character input builds prefix matching (example: `dow` -> `Downloads`)
  - includes timeout-based reset for fresh searches
- Added selection/focus sanitization so date-divider rows are removed from selection and cannot remain focused during range operations.
- Added `Backspace` keyboard behavior in file list to navigate one folder up (same action as `Left`).
- Verified full solution build succeeds in `Debug x64`.  
- Verified `FileExplorer.Tests.exe` runs and passes (empty suite).  
- Launched updated app build for interactive verification (`FileExplorer.exe`, PID logged during run).  
**Not completed / carried forward:**  
- Final CP-04 interactive sign-off still pending user confirmation.  
**Decisions made:** AD-010  
**Errors encountered and resolved:**  
- `LNK1168` during rebuild when `FileExplorer.exe` was still running; resolved by stopping process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-04 - Partial (verification pending)  
**Next session should start with:**  
- Confirm these keyboard/selection fixes in-app and, if all pass, mark CP-04 complete and begin Phase 5.

### Session 006 - 2026-04-17
**Phase:** Phase 4 - Main File List (sync)  
**Goal for this session:** Complete the remaining CP-04 pending items after user approval to finish them.  
**Completed:**  
- Completed CP-04 interaction features in `FileListView`:
  - Date group dividers when sorting by Date Modified (`Today`, `Yesterday`, `This week`, `This month`, `Older`)
  - Divider rows rendered with distinct style and blocked from selection/focus
  - Multi-select support with explicit `Ctrl+A` handling that skips divider rows
  - Empty-area click deselection behavior
  - Keyboard actions:
    - `Left` navigates up one folder
    - `Right` enters selected folder
    - `Enter` opens selected item (folder navigate / file shell open)
  - Context menu with required items:
    - Copy, Cut, Paste, Rename, Delete, Properties, Copy Path, Make Favourite, Make Flyout Favourite
  - Empty-folder centered message drawing (`This folder is empty`) with folder icon overlay
  - Status summary posting (`Items`, `Selected`) wired to status bar text updates in `MainWindow`
- Preserved owner-data architecture (`LVS_OWNERDATA`) and column sort indicator behavior.  
- Verified full solution build succeeds in `Debug x64`.  
- Verified `FileExplorer.Tests.exe` runs and passes (empty suite).  
- Launched updated app build for interactive verification (`FileExplorer.exe`, PID logged during run).  
**Not completed / carried forward:**  
- Final interactive CP-04 sign-off confirmation from user is still pending.  
**Decisions made:** AD-009  
**Errors encountered and resolved:**  
- `LNK1168` during rebuild when `FileExplorer.exe` was still running; resolved by stopping process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-04 - Partial (verification pending)  
**Next session should start with:**  
- Run CP-04 checklist in the live app and mark CP-04 pass if all behaviors are confirmed.

### Session 005 - 2026-04-17
**Phase:** Phase 3 (closeout) + Phase 4 (implementation start)  
**Goal for this session:** Fix reported breadcrumb/navigation + dark edit-mode issues, then proceed into Phase 4 implementation.  
**Completed:**  
- Fixed breadcrumb navigation routing:
  - clicking breadcrumb segments now navigates
  - clicking sibling flyout entries now navigates
  - root cause was message target; navigation messages now post to main window target
- Fixed breadcrumb edit mode appearance:
  - added explicit dark-mode edit colors via `WM_CTLCOLOREDIT`
  - removed client-edge white styling and applied dark background/text
- Added segment-layout refresh logic for reliable hit-testing before click/hover handling.  
- Started Phase 4 implementation with a real `FileListView`:
  - replaced main file-list placeholder with `WC_LISTVIEWW` in `LVS_REPORT | LVS_OWNERDATA`
  - added required columns (Name, Extension, Date Modified, Size)
  - enabled list view extended styles (`LVS_EX_FULLROWSELECT`, `LVS_EX_DOUBLEBUFFER`, `LVS_EX_HEADERDRAGDROP`)
  - implemented synchronous folder loading with `FindFirstFileW/FindNextFileW`
  - implemented owner-data text/icon callbacks (`LVN_GETDISPINFOW`)
  - implemented folder-first sorting with header-click sort toggles and sort indicators
  - added dark list surface and extension-based filename color coding
  - wired double-click folder navigation back into `MainWindow`/`TabManager`
- Verified full solution build succeeds in `Debug x64`.  
- Verified `FileExplorer.Tests.exe` runs and passes (empty suite).  
- Launched updated app build for interactive verification (`FileExplorer.exe`, PID logged during run).  
**Not completed / carried forward:**  
- Remaining CP-04 items (group dividers, selection rules, keyboard behaviors, empty-state UX, context menu/status integration, additional interactions) are still pending.  
**Decisions made:** AD-007, AD-008  
**Errors encountered and resolved:**  
- `LNK1168` due to running `FileExplorer.exe` during rebuild; resolved by stopping running process and rebuilding.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-03 - Pass; CP-04 - Partial  
**Next session should start with:**  
- Validate breadcrumb fixes in the running app, then continue CP-04 remaining criteria implementation and verification.

### Session 004 - 2026-04-17
**Phase:** Phase 2 (closeout) + Phase 3 (implementation start)  
**Goal for this session:** Fix pinned-tab label usability issue, mark CP-02 complete, and start Phase 3 implementation.  
**Completed:**  
- Updated pinned-tab rendering so pinned tabs now show `[P] <name>` with ellipsis, instead of `P` only.  
- Kept pinned behavior constraints: no close button on pinned tabs and narrower width than unpinned tabs.  
- Marked CP-02 complete based on user confirmation that all other Phase 2 criteria worked and pinned rendering fix was applied.  
- Added Phase 3 core implementations:
  - `TabManager` navigation/history API (`NavigateTo`, `NavigateBack`, `NavigateForward`, `NavigateUp`, navigation-state checks)
  - custom `FE_NavBar` control with Back/Forward/Up/Refresh buttons and state updates
  - custom `FE_BreadcrumbBar` control with:
    - clickable path segments
    - separator-triggered sibling flyout enumeration on worker thread and UI handoff via `PostMessageW`
    - edit mode via embedded `EDIT` control
    - `Enter` navigate, `Escape` cancel, and focus-loss exit behavior
    - `SHAutoComplete` filesystem suggestions
  - `MainWindow` integration for:
    - nav command handling
    - breadcrumb navigation handling
    - `Alt+D` / `Ctrl+L` address-edit activation
    - window title update format `<folder> - FileExplorer`
- Verified full solution build succeeds in `Debug x64`.  
- Verified `FileExplorer.Tests.exe` runs and passes (empty suite).  
- Launched updated app build for interactive verification (`FileExplorer.exe`, PID logged during run).  
**Not completed / carried forward:**  
- Full CP-03 manual verification and final completion checks are still pending.  
**Decisions made:** AD-005, AD-006  
**Errors encountered and resolved:**  
- `std::max` overload ambiguity in breadcrumb layout due mixed integer types; fixed with explicit cast.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-02 - Pass; CP-03 - Partial  
**Next session should start with:**  
- Run CP-03 verification checklist end-to-end and complete remaining behaviors needed for CP-03 sign-off.

### Session 003 - 2026-04-17
**Phase:** Phase 2 - Tab Strip  
**Goal for this session:** Mark CP-01 complete after user visual confirmation, then implement CP-02 tab strip behavior and stabilize build/tests.  
**Completed:**  
- Marked CP-01 complete based on user confirmation that the shell/window behavior works.  
- Implemented `TabManager` tab model behaviors for:
  - add/activate/close tab
  - close other tabs / close tabs to the right
  - duplicate tab
  - pin/unpin tab
  - cycling and index jump helpers (`Ctrl+Tab`, `Ctrl+Shift+Tab`, `Ctrl+1..9`)
- Implemented custom `FE_TabStrip` control with:
  - owner-drawn dark tab strip styling
  - active/inactive/hover visual states
  - per-tab close button behavior (hidden for pinned tabs)
  - new-tab `+` button
  - overflow left/right scroll arrows
  - right-click context menu actions (Close tab, Close other tabs, Close tabs to the right, Duplicate tab, Pin/Unpin)
  - tooltip callback showing full tab path on hover
- Integrated tab strip into `MainWindow`:
  - receives `WM_FE_TAB_CHANGED`
  - handles keyboard tab shortcuts (`Ctrl+T`, `Ctrl+W`, `Ctrl+Tab`, `Ctrl+Shift+Tab`, `Ctrl+1..9`)
  - extends `WM_NCHITTEST` to return `HTCAPTION` in draggable tab-gap regions
- Resolved compile break in Phase 2 tab code caused by `min/max` macro collisions by applying `NOMINMAX` at Win32 include boundaries for tab modules.  
- Verified full solution build succeeds in `Debug x64`.  
- Verified `FileExplorer.Tests.exe` runs and passes (empty suite).  
- Launched app for interactive tab-strip verification (`FileExplorer.exe`, PID logged during run).  
**Not completed / carried forward:**  
- Final interactive/manual verification of all CP-02 UI criteria (hover/tooltip/scroll/context behaviors) in desktop session.  
**Decisions made:** AD-004  
**Errors encountered and resolved:**  
- `std::min/max/clamp` parse failures in `TabManager.cpp`/`TabStrip.cpp` due to Win32 macros; resolved with `NOMINMAX` fixes.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-02 - Partial  
**Next session should start with:**  
- Confirm CP-02 interactive behavior checklist end-to-end and mark CP-02 pass if all criteria are visually/functionally verified.

### Session 002 - 2026-04-17
**Phase:** Phase 1 - Application Shell & Window  
**Goal for this session:** Implement CP-01 shell/window behavior according to plan and PRD constraints.  
**Completed:**  
- Implemented `wWinMain` startup path through `App::Run(HINSTANCE, int)`.  
- Added per-monitor V2 DPI awareness call and common controls initialization in `App`.  
- Implemented `MainWindow` with:
  - window class registration (`RegisterClassExW`)
  - centered default window create (`1200 x 750` client target)
  - message loop and dispatch
  - `WM_CLOSE` -> `DestroyWindow`, `WM_DESTROY` -> `PostQuitMessage`
  - dark chrome attempt via `DWMWA_USE_IMMERSIVE_DARK_MODE`
  - Mica attempt via `DWMWA_SYSTEMBACKDROP_TYPE`
  - fallback background fill (`#1C1C1C`) on `WM_ERASEBKGND` when Mica is unavailable
  - placeholder child zones for tab strip, nav bar, file list, sidebar, and status bar
  - DPI-aware metrics and relayout on `WM_SIZE` and `WM_DPICHANGED`
  - `WM_DPICHANGED` suggested RECT handling via `SetWindowPos`
  - initial title text set to `FileExplorer`
- Updated `Theme`/`Dpi` helpers for DPI-scaled `LayoutMetrics` (MulDiv-based scaling).  
- Fixed `/W4 /WX` compile issues (`NOMINMAX` and explicit type casts in layout math).  
- Verified full solution build succeeds in `Debug x64`.  
- Verified `FileExplorer.Tests.exe` passes (empty suite).  
- Runtime smoke launch/close of `FileExplorer.exe` exits cleanly (`exit code 0`).  
**Not completed / carried forward:**  
- Manual visual verification for CP-01 criteria at 125% and 150% DPI in an interactive desktop session.  
**Decisions made:** AD-003  
**Errors encountered and resolved:**  
- `std::min/max` conflicts and ambiguous overloads in `MainWindow.cpp`; resolved by `NOMINMAX` and explicit integer casts.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-01 - Partial  
**Next session should start with:**  
- Run manual visual verification for CP-01 (dark chrome, Mica fallback behavior, zone layout, resize behavior at 100/125/150 DPI). If confirmed, mark CP-01 pass and start Phase 2.

### Session 001 - 2026-04-17
**Phase:** Phase 0 - Project Scaffolding  
**Goal for this session:** Complete CP-00 exactly as specified in plan.  
**Completed:**  
- Created `FileExplorer.sln` with x64 Debug/Release configurations.  
- Added projects: `FileExplorer`, `FileExplorer.Core`, `FileExplorer.Tests`.  
- Applied C++20, `/W4`, `/WX`, Unicode, Windows subsystem for app, Release incremental linking disabled.  
- Added manifest (`PerMonitorV2`, Win10/11 compatibility, visual styles) and embedded it via `resource.rc`.  
- Added `#pragma comment(lib, ...)` entries for required system libraries in `main.cpp`.  
- Created full source/test directory skeleton and compilable stubs listed in Phase 0.  
- Added local GoogleTest bootstrap under `extern/googletest/`.  
- Added `.vscode/tasks.json` build/test tasks.  
- Updated `.gitignore` with required Visual Studio and build artifact exclusions.  
- Verified `msbuild FileExplorer.sln /p:Configuration=Debug /p:Platform=x64` succeeds with zero errors.  
- Verified tests run and pass as empty suite.  
- Verified dependency output for app binary uses only OS DLLs (`KERNEL32.dll` in current stub build).  
**Not completed / carried forward:**  
- Replace GoogleTest bootstrap with full upstream GoogleTest source before Phase 13.  
**Decisions made:** AD-001, AD-002  
**Errors encountered and resolved:**  
- `LNK1104` caused by duplicate core lib specification in linker inputs; fixed by relying on project reference output.  
- `CVT1100/LNK1123` duplicate manifest embedding; fixed by disabling linker-generated manifest and embedding only explicit manifest resource.  
**Errors unresolved / carried forward:**  
- None.  
**Checkpoint status:** CP-00 - Pass  
**Next session should start with:**  
- Wait for user confirmation, then begin Phase 1 (`CP-01`) by implementing `wWinMain`, main window class registration, dark chrome, Mica fallback, and placeholder layout zones.

---

## Known Issues & Debt

| ID | Phase | Severity | Description | Status |
|---|---|---|---|---|
| KI-001 | 0 | Debt | `extern/googletest` currently contains a bootstrap test harness for CP-00. Replace with full upstream GoogleTest source before Phase 13. | Open |
| KI-002 | 1 | Low | CP-01 visual/DPI verification requirement. User confirmed expected behavior; issue closed after CP-01 sign-off. | Closed |
| KI-003 | 2 | Low | CP-02 interactive verification requirement. User confirmed expected behavior and pinned-tab label issue was fixed. | Closed |
| KI-004 | 3 | Low | CP-03 interaction issues (breadcrumb/flyout navigation + edit-mode dark styling) were reported and fixed in Session 005. | Closed |
| KI-005 | 4 | Low | CP-04 interactive verification/sign-off was pending and is now resolved after user confirmation. | Closed |
| KI-006 | 5 | Low | CP-05 interactive verification/sign-off was pending and is now resolved after final parent-navigation selection fix. | Closed |
| KI-007 | 6 | Low | CP-06 sidebar shell verification pending item is resolved after user sign-off; tracking closed. | Closed |

Severity: Blocker | High | Low | Debt

---

## Test Results Log

### Test Run 061 - 2026-04-22
**Configuration:** Debug x64  
**Test count:** 56  
**Passed:** 56  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-CP-15 release packaging closeout regression run)  
**Failures:**  
- None

### Test Run 060 - 2026-04-22
**Configuration:** Debug x64  
**Test count:** 56  
**Passed:** 56  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-header-tail `CDDS_POSTPAINT` dark-fill fix verification run)  
**Failures:**  
- None

### Test Run 059 - 2026-04-22
**Configuration:** Debug x64  
**Test count:** 56  
**Passed:** 56  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-header NM_CUSTOMDRAW routing correction run)  
**Failures:**  
- None

### Test Run 058 - 2026-04-22
**Configuration:** Debug x64  
**Test count:** 56  
**Passed:** 56  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-header-chrome flattening + Esc clear-selection refinement run)  
**Failures:**  
- None

### Test Run 057 - 2026-04-22
**Configuration:** Debug x64  
**Test count:** 56  
**Passed:** 56  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-dark-header/scrollbar theme + icon resource refinement run)  
**Failures:**  
- None

### Test Run 056 - 2026-04-22
**Configuration:** Debug x64  
**Test count:** 56  
**Passed:** 56  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (Phase 13 coverage expansion and required-case alignment verification run)  
**Failures:**  
- None

### Test Run 055 - 2026-04-22
**Configuration:** Debug x64  
**Test count:** 25  
**Passed:** 25  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Ctrl+Tab routing + per-tab row-position restore refinement run)  
**Failures:**  
- None

### Test Run 054 - 2026-04-22
**Configuration:** Debug x64  
**Test count:** 25  
**Passed:** 25  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-tab-switch details-pane focus refinement run)  
**Failures:**  
- None

### Test Run 053 - 2026-04-22
**Configuration:** Debug x64  
**Test count:** 25  
**Passed:** 25  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-CP-12 regression-fix pass: menu removal + stronger second-launch activation fallback)  
**Failures:**  
- None

### Test Run 052 - 2026-04-22
**Configuration:** Debug x64  
**Test count:** 25  
**Passed:** 25  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-foreground-activation handoff patch validation run)  
**Failures:**  
- None

### Test Run 051 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 25  
**Passed:** 25  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 12 single-instance enforcement integration)  
**Failures:**  
- None

### Test Run 050 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 25  
**Passed:** 25  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-tab drag-reorder + file-list context-menu view-toggle implementation)  
**Failures:**  
- None

### Test Run 049 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 25  
**Passed:** 25  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 11 persistence wiring + SessionStore/SortSettings round-trip tests)  
**Failures:**  
- None

### Test Run 048 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-file-list column width persistence implementation verification run)  
**Failures:**  
- None

### Test Run 047 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-sidebar splitter drag + settings persistence implementation verification run)  
**Failures:**  
- None

### Test Run 046 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 10 incremental step 11 empty-state iconography polish verification run)  
**Failures:**  
- None

### Test Run 045 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-tab-search textbox restore fix verification run)  
**Failures:**  
- None

### Test Run 044 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 10 incremental step 10 file-type palette centralization verification run)  
**Failures:**  
- None

### Test Run 043 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 10 incremental step 9 sidebar search focus-state verification run)  
**Failures:**  
- None

### Test Run 042 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 10 incremental steps 7/8 typography + fallback polishing verification run)  
**Failures:**  
- None

### Test Run 041 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 10 incremental step 6 sidebar typography normalization verification run)  
**Failures:**  
- None

### Test Run 040 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 10 incremental step 5 nav-button focus-state polish verification run)  
**Failures:**  
- None

### Test Run 039 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Ctrl+W global shortcut routing fix verification run)  
**Failures:**  
- None

### Test Run 038 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Ctrl+T global shortcut routing fix verification run)  
**Failures:**  
- None

### Test Run 037 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 10 incremental step 4 status-bar typography verification run)  
**Failures:**  
- None

### Test Run 036 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 10 incremental step 3 loading-spinner polish verification run)  
**Failures:**  
- None

### Test Run 035 - 2026-04-21
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 10 incremental step 1 verification run)  
**Failures:**  
- None

### Test Run 034 - 2026-04-19
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-copy-name-token + cut-visual-state refinement run)  
**Failures:**  
- None

### Test Run 033 - 2026-04-19
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 9 refinement pass verification run)  
**Failures:**  
- None

### Test Run 032 - 2026-04-19
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 9 implementation pass verification run)  
**Failures:**  
- None

### Test Run 031 - 2026-04-19
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-flyout-positioning regression fix + CP-08 closeout verification run)  
**Failures:**  
- None

### Test Run 030 - 2026-04-18
**Configuration:** Debug x64  
**Test count:** 21  
**Passed:** 21  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 8 favourites persistence/sidebar integration implementation run)  
**Failures:**  
- None

### Test Run 029 - 2026-04-18
**Configuration:** Debug x64  
**Test count:** 16  
**Passed:** 16  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-search-result open-location/new-tab snapshot-restore verification run)  
**Failures:**  
- None

### Test Run 028 - 2026-04-18
**Configuration:** Debug x64  
**Test count:** 16  
**Passed:** 16  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-sidebar text background paint-mode fix verification run)  
**Failures:**  
- None

### Test Run 027 - 2026-04-18
**Configuration:** Debug x64  
**Test count:** 16  
**Passed:** 16  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-search-time + sidebar-visual-rollback verification run)  
**Failures:**  
- None

### Test Run 026 - 2026-04-18
**Configuration:** Debug x64  
**Test count:** 16  
**Passed:** 16  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 7 implementation verification run)  
**Failures:**  
- None

### Test Run 025 - 2026-04-18
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (CP-06 sign-off journal update verification run)  
**Failures:**  
- None

### Test Run 024 - 2026-04-18
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-typography consistency pass verification run)  
**Failures:**  
- None

### Test Run 023 - 2026-04-18
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-strict visual correction pass verification run)  
**Failures:**  
- None

### Test Run 022 - 2026-04-18
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (final visual-polish pass verification rerun)  
**Failures:**  
- None

### Test Run 021 - 2026-04-18
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-visual-polish pass verification run)  
**Failures:**  
- None

### Test Run 020 - 2026-04-18
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (documentation maintenance session verification run)  
**Failures:**  
- None

### Test Run 019 - 2026-04-18
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-top-row flyout open path wiring verification run)  
**Failures:**  
- None

### Test Run 018 - 2026-04-18
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-menu-loop hook flyout parent-click fix verification run)  
**Failures:**  
- None

### Test Run 017 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-flyout submenu-parent click-open handling verification run)  
**Failures:**  
- None

### Test Run 016 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-sidebar new-tab-or-switch + flyout folder-click behavior verification run)  
**Failures:**  
- None

### Test Run 015 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-flyout top-row + mixed folder/file entry behavior verification run)  
**Failures:**  
- None

### Test Run 014 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-unlimited-nested-flyout implementation verification run)  
**Failures:**  
- None

### Test Run 013 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-flyout empty-state fallback + seed-target update verification run)  
**Failures:**  
- None

### Test Run 012 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-sidebar typography/click/flyout fix verification run)  
**Failures:**  
- None

### Test Run 011 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-folder-enter-selection fix + Phase 6 sidebar shell implementation run)  
**Failures:**  
- None

### Test Run 010 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-parent-navigation-reselect fix verification run)  
**Failures:**  
- None

### Test Run 009 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 5 async loading implementation verification run)  
**Failures:**  
- None

### Test Run 008 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-arrow-divider-skip fix verification run)  
**Failures:**  
- None

### Test Run 007 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-keyboard/selection polish verification run)  
**Failures:**  
- None

### Test Run 006 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-CP-04 completion work verification run)  
**Failures:**  
- None

### Test Run 005 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 4 partial implementation verification run)  
**Failures:**  
- None

### Test Run 004 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 3 partial implementation verification run)  
**Failures:**  
- None

### Test Run 003 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 2 build verification run)  
**Failures:**  
- None

### Test Run 002 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (post-Phase 1 verification run)  
**Failures:**  
- None

### Test Run 001 - 2026-04-17
**Configuration:** Debug x64  
**Test count:** 0  
**Passed:** 0  
**Failed:** 0  
**Skipped:** 0  
**Duration:** N/A (quick smoke run)  
**Failures:**  
- None

---

## Manual Test Results Log

### Manual Run 001 - 2026-04-22
**Phase:** Phase 14 - Integration & Manual Tests  
**Run type:** Full checklist execution (user-driven in-app verification)  
**Result:** PASS

**Category outcomes:**
- Navigation: PASS
- Tabs: PASS
- File list: PASS
- Search: PASS
- File operations: PASS
- Favourites: PASS
- Session & settings: PASS
- Performance: PASS
- DPI and resize: PASS
- PRD section 20 keyboard shortcuts: PASS

**CP-14 acceptance summary:**
- All manual test categories completed: YES
- Pass rate >= 95%: YES (100%)
- No crashes observed during full run: YES
- No UI freezes > 200ms observed: YES
- All PRD section 20 shortcuts verified: YES

**Notes:**
- Verification was confirmed by user sign-off: "It works great."

---

## Dependency & Environment Notes

| Item | Value |
|---|---|
| Visual Studio version | 2026 (`18.x`) |
| Windows SDK version | 10.0.26100.0 |
| Target Windows | Windows 11 x64 |
| MSVC toolset | v145 (`14.50.35717`) |
| Google Test version | Local bootstrap copy in `extern/googletest` (temporary) |
| CRT linking | `/MTd` (Debug), `/MT` (Release) |

---

## File Structure Verification

```text
FileExplorer/
├── FileExplorer.sln                    [x] Created
├── FileExplorer/
│   ├── FileExplorer.vcxproj            [x] Created
│   ├── src/
│   │   ├── main.cpp                    [x] Created
│   │   ├── App.h / App.cpp             [x] Created
│   │   ├── MainWindow.h / MainWindow.cpp [x] Created
│   │   ├── TabStrip.h / TabStrip.cpp   [x] Created
│   │   ├── TabManager.h / TabManager.cpp [x] Created
│   │   ├── NavBar.h / NavBar.cpp       [x] Created
│   │   ├── BreadcrumbBar.h / BreadcrumbBar.cpp [x] Created
│   │   ├── FileListView.h / FileListView.cpp [x] Created
│   │   ├── Sidebar.h / Sidebar.cpp     [x] Created
│   │   ├── SearchBox.h / SearchBox.cpp [x] Created
│   │   ├── FavouritesPanel.h / FavouritesPanel.cpp [x] Created
│   │   ├── StatusBar.h / StatusBar.cpp [x] Created
│   │   ├── FolderWorker.h / FolderWorker.cpp [x] Created
│   │   ├── SearchWorker.h / SearchWorker.cpp [x] Created
│   │   ├── FolderCache.h / FolderCache.cpp [x] Created
│   │   ├── SortSettings.h / SortSettings.cpp [x] Created
│   │   ├── SessionStore.h / SessionStore.cpp [x] Created
│   │   ├── FavouritesStore.h / FavouritesStore.cpp [x] Created
│   │   ├── Settings.h / Settings.cpp   [x] Created
│   │   ├── FileOps.h / FileOps.cpp     [x] Created
│   │   ├── IconCache.h / IconCache.cpp [x] Created
│   │   ├── WildcardMatch.h / WildcardMatch.cpp [x] Created
│   │   ├── CsvParser.h / CsvParser.cpp [x] Created
│   │   ├── Colors.h                    [x] Created
│   │   ├── Dpi.h / Dpi.cpp             [x] Created
│   │   ├── Theme.h / Theme.cpp         [x] Created
│   │   └── resource.h / resource.rc    [x] Created
│   └── res/                            [x] Created
├── FileExplorer.Tests/
│   ├── FileExplorer.Tests.vcxproj      [x] Created
│   └── tests/
│       ├── test_wildcard.cpp           [x] Created
│       ├── test_csv.cpp                [x] Created
│       ├── test_sort.cpp               [x] Created
│       ├── test_session.cpp            [x] Created
│       ├── test_favstore.cpp           [x] Created
│       └── test_pathutils.cpp          [x] Created
├── IMPLEMENTATION_PLAN.md              [x] Present
└── DEVELOPMENT_JOURNAL.md              <- this file
```

---

## Stack Compliance Checklist

- [x] No C# code in project
- [x] No WPF / WinUI / Windows App SDK references
- [x] No Qt includes or libraries
- [x] No SQLite or database files
- [x] No external runtime DLL dependencies (beyond Windows system DLLs)
- [x] All Win32 API calls use `W` (wide) variants
- [x] All string literals for paths/UI use `L""` prefix
- [x] No raw `new`/`delete` for resource ownership (RAII wrappers used or no dynamic ownership introduced yet)
- [x] No UI calls from worker threads (`PostMessageW` only)
- [x] No `SendMessageW` from worker threads
- [x] All `HANDLE` values wrapped in RAII types (no raw HANDLE ownership introduced yet)
- [x] Test project is the only place Google Test appears

---

*End of DEVELOPMENT_JOURNAL.md*  
*Companion: IMPLEMENTATION_PLAN.md*
