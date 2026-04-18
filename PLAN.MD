# FileExplorer — AI Coding Assistant Implementation Plan

**Based on:** PRD v2.5, April 2026  
**Target:** Native Win32 C++20, MSVC, Visual Studio, x64, Windows 11  
**Deployment:** Portable `.exe` / portable folder, no installer  
**Plan version:** 1.0  

---

## ⚠️ MANDATORY STACK REMINDER — READ BEFORE EVERY PHASE

This reminder must be re-read at the start of every phase and every new session.

| Constraint | Rule |
|---|---|
| Language | **C++20 only** |
| Compiler | **MSVC (Visual Studio 2022)** |
| Architecture | **x64 only** |
| UI framework | **Native Win32 API only** |
| Threading | **One UI thread + background worker threads** |
| Persistence | **Lightweight files only (CSV, INI/custom text)** |
| Runtime cache | **In-memory only** |
| Process model | **Single process** |

**The following are PERMANENTLY FORBIDDEN in this project:**

- C#, WPF, WinUI 3, Windows App SDK
- WebView2, Electron, Qt
- SQLite or any external database
- Any background engine process or IPC architecture
- File-content parsing libraries
- External package managers (NuGet is acceptable for test frameworks only)

**All Win32 API calls must use wide-character (`W`) variants.**  
**All string handling must use `std::wstring` or `WCHAR` arrays.**  
**All resource ownership must use RAII (`std::unique_ptr`, custom `Handle` wrappers).**

---

## Companion Document

A `DEVELOPMENT_JOURNAL.md` lives alongside this file in the project root.

**The AI assistant must:**
1. Read `DEVELOPMENT_JOURNAL.md` at the start of every session
2. Update it at the end of every session or after every completed checkpoint
3. Log every architectural decision, deviation, and open issue

The journal is the authoritative record of project state. This plan is the authoritative record of goals and rules.

---

## Project Directory Structure

```
FileExplorer/
├── FileExplorer.sln
├── FileExplorer/
│   ├── FileExplorer.vcxproj
│   ├── src/
│   │   ├── main.cpp
│   │   ├── App.h / App.cpp
│   │   ├── MainWindow.h / MainWindow.cpp
│   │   ├── TabStrip.h / TabStrip.cpp
│   │   ├── TabManager.h / TabManager.cpp
│   │   ├── NavBar.h / NavBar.cpp
│   │   ├── BreadcrumbBar.h / BreadcrumbBar.cpp
│   │   ├── FileListView.h / FileListView.cpp
│   │   ├── Sidebar.h / Sidebar.cpp
│   │   ├── SearchBox.h / SearchBox.cpp
│   │   ├── FavouritesPanel.h / FavouritesPanel.cpp
│   │   ├── StatusBar.h / StatusBar.cpp
│   │   ├── FolderWorker.h / FolderWorker.cpp
│   │   ├── SearchWorker.h / SearchWorker.cpp
│   │   ├── FolderCache.h / FolderCache.cpp
│   │   ├── SortSettings.h / SortSettings.cpp
│   │   ├── SessionStore.h / SessionStore.cpp
│   │   ├── FavouritesStore.h / FavouritesStore.cpp
│   │   ├── Settings.h / Settings.cpp
│   │   ├── FileOps.h / FileOps.cpp
│   │   ├── IconCache.h / IconCache.cpp
│   │   ├── WildcardMatch.h / WildcardMatch.cpp
│   │   ├── CsvParser.h / CsvParser.cpp
│   │   ├── Colors.h
│   │   ├── Dpi.h / Dpi.cpp
│   │   ├── Theme.h / Theme.cpp
│   │   └── resource.h / resource.rc
│   └── res/
│       └── (icons, manifests)
├── FileExplorer.Tests/
│   ├── FileExplorer.Tests.vcxproj
│   └── tests/
│       ├── test_wildcard.cpp
│       ├── test_csv.cpp
│       ├── test_sort.cpp
│       ├── test_session.cpp
│       ├── test_favstore.cpp
│       └── test_pathutils.cpp
├── IMPLEMENTATION_PLAN.md     ← this file
└── DEVELOPMENT_JOURNAL.md     ← always read first
```

---

## Phase Overview

| Phase | Name | Checkpoint |
|---|---|---|
| 0 | Project Scaffolding | `CP-00` |
| 1 | Application Shell & Window | `CP-01` |
| 2 | Tab Strip | `CP-02` |
| 3 | Navigation Bar & Breadcrumb | `CP-03` |
| 4 | Main File List (static, sync) | `CP-04` |
| 5 | Background Threading & Async Loading | `CP-05` |
| 6 | Right Sidebar Shell | `CP-06` |
| 7 | Search Engine | `CP-07` |
| 8 | Favourites System | `CP-08` |
| 9 | File Operations | `CP-09` |
| 10 | Visual Polish & Theme | `CP-10` |
| 11 | Settings & Persistence | `CP-11` |
| 12 | Single-Instance Enforcement | `CP-12` |
| 13 | Unit Tests | `CP-13` |
| 14 | Integration & Manual Tests | `CP-14` |
| 15 | Portable Release Build | `CP-15` |

---

## Phase 0 — Project Scaffolding

**Goal:** An empty, compilable, portable-ready Visual Studio solution exists.

### Tasks

**0.1 — Solution and project creation**
- Create `FileExplorer.sln` in Visual Studio 2022
- Add `FileExplorer` project: Win32 desktop application, C++20, x64 only
- Set subsystem: `Windows` (not console)
- Enable all warnings: `/W4`
- Enable warnings as errors: `/WX` (can relax per-file if needed)
- Set C++ standard: `/std:c++20`
- Disable incremental linking in Release
- Set character set: Unicode (`_UNICODE`, `UNICODE`)
- Output: `FileExplorer.exe`, no DLL dependencies beyond system DLLs

**0.2 — Test project**
- Add `FileExplorer.Tests` project: Console application, C++20, x64
- Integrate Google Test via a local copy in `/extern/googletest/`
  - Do NOT use NuGet if it adds installer dependencies
  - Prefer a Git-submodule or source-copy of gtest
- Link test project to core logic as a static library or shared object files
  - Strategy: Extract all non-UI logic into `FileExplorer.Core` static lib
  - `FileExplorer.exe` links `FileExplorer.Core.lib`
  - `FileExplorer.Tests.exe` also links `FileExplorer.Core.lib`

**0.3 — Manifest and portable config**
- Add application manifest: `FileExplorer.exe.manifest`
  - Declare `supportedOS` for Windows 10/11
  - Enable DPI awareness: `PerMonitorV2`
  - Enable Visual Styles: `comctl32.dll` version 6
- Set `#pragma comment(lib, ...)` for: `comctl32.lib`, `dwmapi.lib`, `uxtheme.lib`, `shell32.lib`, `shlwapi.lib`

**0.4 — Directory skeleton**
- Create all source file stubs as listed in the directory structure above
- Each `.cpp` stub must compile (empty or minimal placeholder)
- Each `.h` stub must have proper `#pragma once` and namespace guards

**0.5 — CI smoke build**
- Confirm: `msbuild FileExplorer.sln /p:Configuration=Debug /p:Platform=x64` succeeds with zero errors

### Checkpoint CP-00

```
CHECKPOINT CP-00 CRITERIA:
[ ] Solution opens and compiles with zero errors in Debug x64
[ ] Test project compiles and runs (empty test suite passes)
[ ] Manifest is present and declares PerMonitorV2 DPI awareness
[ ] Output is a single .exe with no external DLL dependencies beyond OS
[ ] Directory structure matches the plan
[ ] DEVELOPMENT_JOURNAL.md updated: Phase 0 complete, any deviations noted
```

---

## Phase 1 — Application Shell & Window

**Goal:** A resizable Win32 window appears with correct zones, dark background, Mica/DWM effect, and title bar behavior.

### Stack reminder for this phase
- Use `CreateWindowExW` — wide char, always
- DWM: `DwmSetWindowAttribute` with `DWMWA_USE_IMMERSIVE_DARK_MODE` and `DWMWA_SYSTEMBACKDROP_TYPE`
- Mica fallback: if `DwmSetWindowAttribute` for Mica fails, draw a solid dark surface `#1C1C1C`
- DPI: handle `WM_DPICHANGED` for per-monitor DPI scaling from the start

### Tasks

**1.1 — WinMain and message loop**
- Implement `wWinMain` (Unicode entry point)
- Register main window class with `RegisterClassExW`
- Window style: `WS_OVERLAPPEDWINDOW` (resizable, standard chrome)
- Default client area: `1200 × 750` px
- Center window on primary monitor at startup

**1.2 — Dark mode window chrome**
- Call `DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &TRUE, sizeof(TRUE))`
- Attempt Mica backdrop: `DWMWA_SYSTEMBACKDROP_TYPE` = `DWMSBT_MAINWINDOW`
- If Mica is unavailable (Windows build < 22000 or call fails), fill client with `#1C1C1C` via `WM_ERASEBKGND`
- Title bar text: `<current folder name> - FileExplorer` — update via `SetWindowTextW`

**1.3 — Layout zone calculation**
Define `LayoutMetrics` struct (in `Theme.h`):
```cpp
struct LayoutMetrics {
    int tabStripHeight;        // 36 px at 100% DPI
    int navBarHeight;          // 38 px
    int statusBarHeight;       // 24 px
    int sidebarWidth;          // 320 px
    int sidebarResizeGripWidth;// 4 px
    int searchBarHeight;       // 32 px
};
```
- All values scale with DPI: `MulDiv(base, dpi, 96)`
- Recalculate on `WM_DPICHANGED` and `WM_SIZE`

**1.4 — Child zone creation**
Create placeholder child windows (using `CreateWindowExW` with `WC_STATIC` or custom classes) for:
- Tab strip zone (top)
- Navigation bar zone (below tabs)
- Main file list zone (center-left)
- Sidebar zone (right)
- Status bar zone (bottom)

Arrange them correctly in `WM_SIZE`.

**1.5 — DPI handling**
- `SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)`
- In `WM_DPICHANGED`, call `SetWindowPos` with the new suggested RECT
- Recalculate `LayoutMetrics` and resize all child zones

**1.6 — Window close and destroy**
- `WM_CLOSE`: prompt for unsaved state later; for now, call `DestroyWindow`
- `WM_DESTROY`: call `PostQuitMessage(0)`

### Checkpoint CP-01

```
CHECKPOINT CP-01 CRITERIA:
[ ] Window opens with dark chrome, Mica or solid dark fallback
[ ] Window title shows "FileExplorer" initially
[ ] Correct layout zones are visible as placeholder rectangles
[ ] Window resizes correctly — all zones adapt
[ ] DPI changes are handled (no garbled layout at 125%, 150%)
[ ] WM_CLOSE destroys window and exits cleanly
[ ] DEVELOPMENT_JOURNAL.md updated
```

---

## Phase 2 — Tab Strip

**Goal:** A custom-drawn tab strip with add/close/pin behavior, keyboard shortcuts, and right-click context menu.

### Stack reminder
- Custom-drawn owner-draw tabs, NOT `SysTabControl32` — it does not look like Win11 browser tabs
- Drawing in `WM_PAINT` with `BeginPaint`/`EndPaint`
- Hit-test manually in `WM_LBUTTONDOWN`, `WM_LBUTTONUP`, `WM_MOUSEMOVE`
- Use `TrackMouseEvent` for hover detection and leave events

### Data model

```cpp
struct TabState {
    std::wstring path;          // current folder path
    std::wstring displayName;   // folder name for tab label
    bool pinned = false;
    std::vector<std::wstring> history;  // navigation back/forward stack
    int historyIndex = -1;
    // scroll/sort state carried per tab
};

class TabManager {
    std::vector<TabState> tabs_;
    int activeIndex_ = 0;
public:
    // CRUD: add, close, duplicate, pin/unpin
    // activate(index), next(), previous()
    // jumpTo(1-8, last)
};
```

### Tasks

**2.1 — TabStrip custom control class**
- Register custom window class `L"FE_TabStrip"`
- Owner-draw tab items with per-item `RECT` tracking
- Draw tab strip background matching window surface
- Draw active tab with elevated surface (`#2D2D2D`), brighter text
- Draw inactive tabs with muted text, hover highlight

**2.2 — Tab item hit testing**
- On `WM_LBUTTONDOWN`: determine which tab was clicked vs close button vs `+` button
- Store per-tab `RECT` in a list for fast hit testing
- Invalidate on hover change

**2.3 — Tab close button**
- Render `×` glyph or Fluent close icon per tab (except pinned)
- Close button only visible on hover of that tab or active tab
- On click: remove tab from `TabManager`, activate adjacent tab

**2.4 — New tab button (`+`)**
- Positioned after last tab
- On click: `TabManager::addTab()` with default path (e.g., user home)
- Notify `MainWindow` to load the new tab's folder

**2.5 — Pinned tab behavior**
- Pinned tabs: narrower width (~40 px), show pin icon, no close button
- Cannot be closed via close button or `Ctrl+W`
- Toggle pin via tab context menu

**2.6 — Tab context menu**
- Right-click tab → `TrackPopupMenuEx`
- Items: Close tab, Close other tabs, Close tabs to the right, Duplicate tab, Pin/Unpin
- Implement all actions in `TabManager`

**2.7 — Keyboard shortcuts**
- Trap `WM_KEYDOWN` at main window level for tab shortcuts
- `Ctrl+T` → new tab
- `Ctrl+W` → close current (block if last tab or pinned)
- `Ctrl+Tab` / `Ctrl+Shift+Tab` → cycle tabs
- `Ctrl+1`…`Ctrl+8` → jump to tab index
- `Ctrl+9` → last tab

**2.8 — Window drag region**
- Extend `WM_NCHITTEST` so that non-tab areas in the tab strip return `HTCAPTION`
- Tab items themselves return `HTCLIENT` (to catch mouse events)

**2.9 — Tab strip scrolling**
- If tabs overflow the strip width, add horizontal scroll arrows at the ends
- Do not truncate tabs completely

### Checkpoint CP-02

```
CHECKPOINT CP-02 CRITERIA:
[ ] Tab strip renders with correct dark theme styling
[ ] Active tab is visually distinct from inactive tabs
[ ] Hover states work correctly on tabs and close buttons
[ ] New tab button adds a tab
[ ] Close button removes a tab; last tab cannot be closed
[ ] Pinned tabs are narrower, have no close button, cannot be closed
[ ] Right-click context menu works (all 5 items functional)
[ ] Ctrl+T, Ctrl+W, Ctrl+Tab, Ctrl+Shift+Tab all work
[ ] Ctrl+1...9 jump to correct tabs
[ ] Dragging window by tab strip gap area works
[ ] Tooltip shows full path on tab hover
[ ] DEVELOPMENT_JOURNAL.md updated
```

---

## Phase 3 — Navigation Bar & Breadcrumb

**Goal:** Navigation buttons (Back/Forward/Up/Refresh) and a dual-mode breadcrumb/address bar.

### Tasks

**3.1 — NavBar container**
- Custom window class `L"FE_NavBar"`
- Draw background matching nav surface

**3.2 — Navigation buttons**
- Four compact icon buttons: Back, Forward, Up, Refresh
- Use `DrawThemeBackground` or custom painter with Fluent-style glyphs
- States: normal, hover, pressed, disabled
- Wire `WM_COMMAND` IDs: `ID_NAV_BACK`, `ID_NAV_FORWARD`, `ID_NAV_UP`, `ID_NAV_REFRESH`
- Disable Back when `historyIndex == 0`, Forward when at end, Up when at drive root

**3.3 — BreadcrumbBar — display mode**
- Custom control class `L"FE_BreadcrumbBar"`
- Parse current path into segments: `[C:] [›] [Projects] [›] [Current Job] [›] [Docs]`
- Each segment: clickable `RECT` that navigates to that prefix
- Each `›` separator: acts as sibling flyout trigger
- Draw using `TextOutW` with `Segoe UI Variable`, 12 px
- Highlight hovered segment with soft background pill

**3.4 — Sibling flyout**
- On click of `›` between segment N and N+1:
  - Launch background task: enumerate siblings at the level represented by segment N
  - Result posted to UI thread via `PostMessage`
  - Show `TrackPopupMenuEx` flyout populated with sibling folders, sorted alphabetically
  - Item click → navigate to that sibling folder
- Flyout must not block the UI thread

**3.5 — BreadcrumbBar — edit mode**
- Click on non-segment area (or `Alt+D` / `Ctrl+L`) → switch to edit mode
- Show a full-path `EDIT` control (Win32 `WC_EDIT`), pre-filled with current path, full text selected
- Background: slightly elevated, distinct from breadcrumb mode
- Autocomplete suggestions: enumerate real file system via `SHAutoComplete` or custom `WM_CHAR` handler
- `Enter` → navigate; `Escape` → cancel, return to breadcrumb mode
- Losing focus (`WM_KILLFOCUS`) → return to breadcrumb mode

**3.6 — Keyboard wiring**
- `Alt+D` / `Ctrl+L` → activate edit mode on breadcrumb bar (globally trapped)
- `Ctrl+F` → focus search field in sidebar

### Checkpoint CP-03

```
CHECKPOINT CP-03 CRITERIA:
[ ] Back/Forward/Up/Refresh buttons render with correct states
[ ] Back and Forward navigate through per-tab history correctly
[ ] Up navigates to parent folder
[ ] Breadcrumb displays path as clickable segments
[ ] Clicking a segment navigates to that prefix
[ ] Sibling flyout opens on › click with correct siblings listed
[ ] Sibling click navigates immediately
[ ] Clicking blank area of breadcrumb switches to edit mode
[ ] Alt+D / Ctrl+L focus edit mode
[ ] Edit mode shows full path, selected; Enter navigates
[ ] Escape in edit mode restores breadcrumb
[ ] Window title updates to <folder name> - FileExplorer
[ ] DEVELOPMENT_JOURNAL.md updated
```

---

## Phase 4 — Main File List (Static / Synchronous)

**Goal:** The main pane shows folder contents using a virtualized list (owner-data). This phase loads synchronously. Phase 5 makes it async.

### Stack reminder
- Use Win32 `ListView` (`WC_LISTVIEW`) in report mode with `LVS_OWNERDATA` (virtual/owner-data mode)
- `LVS_OWNERDATA` is mandatory for performance with large folders
- `LVN_GETDISPINFO` provides data on demand — never cache data in the ListView itself
- All custom drawing via `NM_CUSTOMDRAW` for row-level color overrides

### Data model

```cpp
struct FileEntry {
    std::wstring name;
    std::wstring extension;
    FILETIME     modifiedTime;
    ULONGLONG    sizeBytes;
    DWORD        attributes;   // FILE_ATTRIBUTE_*
    bool         isFolder;
    // icon index into shared ImageList (resolved lazily)
    int          iconIndex = -1;
};

enum class SortColumn { Name, Extension, DateModified, Size };
enum class SortDirection { Ascending, Descending };
```

**4.1 — ListView setup**
- Create `WC_LISTVIEWW` with styles: `LVS_REPORT | LVS_OWNERDATA | LVS_SHOWSELALWAYS`
- Extended styles: `LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_HEADERDRAGDROP`
- Columns: Name (260 px), Extension (60 px), Date Modified (140 px), Size (90 px)
- Header click → sort; visual chevron in header via `HDM_SETITEM` / `HDF_SORTUP` / `HDF_SORTDOWN`

**4.2 — Owner data (LVN_GETDISPINFO)**
- Maintain `std::vector<FileEntry> displayList_` as the sorted data source
- In `LVN_GETDISPINFO`, return correct text per column, icon index for Name column
- Icon: use `SHGetFileInfo` with `SHGFI_SMALLICON | SHGFI_SYSICONINDEX`, cache in `IconCache`

**4.3 — Folder loading (sync — will be replaced in Phase 5)**
- `LoadFolder(const std::wstring& path)`:
  - Use `FindFirstFileW` / `FindNextFileW` to enumerate
  - Build `std::vector<FileEntry>`
  - Apply hidden/system file filter per settings
  - Sort by current sort column and direction
  - Set `ListView_SetItemCountEx(hwnd, count, LVSICF_NOINVALIDATEALL)`
  - Invalidate list

**4.4 — Sorting**
- On column header click: toggle sort direction if same column, else set new column
- `std::sort` with lambda comparator per `SortColumn`
- Folder-first sort: folders always appear before files regardless of sort column
- Store sort preference per folder (forwarded to `SortSettings` in Phase 11)

**4.5 — File-name color coding**
- Handle `NM_CUSTOMDRAW` at `CDDS_ITEMPREPAINT`
- For each row, inspect `FileEntry.extension`:
  - `.pdf` → `#E06C75` (red)
  - `.xls`, `.xlsx`, `.xlsm` → `#98C379` (green)
  - `.doc`, `.docx` → `#61AFEF` (blue)
  - `.zip`, `.7z`, `.rar`, `.tar`, `.gz` → `#E5C07B` (yellow)
  - `.txt`, `.md`, `.log` → `#56B6C2` (teal)
  - folders → default
  - everything else → default text color
- Color coding only active in dark mode

**4.6 — Selection behavior**
- `LVN_ITEMCHANGED` → update status bar selection summary
- Multi-select: Shift+Click, Ctrl+Click, Ctrl+A (all supported by `LVS_EX_FULLROWSELECT`)
- Click on empty area: deselect all (custom `WM_LBUTTONDOWN` handling)
- Keyboard: Up/Down move focus, Left = navigate up, Right = enter folder, Enter = open item

**4.7 — Date group dividers**
- When sorted by Date Modified, insert `FE_GROUP_DIVIDER` pseudo-entries:
  - Today, Yesterday, This week, This month, Older
- Divider rows: `NM_CUSTOMDRAW` draws them differently (no icon, distinct background, smaller bold text)
- Divider rows are skipped during selection and keyboard navigation

**4.8 — Empty folder state**
- If `displayList_` is empty after loading: draw centered empty-state message
  - Large folder icon (32 px)
  - Text: "This folder is empty"
  - Drawn in `WM_PAINT` fallback when ListView has zero items

**4.9 — Double-click and Enter**
- `NM_DBLCLK` on folder item → `TabManager::navigate(path)`
- `NM_DBLCLK` on file → `ShellExecuteW(NULL, L"open", path, ...)`
- `.lnk` files → resolve shortcut target via `IShellLink`, navigate or open target

**4.10 — Context menu**
- `WM_CONTEXTMENU` on list item:
  - Copy, Cut, Paste, Rename (F2), Delete, Properties (Alt+Enter), Copy Path
  - Make Favourite, Make Flyout Favourite
  - Shell integration: `IContextMenu` invocation for standard shell items

### Checkpoint CP-04

```
CHECKPOINT CP-04 CRITERIA:
[ ] File list loads and displays contents of a folder
[ ] All 4 columns correct (Name, Extension, Date Modified, Size)
[ ] Folder-first sort applied
[ ] Clicking column header sorts; chevron shows direction
[ ] File name color coding correct in dark mode
[ ] Date group dividers appear when sorted by Date Modified
[ ] Divider rows are not selectable
[ ] Multi-select works (click, Shift+Click, Ctrl+Click, Ctrl+A)
[ ] Double-click folder navigates into it
[ ] Double-click file opens with default app
[ ] Left key = up, Right key = enter folder, Enter = open
[ ] Context menu appears with correct items
[ ] Empty folder state renders centered message
[ ] Status bar shows item count and selection summary
[ ] DEVELOPMENT_JOURNAL.md updated
```

---

## Phase 5 — Background Threading & Async Folder Loading

**Goal:** Folder enumeration never blocks the UI thread. Navigation feels instant.

### Stack reminder
- UI thread only: Win32 message pump, all UI calls
- Worker threads: `std::thread` or `_beginthreadex` for enumeration/search
- Communication: `PostMessageW` from worker to UI thread (no shared UI state)
- Cancellation: generation counter (`std::atomic<uint64_t>`)

### Architecture

```
[UI Thread]
  TabManager::navigate(path)
    → increments generation counter
    → posts placeholder "loading" state to ListView
    → starts FolderWorker thread

[FolderWorker Thread]
  captures generation at start
  enumerates with FindFirstFileW/FindNextFileW
  if (generation != captured) → abort, discard results
  PostMessageW(hwndMain, WM_FE_FOLDER_LOADED, generation, results*)

[UI Thread — WM_FE_FOLDER_LOADED]
  if (generation == current) → apply results, update ListView
  else → discard
```

**5.1 — Worker message definitions**
```cpp
// Define in a shared header
#define WM_FE_FOLDER_LOADED     (WM_APP + 1)
#define WM_FE_SEARCH_RESULT     (WM_APP + 2)
#define WM_FE_SEARCH_DONE       (WM_APP + 3)
#define WM_FE_SIBLINGS_READY    (WM_APP + 4)
```

**5.2 — FolderWorker**
- `class FolderWorker`: launched per navigation event
- Constructor takes `path`, `generation`, `hwndTarget`
- Runs `FindFirstFileW` loop, builds `std::vector<FileEntry>` on heap
- Before posting: sorts by default/remembered sort settings
- Posts `WM_FE_FOLDER_LOADED` with `wParam = generation`, `lParam = ptr to vector`
- UI thread takes ownership of pointer, deletes after use

**5.3 — Generation-based stale suppression**
- `MainWindow` holds `std::atomic<uint64_t> folderGeneration_`
- Every new `navigate()` increments it
- Workers check it at start and before posting
- UI handler discards if generation mismatched

**5.4 — Loading state visual**
- When navigation starts: show centered loading spinner (animated via `WM_TIMER` or `SetTimer`)
- Keep previous folder contents visible during load (do not blank the list immediately)
- Replace spinner + contents atomically when results arrive

**5.5 — Incremental refresh**
- On F5 / Refresh: compare new listing to current `displayList_`
  - Added items: insert into correct sorted position
  - Removed items: remove from list
  - Changed items: update in place
- Preserve selection indices where names still match
- Preserve scroll position: `ListView_GetTopIndex`, restore after update

**5.6 — File system change notifications**
- Use `FindFirstChangeNotificationW` or `ReadDirectoryChangesW` on the active folder
- On `FILE_NOTIFY_CHANGE_FILE_NAME | FILE_NOTIFY_CHANGE_DIR_NAME | FILE_NOTIFY_CHANGE_LAST_WRITE`:
  - Debounce: 500 ms timer before triggering incremental refresh
  - Only trigger for the currently visible folder, not all tabs

### Checkpoint CP-05

```
CHECKPOINT CP-05 CRITERIA:
[ ] Navigating to a large folder (1000+ files) does NOT freeze UI
[ ] Loading spinner appears during enumeration
[ ] Previous folder content stays visible until new results arrive
[ ] Rapid tab switching does not crash or show stale data
[ ] Cancellation: navigating away mid-load discards old results cleanly
[ ] F5 / Refresh updates list incrementally without full blank rebuild
[ ] Selection preserved after refresh where possible
[ ] File system changes trigger incremental refresh (add/delete a file, watch list update)
[ ] No WM_FE_* message leaks (all heap pointers freed by UI thread)
[ ] DEVELOPMENT_JOURNAL.md updated
```

---

## Phase 6 — Right Sidebar Shell

**Goal:** The right sidebar is always visible with three sections: Search, Flyout Favourites, Regular Favourites.

### Tasks

**6.1 — Sidebar container**
- Custom window class `L"FE_Sidebar"`
- Default width: 320 px (DPI-scaled), fixed in v1
- Background: slightly different surface than main area (`#212121` vs `#1A1A1A`)
- 1 px left border separator

**6.2 — Sidebar layout**
```
┌───────────────────────┐
│ [Search input]        │  ← fixed height 32 px
├───────────────────────┤
│ Flyout Favourites     │  ← section header, bold 11px
│  [item] ▶             │  ← flyout favourite items
│  [item] ▶             │
│   ...                 │
├───────────────────────┤  ← 1px divider
│ Favourites            │  ← section header
│  [item]               │
│  [item]               │
│   ...                 │
└───────────────────────┘
```

**6.3 — Section headers**
- Bold, 11 px, all-caps or title-case
- Subtle left padding (12 px), top spacing (8 px)
- Section header text: `Flyout Favourites` and `Favourites`

**6.4 — Sidebar items**
- Custom-drawn item rows: 28 px tall, 16 px left-pad, 12 px text
- Hover: soft background `#2A2A2A`
- Active item (currently open folder match): accent underline or stronger highlight
- Right-click menu on items: `Remove`, `Rename`

**6.5 — Sidebar scrolling**
- If items overflow the available height, enable vertical scrolling per section
- Use `WS_VSCROLL` or owner-scroll via `WM_VSCROLL`

### Checkpoint CP-06

```
CHECKPOINT CP-06 CRITERIA:
[ ] Sidebar is visible and correctly positioned (right side, 320px wide)
[ ] Three sections render with headers
[ ] Search section placeholder renders correctly
[ ] Sidebar items render with correct hover states
[ ] 1px divider between Flyout Favourites and Favourites sections
[ ] DEVELOPMENT_JOURNAL.md updated
```

---

## Phase 7 — Search Engine

**Goal:** Recursive wildcard filename search, runs on background thread, results appear incrementally.

### Core algorithm (`WildcardMatch.cpp`)

```cpp
// Case-insensitive wildcard match: * matches any sequence, ? matches one char
bool WildcardMatch(const wchar_t* pattern, const wchar_t* str);
```

Implementation notes:
- Iterative (not recursive) to handle deep paths without stack overflow
- Case-fold both strings to lowercase before comparison
- Handle `*` as zero-or-more chars, `?` as exactly one char
- Must handle paths with `*` at start, middle, and end

### SearchWorker architecture

```cpp
class SearchWorker {
    std::atomic<bool> cancelled_{false};
    std::wstring      rootPath_;
    std::wstring      pattern_;
    HWND              hwndTarget_;
    uint64_t          generation_;
public:
    void run();   // called on worker thread
    void cancel();
};
```

**7.1 — Search trigger**
- User types in search box, presses Enter
- `Ctrl+F` focuses search box
- Previous search worker is cancelled (`.cancel()`)
- New `SearchWorker` started
- File list switches to "search results" mode (flag on `FileListView`)

**7.2 — Recursive enumeration in worker**
- Start with `rootPath_`
- Use stack-based DFS (not recursion) via `std::stack<std::wstring>`
- For each directory: `FindFirstFileW(L"\\*", ...)` → iterate files
- Match file names against pattern using `WildcardMatch`
- Batch results: collect 50 items, then `PostMessageW(WM_FE_SEARCH_RESULT, gen, results*)`
- Check `cancelled_` after every directory

**7.3 — Results mode**
- `FileListView` enters search results mode:
  - Column set changes: Name, Extension, Date Modified, Size, **Path** (full relative path)
  - Or: show full path in Name column as secondary line if column space is tight
- Results appended incrementally to `displayList_`
- Status bar: `126 search results` (updating)

**7.4 — Search done**
- Worker posts `WM_FE_SEARCH_DONE`
- Status bar shows final count: `126 search results`
- Empty results: show empty state "No files found matching …"

**7.5 — Exit search mode**
- Clear search box + Escape → clear results, reload current folder
- Navigate to a folder (click, breadcrumb) → exit search mode automatically

**7.6 — Search mode visuals**
- Search box: active/focused styling, clear button (`×`) when text present
- File list header shows "Search Results in: [root path]" at top as a subtle label

### Checkpoint CP-07

```
CHECKPOINT CP-07 CRITERIA:
[ ] Search runs recursively without blocking UI
[ ] Wildcard * and ? work correctly (test cases in unit tests)
[ ] Results appear incrementally in the list
[ ] Cancellation: starting a new search cancels the previous cleanly
[ ] Path column visible in search results mode
[ ] Status bar updates count during and after search
[ ] Empty result state shown when nothing found
[ ] Escape / clear search box exits search mode
[ ] Navigation exits search mode
[ ] Ctrl+F focuses search field
[ ] Very large trees (10,000+ files) search without UI stall
[ ] DEVELOPMENT_JOURNAL.md updated
```

---

## Phase 8 — Favourites System

**Goal:** Both favourites systems (flyout + regular) are fully functional and persisted in CSV.

### CsvParser (`CsvParser.cpp`)

Must handle:
- Fields with commas inside double quotes: `"path,with,commas","name"`
- Newlines: `\r\n` and `\n`
- Leading/trailing whitespace trimming
- Missing file → empty list (not an error)
- Write: produce minimal valid CSV

### Data model

```cpp
struct FavouriteEntry {
    std::wstring path;
    std::wstring friendlyName;
};

class FavouritesStore {
    std::vector<FavouriteEntry> flyouts_;
    std::vector<FavouriteEntry> regulars_;
    std::wstring flyoutsPath_;   // CSV file path
    std::wstring regularsPath_;  // CSV file path
public:
    void load();
    void save();
    void addFlyout(const std::wstring& path);
    void addRegular(const std::wstring& path);
    void remove(FavouriteType, int index);
    void rename(FavouriteType, int index, const std::wstring& newName);
    const std::vector<FavouriteEntry>& getFlyouts() const;
    const std::vector<FavouriteEntry>& getRegulars() const;
};
```

**8.1 — CSV file locations**
- Store next to `.exe` (portable): `flyout_favourites.csv`, `regular_favourites.csv`
- Fall back to `%APPDATA%\FileExplorer\` if exe directory is read-only

**8.2 — Regular favourites**
- Right-click folder in list → `Make Favourite`
- Added to `FavouritesStore::regulars_`, CSV saved, sidebar refreshed
- Sidebar item click → navigate to folder, or activate matching tab
- Right-click sidebar item: `Remove`, `Rename`
- Display sorted alphabetically by `friendlyName`

**8.3 — Flyout favourites**
- Right-click folder → `Make Flyout Favourite`
- Added to `FavouritesStore::flyouts_`, CSV saved, sidebar refreshed
- Sidebar item click → opens flyout popup (not immediately navigates)
- Flyout popup behavior:
  - First level: explicit entry "Open [folder name]" at top + subfolders listed
  - Background worker enumerates subfolders for each flyout level on demand
  - Clicking a subfolder → navigate immediately, close flyout
  - Clicking "Open [folder name]" at top level → navigate to root, close flyout
  - Subfolders themselves may contain subfolders → nested flyout levels
- Display sorted alphabetically by `friendlyName`

**8.4 — Friendly name editing**
- Right-click sidebar item → `Rename`
- Inline edit or small popup input: `InputBox`-style dialog with `WC_EDIT`
- On confirm: update `FavouritesStore`, save CSV

**8.5 — Sidebar refresh**
- After any `FavouritesStore` mutation, invalidate sidebar for re-render
- Sidebar reads directly from `FavouritesStore` sorted list

### Checkpoint CP-08

```
CHECKPOINT CP-08 CRITERIA:
[ ] Regular favourites: add via context menu, appear in sidebar
[ ] Regular favourite click navigates (or activates existing tab)
[ ] Regular favourite right-click: remove, rename work
[ ] Flyout favourites: add via context menu, appear in sidebar
[ ] Flyout click shows popup with "Open [name]" + subfolders
[ ] Subfolders in flyout are loaded async (not blocking)
[ ] Clicking a subfolder navigates and closes flyout
[ ] Both lists sorted alphabetically by friendly name
[ ] CSV files written and readable after app restart
[ ] CSV survives special chars in paths and names
[ ] Unit tests for CsvParser pass
[ ] DEVELOPMENT_JOURNAL.md updated
```

---

## Phase 9 — File Operations

**Goal:** Copy, cut, paste, delete, rename, new folder, properties, and copy path all work correctly.

### Tasks

**9.1 — Clipboard integration**
- Copy (`Ctrl+C`): put selected file paths on clipboard using `CF_HDROP` format
  - `DROPFILES` struct + null-separated paths + double null terminator
- Cut (`Ctrl+X`): same as Copy plus set `Preferred DropEffect` = `DROPEFFECT_MOVE`
- Paste (`Ctrl+V`): read `CF_HDROP` from clipboard, copy or move files using `SHFileOperationW`
  - `FOF_ALLOWUNDO` flag to support Recycle Bin for deletes
  - `FOF_NOCONFIRMATION` for non-destructive pastes; confirm overwrites

**9.2 — Delete**
- `Delete` key: send to Recycle Bin via `SHFileOperationW` with `FO_DELETE | FOF_ALLOWUNDO`
- `Shift+Delete`: permanent delete with user confirmation dialog (`MessageBoxW`)

**9.3 — Rename (F2 / context menu)**
- `F2` or context menu `Rename` → enter rename mode on selected item
- Inline rename: create an `EDIT` control overlaid on the Name cell of the row
  - Position using `ListView_GetSubItemRect`
  - Pre-fill with current name (without extension if setting says hide ext)
  - `Enter` → `MoveFileW(oldPath, newPath)`; on success, update `displayList_` entry in place
  - `Escape` → cancel, destroy edit control
  - On error: show `MessageBoxW` with error string, revert

**9.4 — New folder**
- Right-click background → `New Folder`
- Generate unique name `New folder`, `New folder (2)`, etc.
- `CreateDirectoryW(newPath, NULL)`
- Insert into `displayList_` in correct sorted position
- Immediately enter rename mode on the new item

**9.5 — Properties**
- `Alt+Enter` or context menu → `ShellExecuteW(NULL, L"properties", path, ...)`
- Or `SHObjectProperties` for shell-integrated properties dialog

**9.6 — Copy path**
- Context menu `Copy path`: put full path string on clipboard as plain text (`CF_UNICODETEXT`)

**9.7 — Shell context menu integration**
- For right-clicked items, optionally invoke `IContextMenu` from shell
  - Query `IShellFolder`, `GetUIObjectOf`, `IContextMenu::QueryContextMenu`
  - Append shell items to custom context menu
  - `IContextMenu::InvokeCommand` for selected item
  - This provides standard shell actions (Open with, Send to, etc.)

**9.8 — Undo (Ctrl+Z)**
- For delete (to Recycle Bin) and copy/move: `SHFileOperationW` with `FOF_ALLOWUNDO` already registers undo in shell
- For rename: maintain a local undo stack with `{ oldPath, newPath }` pairs
- `Ctrl+Z` → `MoveFileW(newPath, oldPath)` and update `displayList_`

### Checkpoint CP-09

```
CHECKPOINT CP-09 CRITERIA:
[ ] Ctrl+C copies selected files to clipboard (CF_HDROP)
[ ] Ctrl+X marks files for cut (CF_HDROP + PreferredDropEffect)
[ ] Ctrl+V pastes (copy or move) into current folder
[ ] Pasted items appear in list immediately
[ ] Delete sends to Recycle Bin
[ ] Shift+Delete permanently deletes after confirmation
[ ] F2 enters inline rename; Enter confirms, Escape cancels
[ ] Rename failure shows error and reverts
[ ] New Folder creates folder, enters rename mode
[ ] Properties dialog opens via Alt+Enter
[ ] Copy path puts full path on clipboard
[ ] Shell context menu integration works for basic shell actions
[ ] DEVELOPMENT_JOURNAL.md updated
```

---

## Phase 10 — Visual Polish & Theme

**Goal:** The app fully looks and feels like a Windows 11 native tool: typography, spacing, colors, icons, states, Mica.

### Tasks

**10.1 — Typography**
- Create `HFONT` instances for all text sizes:
  - `fontBody12` → Segoe UI Variable, 12 px, regular
  - `fontSemibold11` → Segoe UI Variable, 11 px, semibold
  - `fontLarge18` → Segoe UI Variable, 18 px, regular (empty state title)
  - `fontStatus11` → Segoe UI Variable, 11 px, regular
- Recreate fonts on `WM_DPICHANGED`
- Set anti-aliased text rendering via `LOGFONT.lfQuality = CLEARTYPE_QUALITY`

**10.2 — Color palette (dark mode)**

Define all colors in `Colors.h` as `COLORREF`:

```cpp
namespace Colors {
    // Surfaces
    constexpr COLORREF BG_BASE        = 0x001C1C1C;  // deepest bg
    constexpr COLORREF BG_ELEVATED    = 0x00242424;  // nav/tab surface
    constexpr COLORREF BG_SIDEBAR     = 0x00212121;
    constexpr COLORREF BG_STATUS      = 0x001A1A1A;
    constexpr COLORREF BG_ROW_HOVER   = 0x00303030;
    constexpr COLORREF BG_ROW_SEL     = 0x003D3D3D;

    // Text
    constexpr COLORREF TEXT_PRIMARY   = 0x00F0F0F0;
    constexpr COLORREF TEXT_SECONDARY = 0x00A0A0A0;
    constexpr COLORREF TEXT_DISABLED  = 0x00606060;

    // File type colors (dark mode)
    constexpr COLORREF FILE_PDF       = 0x0075_6CE0;  // #E06C75 in BGR
    // (All values stored as Win32 BGR, not RGB)
    ...
}
```

> NOTE: Win32 COLORREF is `0x00BBGGRR` — all hex values must be BGR order.

**10.3 — Icon management**
- Use `SHGetFileInfo` + `ImageList_GetIcon` for shell file type icons (16 px)
- Cache `HICON` per extension in `IconCache` (hash map)
- For navigation buttons: load from embedded resource or draw using Unicode symbols with Segoe MDL2 Assets font
  - Back: `U+E72B`, Forward: `U+E72A`, Up: `U+E74A`, Refresh: `U+E72C`, Search: `U+E721`
  - Pin: `U+E840`, Close: `U+E711`, New: `U+E710`

**10.4 — Mica / DWM backdrop**
```cpp
void TryApplyMica(HWND hwnd) {
    BOOL dark = TRUE;
    DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &dark, sizeof(dark));
    
    DWM_SYSTEMBACKDROP_TYPE backdrop = DWMSBT_MAINWINDOW; // Mica
    HRESULT hr = DwmSetWindowAttribute(hwnd, DWMWA_SYSTEMBACKDROP_TYPE,
                                        &backdrop, sizeof(backdrop));
    if (FAILED(hr)) {
        // Fallback: extend frame into client area with 0-thickness borders
        // and paint with solid BG_BASE
        MARGINS m = {0, 0, 0, 0};
        DwmExtendFrameIntoClientArea(hwnd, &m);
    }
}
```

**10.5 — Loading spinner**
- Simple animated arc drawn in `WM_PAINT` using `Arc`/`Ellipse` GDI calls
- 8-step rotation driven by `WM_TIMER` at 80 ms interval
- Centered in file list area, 24 px diameter
- Stop timer and hide when results arrive

**10.6 — Hover and focus states**
- All custom controls must handle: normal, hover, pressed, focused, selected, disabled
- Use `TrackMouseEvent` for `WM_MOUSELEAVE` detection
- Draw focus rectangle using `DrawFocusRect` or custom 1 px border

**10.7 — Separator rendering**
- 1 px horizontal separators between zones: draw with `MoveToEx` / `LineTo` using `TEXT_DISABLED` color
- No thick borders, no bevels

**10.8 — Empty state graphics**
- Draw large folder outline icon (using Unicode `U+E8B7` from Segoe MDL2 Assets, 48 px)
- "This folder is empty" in `fontLarge18`
- "Drag files here or navigate to a folder" in `fontBody12`, `TEXT_SECONDARY`

### Checkpoint CP-10

```
CHECKPOINT CP-10 CRITERIA:
[ ] All text uses Segoe UI Variable with correct sizes per zone
[ ] Dark surface hierarchy is visually distinct between zones
[ ] File type color coding matches PRD specification
[ ] Navigation button icons render clearly (Segoe MDL2 Assets or equivalent)
[ ] Mica effect applied where supported; solid dark fallback otherwise
[ ] Loading spinner animates correctly and stops cleanly
[ ] All interactive elements show hover/press/focus states
[ ] Separators are 1px, no bevels
[ ] Empty folder state renders correctly
[ ] DPI scaling: all sizes scale correctly at 100%, 125%, 150%
[ ] DEVELOPMENT_JOURNAL.md updated
```

---

## Phase 11 — Settings & Persistence

**Goal:** Per-folder sort, session restore, pane widths, and user settings persist correctly across sessions.

### File locations (portable-first)

All settings files next to `.exe`:
- `settings.ini` — general settings
- `session.ini` — tab session state
- `sort_settings.csv` — per-folder sort preferences
- `flyout_favourites.csv` — flyout favourites
- `regular_favourites.csv` — regular favourites

If exe directory is read-only (Program Files), fall back to `%APPDATA%\FileExplorer\`.

### Tasks

**11.1 — Settings file (`settings.ini`)**

Simple INI-style parser (no external library):
```ini
[General]
ShowHiddenFiles=0
ShowExtensions=1
Theme=dark
SidebarWidth=320

[Window]
Left=100
Top=80
Width=1200
Height=750
```

Parse on startup, write on clean exit.

**11.2 — Per-folder sort settings (`sort_settings.csv`)**

CSV: `path,sort_column,sort_direction`

`SortSettings` class:
```cpp
class SortSettings {
    std::unordered_map<std::wstring, std::pair<SortColumn, SortDirection>> map_;
public:
    void load(const std::wstring& csvPath);
    void save(const std::wstring& csvPath);
    void set(const std::wstring& folderPath, SortColumn col, SortDirection dir);
    std::pair<SortColumn, SortDirection> get(const std::wstring& folderPath) const;
};
```

- Loaded on startup, saved on clean exit or after every sort change
- Limit stored entries: cap at 500 most-recently-used to prevent unbounded growth

**11.3 — Session restore (`session.ini`)**

```ini
[Session]
TabCount=3
ActiveTab=1

[Tab0]
Path=C:\Projects\Alpha
Pinned=0

[Tab1]
Path=D:\Work\Current
Pinned=1

[Tab2]
Path=C:\Users\NT\Documents
Pinned=0
```

- Written on clean exit (`WM_DESTROY` or clean `WM_CLOSE`)
- Read on startup, tabs restored before first window paint
- If session file missing or corrupt: open single default tab (user home)

**11.4 — SessionStore class**

```cpp
class SessionStore {
public:
    struct SessionTab {
        std::wstring path;
        bool pinned;
    };
    
    void save(const std::vector<SessionTab>& tabs, int activeIndex,
              const std::wstring& filePath);
    std::vector<SessionTab> load(const std::wstring& filePath, int& activeIndex);
};
```

**11.5 — Settings dialog / menu**
- Simple menu items under a `Settings` or `View` menu:
  - Show hidden files (toggle)
  - Show file extensions (toggle)
- These update the in-memory settings and trigger folder refresh
- Persisted to `settings.ini` on exit

### Checkpoint CP-11

```
CHECKPOINT CP-11 CRITERIA:
[ ] App reopens with same tabs and folders as last session
[ ] Tab order and pinned state restored
[ ] Per-folder sort preferences restored after navigation and restart
[ ] Show hidden files toggle works and persists
[ ] Show extensions toggle works and persists
[ ] Sidebar width saved and restored if resizable
[ ] Window size/position restored
[ ] All CSV/INI files written to exe directory (or AppData fallback)
[ ] Corrupt/missing settings handled gracefully (no crash, sensible defaults)
[ ] Unit tests: SessionStore round-trip, SortSettings round-trip
[ ] DEVELOPMENT_JOURNAL.md updated
```

---

## Phase 12 — Single-Instance Enforcement

**Goal:** Only one FileExplorer window; second launch opens path in new tab of existing instance.

### Tasks

**12.1 — Named mutex**
- On startup: `CreateMutexW(NULL, TRUE, L"FileExplorerSingleInstanceMutex_v1")`
- If `GetLastError() == ERROR_ALREADY_EXISTS`: second instance is running

**12.2 — IPC via named pipe**
- First instance creates: `CreateNamedPipeW(L"\\\\.\\pipe\\FileExplorerIPC_v1", ...)`
- Reads incoming path strings in a background thread
- Second instance: connects to pipe, writes requested path (or empty string), exits

**12.3 — Focus and tab activation**
- On receiving a path via pipe:
  - `SetForegroundWindow(hwndMain)`
  - Check open tabs: if path already open, activate that tab
  - Else: create new tab for the path

### Checkpoint CP-12

```
CHECKPOINT CP-12 CRITERIA:
[ ] Launching a second instance focuses the first window
[ ] Passing a path argument opens/activates it as a tab
[ ] No second window appears
[ ] Named mutex and pipe created cleanly; no handle leaks
[ ] DEVELOPMENT_JOURNAL.md updated
```

---

## Phase 13 — Unit Tests

**Goal:** All non-UI logic is covered by automated tests. Tests run with a single build step.

### Test project setup
- Google Test, locally sourced
- Test executable: `FileExplorer.Tests.exe`
- All tests run in `Debug x64`
- No UI, no actual file system for pure logic tests (use temp dirs where needed)

### Test files and required coverage

**`test_wildcard.cpp`** — `WildcardMatch`
```cpp
// Required test cases:
TEST(Wildcard, ExactMatch)
TEST(Wildcard, StarMatchesEmpty)
TEST(Wildcard, StarMatchesSequence)
TEST(Wildcard, QuestionMatchesOne)
TEST(Wildcard, QuestionNoMatchEmpty)
TEST(Wildcard, StarAtStart)
TEST(Wildcard, StarAtEnd)
TEST(Wildcard, MultipleStars)
TEST(Wildcard, CaseInsensitive)
TEST(Wildcard, ExtensionPattern)      // "*.pdf" matches "report.pdf"
TEST(Wildcard, DoubleExtension)       // "*.tar.gz" matches "archive.tar.gz"
TEST(Wildcard, NoMatchDifferentExt)   // "*.pdf" does NOT match "report.docx"
TEST(Wildcard, QuestionInMiddle)      // "rep?rt.pdf" matches "report.pdf"
TEST(Wildcard, EmptyPattern)
TEST(Wildcard, EmptyString)
TEST(Wildcard, LongPath)              // 260+ char path
```

**`test_csv.cpp`** — `CsvParser`
```cpp
TEST(Csv, ParseSimpleRow)
TEST(Csv, ParseQuotedFields)
TEST(Csv, ParseCommaInQuotes)
TEST(Csv, ParseNewlineInQuotes)
TEST(Csv, ParseEmptyField)
TEST(Csv, ParseMultipleRows)
TEST(Csv, WriteAndReadRoundtrip)
TEST(Csv, WriteSpecialChars)          // paths with unicode
TEST(Csv, ParseCRLF)
TEST(Csv, ParseLFOnly)
TEST(Csv, MissingFile)                // returns empty list, no throw
TEST(Csv, CorruptLine)               // handles gracefully
```

**`test_sort.cpp`** — Sort logic
```cpp
TEST(Sort, SortByNameAscending)
TEST(Sort, SortByNameDescending)
TEST(Sort, SortByExtension)
TEST(Sort, SortByDate)
TEST(Sort, SortBySize)
TEST(Sort, FolderAlwaysFirst)        // folders before files regardless of sort
TEST(Sort, DateGroupBuckets)         // Today/Yesterday/etc. assigned correctly
TEST(Sort, MixedCaseNames)           // case-insensitive name sort
```

**`test_session.cpp`** — `SessionStore`
```cpp
TEST(Session, SaveAndLoad)
TEST(Session, LoadMissingFile)        // returns empty tabs, no throw
TEST(Session, LoadCorruptFile)        // returns empty/defaults
TEST(Session, ActiveTabPreserved)
TEST(Session, PinnedStatePreserved)
TEST(Session, MaxTabsRoundtrip)       // e.g. 20 tabs
```

**`test_favstore.cpp`** — `FavouritesStore`
```cpp
TEST(Favstore, AddAndLoadFlyout)
TEST(Favstore, AddAndLoadRegular)
TEST(Favstore, RemoveFlyout)
TEST(Favstore, RenameRegular)
TEST(Favstore, AlphabeticOrder)       // sorted after add
TEST(Favstore, DuplicatePathIgnored)  // or merged
TEST(Favstore, SpecialPathChars)
TEST(Favstore, PersistAndReload)      // write CSV, create new instance, verify
```

**`test_pathutils.cpp`** — Path utility helpers
```cpp
TEST(Path, ExtractExtension)
TEST(Path, ExtractFolderName)
TEST(Path, IsRoot)                    // "C:\" is root
TEST(Path, ParentPath)
TEST(Path, NormalizeSeparators)
TEST(Path, LongPath)
```

### Running tests

```bat
msbuild FileExplorer.Tests.vcxproj /p:Configuration=Debug /p:Platform=x64
.\x64\Debug\FileExplorer.Tests.exe
```

Expected output: `[==========] All tests passed.`

### Checkpoint CP-13

```
CHECKPOINT CP-13 CRITERIA:
[ ] All test files compile without errors
[ ] All listed test cases are present and named exactly as specified
[ ] WildcardMatch: all 15 test cases pass
[ ] CsvParser: all 12 test cases pass
[ ] Sort logic: all 8 test cases pass
[ ] SessionStore: all 6 test cases pass
[ ] FavouritesStore: all 8 test cases pass
[ ] PathUtils: all 6 test cases pass
[ ] Test runner exits with code 0
[ ] DEVELOPMENT_JOURNAL.md updated with test results
```

---

## Phase 14 — Integration & Manual Test Checklist

**Goal:** Functional end-to-end verification before release build.

### Manual test matrix

Run each case manually. Document pass/fail in `DEVELOPMENT_JOURNAL.md`.

**Navigation**
- [ ] Open app → restores last session
- [ ] Navigate to a deep folder (5+ levels)
- [ ] Use breadcrumb segments to jump up the tree
- [ ] Use sibling flyout to jump sideways
- [ ] Edit address bar manually, press Enter
- [ ] Press Escape in address bar → breadcrumb restored
- [ ] Back and Forward through navigation history
- [ ] Alt+Up navigates to parent

**Tabs**
- [ ] Open 5 tabs, each in a different folder
- [ ] Switch between tabs → each remembers its own path
- [ ] Close tabs, pin a tab, unpin it
- [ ] Duplicate a tab
- [ ] Context menu: close other tabs, close tabs to the right
- [ ] Ctrl+1 through Ctrl+8 and Ctrl+9

**File list**
- [ ] Open a folder with 500+ files → loads without UI freeze
- [ ] Sort by each column; visual chevron correct
- [ ] Sort persists after navigate-away and navigate-back
- [ ] Color coding visible: PDF red, Excel green, Word blue, archive yellow, text teal
- [ ] Date groups visible when sorted by Date Modified
- [ ] Select with Shift+Click range, Ctrl+Click individual
- [ ] Ctrl+A selects all
- [ ] Click blank space deselects

**Search**
- [ ] Search in a folder with 3000+ files using `*.pdf`
- [ ] Results appear incrementally
- [ ] Start new search while first is running → first is cancelled cleanly
- [ ] Search `report*` → finds all matching files
- [ ] Clear search → returns to folder view
- [ ] Empty result shows empty state message
- [ ] Search results show full path column

**File operations**
- [ ] Copy and paste a file
- [ ] Cut and paste a file (move)
- [ ] Delete a file → appears in Recycle Bin
- [ ] Shift+Delete → permanent (confirmation shown)
- [ ] Rename via F2 → confirm with Enter
- [ ] Rename via F2 → cancel with Escape
- [ ] Create New Folder → enters rename mode immediately
- [ ] Copy path → verify clipboard content

**Favourites**
- [ ] Mark a folder as Regular Favourite → appears in sidebar
- [ ] Click it → navigates or activates tab
- [ ] Mark a folder as Flyout Favourite → appears in sidebar
- [ ] Click it → flyout opens with "Open [name]" + subfolders
- [ ] Rename a favourite in sidebar
- [ ] Remove a favourite
- [ ] Restart app → favourites preserved

**Session & Settings**
- [ ] Close and reopen → tabs restored
- [ ] Toggle Show Hidden Files → hidden files appear/disappear
- [ ] Toggle Show Extensions
- [ ] Per-folder sort: set sort in folder A, navigate away, come back → sort preserved

**Performance**
- [ ] Navigate to folder with 5000+ files: list appears in < 2 sec
- [ ] Sort 5000-item list: < 100 ms, no UI freeze
- [ ] Rapid tab switching (5 tabs, clicking quickly) → no crashes
- [ ] Search 50,000 files: results appear incrementally, no UI freeze

**DPI and resize**
- [ ] At 125% DPI: all zones scale correctly
- [ ] At 150% DPI: all zones scale correctly
- [ ] Resize window to minimum: no layout corruption
- [ ] Maximize: all zones fill correctly

### Checkpoint CP-14

```
CHECKPOINT CP-14 CRITERIA:
[ ] All manual test categories completed
[ ] Pass rate ≥ 95% (any failures documented in journal with fix plan)
[ ] No crashes observed during entire manual test run
[ ] No UI freezes > 200ms observed
[ ] All keyboard shortcuts listed in PRD section 20 verified working
[ ] DEVELOPMENT_JOURNAL.md: full manual test results recorded
```

---

## Phase 15 — Portable Release Build

**Goal:** A portable, self-contained build that runs on any Windows 11 x64 machine from a folder.

### Tasks

**15.1 — Release build configuration**
- Configuration: `Release`, `x64`
- Enable optimizations: `/O2`
- Disable debug info in binary: `/DEBUG:NONE` or separate `.pdb`
- Static CRT: `/MT` (link CRT statically — no `vcredist` dependency)
  - Or: `/MD` with documented CRT dependency
  - Prefer `/MT` for true portability

**15.2 — Dependency verification**
```bat
dumpbin /dependents FileExplorer.exe
```
Expected: only Windows system DLLs (`kernel32.dll`, `user32.dll`, `shell32.dll`, `comctl32.dll`, `dwmapi.dll`, `uxtheme.dll`, `shlwapi.dll`)

**15.3 — Portable folder**
```
FileExplorer_v1.0_portable/
├── FileExplorer.exe
├── FileExplorer.exe.manifest (embedded)
└── README.txt
```

Data files are created on first run:
- `settings.ini`
- `session.ini`
- `flyout_favourites.csv`
- `regular_favourites.csv`
- `sort_settings.csv`

**15.4 — Smoke test on clean machine**
- Copy portable folder to a machine without Visual Studio
- Launch `FileExplorer.exe`
- Navigate, search, add a favourite → verify all work

**15.5 — Version resource**
- Add `VERSIONINFO` resource:
  - ProductName: `FileExplorer`
  - FileVersion: `1.0.0.0`
  - CompanyName: (owner)

### Checkpoint CP-15

```
CHECKPOINT CP-15 CRITERIA:
[ ] Release build compiles without errors or warnings
[ ] .exe runs on clean Windows 11 machine (no VS runtime installed)
[ ] dumpbin shows no non-system DLL dependencies
[ ] Portable folder contains only .exe (all data created on first run)
[ ] Version resource embedded
[ ] Size of .exe is reasonable (< 5 MB)
[ ] DEVELOPMENT_JOURNAL.md: final release notes
```

---

## Cross-Cutting Rules (Enforced Every Phase)

These rules apply to every line of code, every session, every phase. An AI assistant must check these before marking any checkpoint complete.

### Memory safety
- No raw `new` / `delete` for resource ownership — use `std::unique_ptr` or custom RAII wrappers
- Every `HANDLE` returned by Win32 must be wrapped: `std::unique_ptr<void, HandleDeleter>`
- Every `HGLOBAL`, `HICON`, `HFONT` must have a matching free/destroy
- No buffer overflows: use `StringCbCopyW` / `StringCchCopyW` or `std::wstring` — never `wcscpy`

### Threading safety
- UI controls are only ever accessed from the UI thread
- Worker threads only communicate via `PostMessageW`
- Shared data accessed from multiple threads must use `std::mutex` or `std::atomic`
- No `SendMessageW` from worker threads to UI (can deadlock)

### Error handling
- All Win32 API calls that can fail must check return values
- On failure: log to `OutputDebugStringW` at minimum; surface to user if actionable
- No silent failures in file operations

### Code style
- All source files: UTF-8 with BOM
- All string literals involving paths or UI: `L"..."` prefix
- No `#define` for constants that can be `constexpr`
- No `using namespace std;` in header files

---

## AI Assistant Behavioral Instructions

The following instructions govern how an AI assistant must work on this project.

### At the start of every session:
1. Read `DEVELOPMENT_JOURNAL.md` — understand current state
2. Read this plan — re-read the phase you are working on
3. Re-read the **MANDATORY STACK REMINDER** at the top
4. Identify which checkpoint is the current target
5. Do not skip phases or checkpoints

### During work:
- Write code → compile → run (if applicable) → fix errors → repeat
- Do not commit code with compile errors
- Do not leave unimplemented stubs (use `static_assert(false)` or `TODO` comments that are surfaced, not silent)
- After completing any task, run the relevant unit tests
- Every architectural decision → document in journal

### At the end of every session:
1. Run all unit tests
2. Verify the current checkpoint criteria
3. Update `DEVELOPMENT_JOURNAL.md`:
   - What was completed
   - What failed and why
   - Any decisions made
   - Next action item

### If the plan conflicts with itself or with the PRD:
- PRD wins over this plan on product behavior
- This plan wins over the PRD on technical implementation constraints
- Document the conflict in the journal and proceed with the higher-authority rule

### Never:
- Introduce external libraries beyond Google Test (no vcpkg, no NuGet for runtime code)
- Use managed C++/CLI
- Use any UI framework other than raw Win32
- Skip the journal update at end of session

---

*End of IMPLEMENTATION_PLAN.md*  
*Companion: DEVELOPMENT_JOURNAL.md*