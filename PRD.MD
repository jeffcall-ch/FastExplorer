# FileExplorer — Native Win32 C++ Product & Technical Requirements Document

**Version:** 2.5  
**Date:** April 2026  
**Status:** Engineering ready  
**Audience:** Product owner, designer, AI coding assistant, developer  
**Target platform:** Windows 11 x64  
**Target stack:** Native Win32 desktop application in C++20 using MSVC  
**Deployment target:** Portable `.exe` or portable folder, no installer required

---

## 1. Purpose of this document

This document defines the product, interaction, visual design, and technical constraints for **FileExplorer**.

The application is intentionally focused on:

- lightning-fast folder navigation
- tabbed browsing
- very fast recursive filename search
- strong visual clarity
- remembered per-folder view behavior
- efficient favourites and breadcrumb navigation
- portable deployment

This document includes both:

- what the user sees and how it behaves
- technical constraints that the implementation must obey

The application does **not** perform file-content extraction or metadata parsing.

---

## 2. Product overview

FileExplorer is a **Windows desktop file manager** for users who work with many folders, project documents, PDFs, Office files, archives, and structured engineering files.

Its purpose is to be:

- faster than Windows Explorer in daily folder navigation
- easier to scan visually
- better at handling tabs, favourites, and search
- more stable and predictable in large project folders
- highly responsive in deep folder trees
- portable and lightweight

The product intentionally avoids expensive background parsing and focuses on:

- very fast folder switching
- very fast recursive filename search
- strong visual clarity
- remembered per-folder sort behavior
- efficient favourites and breadcrumb navigation

---

## 3. Primary users and key use cases

Primary user examples:

- project engineer
- document controller
- design coordinator
- office-heavy technical user
- engineer working in deep nested project folder structures

Core use cases:

- switch quickly between several project folders
- jump sideways between sibling folders without going up and down the tree
- search recursively in the current subtree using wildcards
- keep favourite folders permanently available
- drill down quickly from a top-level favourite through subfolders and files
- work with very large file sets without the UI lagging
- preserve sort and layout behavior per folder

---

## 4. Product principles

The product must follow these principles:

1. **Feels native to Windows 11**
2. **Looks modern, compact, and professional**
3. **Never blocks the UI**
4. **Shows structure clearly**
5. **Prioritizes speed over visual gimmicks**
6. **Preserves context when changing folders or tabs**
7. **Makes large lists scannable**
8. **Avoids unnecessary blank states and flicker**
9. **Stays portable and lightweight**
10. **Keeps all heavy work out of the UI thread**

---

## 5. Implementation stack and technical constraints

The application shall be built as:

- **native Win32 desktop application**
- **C++20**
- **MSVC**
- **Visual Studio**
- **x64 only**
- **single process**

The following are not allowed in v1:

- C#
- WPF
- WinUI 3
- Windows App SDK
- WebView2
- Electron
- Qt
- SQLite
- external database/indexing service
- second background engine process
- IPC architecture
- file-content parsing libraries for metadata extraction

All runtime caching shall be **in memory only**.

Persistent non-cache settings may be written to lightweight files.

---

## 6. Application shell and window layout

## 6.1 General shell

The application opens as a standard **resizable Windows 11 desktop window** with native window behavior.

The visual target is:

- Windows 11 look and feel
- Fluent-style iconography
- native-feeling window chrome
- tab strip visually merged into the title region
- dark theme by default
- Mica-like or Mica-backed look where achievable with plain Win32 and DWM, with graceful fallback to a solid dark surface if unsupported or undesirable

The title text shall show:

`<current folder name> - FileExplorer`

Example:

`Current Job - FileExplorer`

## 6.2 Overall zone structure

From top to bottom, the application is organized like this:

```text
┌────────────────────────────────────────────────────────────────────────────┐
│ [Pinned Tab] [Tab 1] [Tab 2] [+]                               [— □ ×]   │
├────────────────────────────────────────────────────────────────────────────┤
│ ◄ ► ▲ ⟳   [Breadcrumb / address bar]                                      │
├───────────────────────────────────────────────────────┬────────────────────┤
│                                                       │ Search             │
│                                                       ├────────────────────┤
│                                                       │ Flyout favourites  │
│ Main details file list                                │                    │
│                                                       ├────────────────────┤
│                                                       │ Favourites         │
│                                                       │                    │
├───────────────────────────────────────────────────────┴────────────────────┤
│ 24 items · 3 selected (1.2 MB)                         5.4 GB free on C:   │
└────────────────────────────────────────────────────────────────────────────┘
```

## 6.3 Pane structure

There is:

- **no left tree view**
- **main details pane** occupying the large left/center area
- **right inspector sidebar** always visible
- **status bar** at the bottom

The right sidebar is a focused utility pane for search and favourites.

## 6.4 Default dimensions

Default target dimensions and heights:

- right sidebar default width: **320 px**
- status bar height: **24 px**
- tab strip target height: **34–36 px**
- navigation bar target height: **36–40 px**
- search bar height: **30–32 px**
- compact file row height target: **24–28 px**
- file icon size in list: **16 px**

The sidebar may be resizable later, but in v1 it may remain fixed if needed for simplicity. If resizable, its width must be remembered between sessions.

---

## 7. Visual design specification

## 7.1 Visual language

The visual language shall stay close to **Windows 11**:

- rounded corners where natural for custom controls
- Fluent-style icons
- subtle separators
- layered surfaces
- restrained contrast
- no heavy ornamentation
- modern compact density
- clean typography
- soft hover states
- clear active and selected states

The app must feel like a **serious engineering productivity tool**, not a consumer-style file browser.

## 7.2 Theme behavior

### Default theme

- **Dark theme by default**

### Optional later

- light theme support

If light theme is implemented, it should retain the same layout and visual hierarchy.

## 7.3 Typography

Default typography shall follow Windows 11 conventions as closely as possible.

Preferred font family:

- **Segoe UI Variable**
- fallback: **Segoe UI**

Recommended visual sizes:

- tab text: **12 px**
- breadcrumb text: **12 px**
- list row text: **12 px**
- sidebar item text: **12 px**
- sidebar section headers: **11–12 px semibold**
- status bar text: **11 px**
- date group divider text: **11 px**
- empty-state title text: **18–20 px**
- empty-state subtitle text: **12–13 px**

Font rendering must be crisp and readable at 100% scaling and remain stable under Windows DPI scaling.

## 7.4 Surfaces

Visual surface hierarchy:

1. title/tab surface
2. navigation surface
3. main content surface
4. sidebar surface
5. status surface

These should be subtly separated with tone, transparency, or 1 px separators, not heavy borders.

## 7.5 Borders and separators

Use subtle separators:

- 1 px dividers between major zones
- no thick borders
- no beveled legacy borders
- horizontal divider between flyout favourites and regular favourites in sidebar
- divider rows in grouped file lists

## 7.6 Icons

Use Fluent-style, Windows-native-feeling icons.

Minimum icon usage:

- tab pin icon
- close button icon
- back / forward / up / refresh
- search
- chevrons in breadcrumb
- favourite / pin / flyout markers
- new tab
- context actions where custom icons are used

File icons in list view:

- 16 px
- aligned left in Name column
- shell or appropriate file-type icons where practical

## 7.7 Colors

The PRD preserves the original visual scanning logic.

### General UI colors

Use Windows 11 dark theme semantics:

- dark neutral background
- slightly lighter elevated surfaces
- clear selected row background
- subdued inactive text
- bright readable primary text
- soft hover background
- accent used sparingly

Exact hex values may be finalized in design implementation, but the relative hierarchy must remain stable.

### File-name color coding in dark mode

In the **Name** column, file names are colored by file type:

- PDF → **red**
- Excel → **green**
- Word → **blue**
- archives (`.zip`, `.7z`, etc.) → **yellow**
- text / markdown → **teal**
- everything else → default text color

### Light mode rule

If light mode is implemented:

- only **PDF** remains color-coded
- all other file names use default text color

These colors must aid scanning without becoming visually noisy.

## 7.8 Hover and selection visuals

Interactive elements must show clear but restrained state changes.

Required state types:

- normal
- hover
- pressed
- focused
- selected
- disabled
- active tab
- inactive tab

The selected list row must be unmistakable in both active and inactive window states.

## 7.9 Empty and loading visuals

### Empty folder state

When a folder is empty, show:

- large centered icon
- short friendly message
- secondary text if needed

### Loading state

When a folder is loading:

- show a centered circular progress indicator
- keep previous content visible until the new contents are ready enough
- do not flash a blank screen

---

## 8. Tab strip

## 8.1 Tab strip position

The tab strip is at the very top of the window and visually merges into the title area.

The drag region for moving the window sits around and between tab elements where appropriate.

## 8.2 Tab appearance

Tabs behave like a browser tab strip.

Each tab shows:

- folder name as the tab label
- full path as tooltip on hover
- close button unless pinned

Pinned tabs are:

- narrower
- show pin icon instead of folder name or a very compact label
- not closable via the normal close button

Active tab visual:

- brighter text
- stronger foreground contrast
- slightly elevated or highlighted surface

Inactive tabs:

- lower contrast
- still legible
- hover highlight on mouse over

## 8.3 New tab button

There is a visible **+** button in the tab strip.

It creates a new tab.

## 8.4 Tab context menu

Right-clicking a tab shows:

- Close tab
- Close other tabs
- Close tabs to the right
- Duplicate tab
- Pin / Unpin tab

## 8.5 Minimum tab rules

- the last remaining tab cannot be closed
- pinned tabs cannot be closed through normal close actions
- when opening a folder from favourites or navigation, the app checks whether that folder already exists in an open tab and activates it instead of opening a duplicate where appropriate

## 8.6 Tab keyboard shortcuts

Required:

- Ctrl+T → new tab
- Ctrl+W → close current tab
- Ctrl+Tab → next tab
- Ctrl+Shift+Tab → previous tab
- Ctrl+1 … Ctrl+8 → jump to tab 1–8
- Ctrl+9 → jump to last tab

---

## 9. Navigation bar

## 9.1 Position

The navigation bar sits below the tab strip.

## 9.2 Left navigation button group

Buttons on the left:

- Back
- Forward
- Up
- Refresh

Visual rules:

- compact icon buttons
- greyed out when unavailable
- hover and pressed states
- no oversized toolbar look

Actions:

- Back → go to previous folder in current tab history
- Forward → go to next folder in current tab history
- Up → go to parent folder
- Refresh → refresh current folder incrementally

## 9.3 Address bar / breadcrumb

The center section is the current path control.

### Normal state

The path is shown as clickable breadcrumb segments:

`C: › Projects › Current Job › Site Drawings`

Behavior:

- each segment is clickable
- clicking a segment navigates directly to that folder
- the current/rightmost segment is always visible
- long paths scroll horizontally if needed
- there is a trailing navigation affordance at the end

### Sibling flyout affordance

Next to each breadcrumb level there is a **small compact rectangular button / chevron trigger**.

Clicking it opens a flyout containing **sibling directories at that level**.

Example:  
If current path is:

`D:\Projects\Plant\Docs`

then clicking the sibling trigger next to `Plant` shows sibling folders located under:

`D:\Projects\`

This gives the user a fast sideways jump in the tree.

The visual treatment should stay close to the original chevron-dropdown idea, but may be rendered as a small rectangular flyout trigger.

### Edit state

Clicking in the address bar switches to an editable full-path text field.

Behavior:

- the full current path is shown as text
- the full path is automatically selected immediately
- the user can copy it immediately without manually selecting the whole line
- typing replaces the selected path
- typing shows autocomplete suggestions from the real file system
- Enter navigates
- Escape cancels and returns to breadcrumb mode
- losing focus returns to breadcrumb mode

The control must clearly distinguish:

- breadcrumb display mode
- path edit mode

---

## 10. Right inspector sidebar

The right sidebar is always visible and remains a fixed, narrow utility panel.

Default width:

- **320 px**

From top to bottom:

1. Search
2. Flyout favourites
3. Regular favourites

A 1 px divider separates the two favourites sections.

## 10.1 Search section

At the top is a search field labelled or placeholder-texted:

`Search...`

Visual requirements:

- fixed-height modern input field
- search icon on the left if appropriate
- clean Windows 11 style
- clear focus ring / active state
- aligned with sidebar padding and section spacing

## 10.2 Flyout favourites section

This is the upper favourites/navigation section.

Purpose:

- fast drill-down from top-level anchor folders

Behavior:

- user clicks a flyout favourite to start the flyout
- the top-level click opens the flyout, not the folder immediately
- inside the first flyout level, there must be an explicit entry to open the top-level favourite folder itself
- once the user is inside deeper flyout levels, clicking any folder or file opens it immediately
- when a deeper flyout entry is clicked, the full flyout closes immediately to clear the screen
- folders may themselves fly out further
- folders can be added to this section via a dedicated action such as:
  - `Make Flyout Favourite`

Data storage:

- flyout favourites are stored in a CSV file
- required columns:
  - `path`
  - `friendly_name`
- default friendly name is the folder name when the favourite is created
- friendly name can be changed through right-click
- the application keeps this CSV updated automatically
- this file is part of the portability concept

Sorting:

- flyout favourites are sorted alphabetically in the UI by friendly name

Items should look like navigation targets, not like ordinary file list rows.

## 10.3 Regular favourites section

At the bottom is the standard favourites list.

Purpose:

- one-click direct pinned folders

Behavior:

- right-click a folder → `Make Favourite`
- the item appears in the regular favourites list
- clicking it opens that folder or activates its existing tab if already open

Data storage:

- regular favourites are stored in a separate CSV file
- required columns:
  - `path`
  - `friendly_name`
- default friendly name is the folder name when the favourite is created
- friendly name can be changed through right-click
- the application keeps this CSV updated automatically
- this file is part of the portability concept

Sorting:

- regular favourites are sorted alphabetically in the UI by friendly name

Section header:

- **Favourites**

A bold header and subtle divider should visually separate it from the flyout favourites section.

## 10.4 Sidebar visuals

Sidebar must feel calm and structured:

- consistent padding
- section headers
- subtle separators
- no noisy frames
- clearly clickable navigation items

Pane width shall be remembered between sessions if resizable.

---

## 11. Main file list

## 11.1 General behavior

The main central area displays the current folder contents in a **details/columns view**.

This is the default and primary view.

No icon grid view is required in v1.

## 11.2 Default columns

The full intended column set for v1 is:

- **Name**
- **Extension**
- **Date modified**
- **Size**

Notes:

- `Name` shows file or folder name
- `Extension` shows file extension where applicable
- `Date modified` shows date and time
- `Size` is blank for folders unless folder size is later added

## 11.3 Column visuals and behavior

Columns must support:

- drag-to-resize width
- click-to-sort
- visual sort chevron in header
- clear hover state on headers
- consistent alignment and spacing

Sort preferences are remembered **per folder** and restored:

- when navigating back
- when reopening the app

This persistence is stored in a lightweight settings file, not in a database.

## 11.4 Row visuals

Rows are compact.

Requirements:

- short row height for dense scanning
- 16 px icon in Name column
- clear alternating or uniform row treatment depending on theme
- selected row strongly visible
- multi-select clearly visible
- keyboard focus state visible

## 11.5 File-name color coding

In dark mode, the **Name** text is color coded by file type:

- PDF → red
- Excel → green
- Word → blue
- archives → yellow
- text / markdown → teal
- all other types → normal text color

Folders keep the standard folder color/text treatment.

## 11.6 Hidden and system files

Hidden and system files are shown or hidden according to settings.

## 11.7 File extensions display

Extensions can be shown or hidden via settings, but the default is:

- **Show file extensions = On**

The separate `Extension` column remains supported.

## 11.8 Date group dividers

When sorted by **Date modified**, files may be grouped with divider rows such as:

- Today
- Yesterday
- This week
- Last month

Divider rows are:

- not selectable
- visually distinct
- smaller text
- subtle color

## 11.9 Selection behavior

The list must feel familiar to Windows users.

Required selection interactions:

- click → select one item
- Ctrl+Click → add/remove one item
- Shift+Click → select range
- Ctrl+A → select all
- click blank space → deselect
- arrow keys → navigate at all times
- Up / Down → move row highlight
- Left → go one folder up
- Right → step into selected folder
- Enter → open selected item; if it is a folder, step into it
- Shift + arrows → extend selection
- mouse drag/fence selection → supported

Left and Right navigation must work consistently as direct navigation shortcuts, not only as text-caret behavior.

## 11.10 Opening items

- double-click folder → navigate into folder in current tab
- double-click file → open with default application
- shortcut (`.lnk`) files → resolve and open/navigate to target where appropriate

## 11.11 Empty folder state

When a folder is empty:

- centered empty-state icon
- centered short message
- visually calm, not alarming

## 11.12 Loading behavior

When loading a new folder:

- previous folder contents stay visible until replacement is ready enough
- centered loading spinner may appear
- no blank flash
- no heavy blocking overlay

## 11.13 Row count

Bottom-left status must show the number of rows currently displayed.

Examples:

- `12 items`
- `126 search results`
- `24 items · 3 selected (1.2 MB)`

---

## 12. Search

## 12.1 Search location

The search field is located at the **top of the right inspector sidebar**.

## 12.2 Search scope

Search covers:

- filenames
- optionally full relative paths in search matching later if useful

It does **not** search file contents in v1.

Search is recursive under the current folder.

## 12.3 Search behavior

1. user types query
2. presses Enter
3. search runs recursively in the current folder subtree
4. file list switches to search results mode
5. results appear incrementally
6. search is case-insensitive
7. wildcards such as `*` and `?` are supported
8. editing or starting another search cancels the previous search
9. search must remain extremely fast

## 12.4 Search results mode visuals

In search mode, the main list switches from folder listing to a flat results list.

Full path visibility is required for results.

This can be implemented by:

- adding a `Path` column in search mode
or
- showing full path in a dedicated visible column arrangement

The result list must still look like the same main file list, not like a different tool window.

## 12.5 Exiting search mode

Clearing the search box and confirming, or pressing Escape, returns to the current folder listing.

Navigating to a folder also clears search mode.

---

## 13. Favourites

There are two separate favourites systems.

## 13.1 Flyout favourites

Purpose:

- structured drill-down navigation from top-level anchors

Behavior:

- folders can be marked as flyout favourites
- one click starts the flyout
- the top-level flyout includes an explicit entry to open the top-level favourite folder itself
- deeper flyout clicks open the selected folder or file immediately
- after opening from a deeper flyout level, the flyout closes immediately
- flyout favourites are displayed alphabetically by friendly name

Storage:

- stored in a CSV file with columns:
  - `path`
  - `friendly_name`
- default friendly name is the folder name
- friendly name is editable by right-click
- CSV is maintained automatically by the application

## 13.2 Regular favourites

Purpose:

- one-click direct pinned folders

Behavior:

- right-click folder → `Make Favourite`
- appears in regular favourites section
- one click opens folder or activates matching existing tab
- regular favourites are displayed alphabetically by friendly name

Storage:

- stored in a separate CSV file with columns:
  - `path`
  - `friendly_name`
- default friendly name is the folder name
- friendly name is editable by right-click
- CSV is maintained automatically by the application

## 13.3 Visual behavior

Both favourites systems should be visually distinct:

- flyout favourites feel exploratory and hierarchical
- regular favourites feel static and direct

The right sidebar must make this distinction obvious.

---

## 14. Refresh behavior

Refresh must be incremental and row-aware wherever possible.

Refresh is triggered by:

- pressing Refresh button
- pressing F5
- changing tab
- returning to a previously open tab
- app regaining focus after switching away

Behavioral goal:

- update changed rows
- add/remove changed items
- avoid full blank rebuild
- preserve selection where practical
- preserve scroll position where practical

---

## 15. File operations and context actions

## 15.1 Copy / cut / paste

Supported:

- Ctrl+C
- Ctrl+X
- Ctrl+V

Visual behavior goal:

- pasted items appear immediately in destination view
- moved items disappear immediately from source view as soon as the operation is logically committed

## 15.2 Delete

Supported:

- Delete → Recycle Bin
- Shift+Delete → permanent delete

## 15.3 Rename

Supported:

- F2
- context menu Rename

Behavior:

- inline rename in Name column
- Enter confirms
- Escape cancels
- row repositions immediately if sort order changes
- failure reverts and shows error

## 15.4 New Folder

Right-click background or appropriate context must support:

- New Folder

New folder appears immediately in correct sorted position and enters rename mode if possible.

## 15.5 Properties

Supported:

- Alt+Enter
- context menu → Properties

Opens standard Windows properties dialog.

## 15.6 Copy path

Context menu must support:

- Copy path

This copies the full path of the selected item.

## 15.7 Shell-integrated actions

Where practical, right-click may include standard shell actions.

---

## 16. Status bar

The status bar runs along the bottom of the window.

Height:

- **24 px**

Left side shows:

- item count
- selection summary
- optional selected total size

Examples:

- `24 items`
- `24 items · 3 selected`
- `24 items · 3 selected (1.2 MB)`

Right side shows:

- free space on current drive

Example:

- `5.4 GB free on C:`

Updates immediately after:

- selection change
- folder refresh
- file operation
- search result updates

---

## 17. Settings

Settings may be exposed through menus and shortcuts.

Required settings:

- Show hidden files → default Off
- Show file extensions → default On
- dark theme default
- optional light theme later

Remembered settings:

- pane widths
- per-folder sort mode and direction
- window size/position optionally

Favourites are not stored in the generic settings store. They are stored in their dedicated CSV files.

---

## 18. Session restore

On reopen, the application restores the previous session:

- open tabs
- folders shown in each tab
- tab order
- pinned tabs if supported in restore
- active tab if practical

This should feel like reopening a browser session.

---

## 19. Single-instance behavior

If a second instance is launched:

- existing window is focused
- requested folder opens in a new tab or activates existing matching tab
- second full window is not opened

---

## 20. Keyboard shortcuts

Required keyboard behavior:

- Ctrl+T → new tab
- Ctrl+W → close tab
- Ctrl+Tab → next tab
- Ctrl+Shift+Tab → previous tab
- Ctrl+1 … Ctrl+8 → jump to tab
- Ctrl+9 → jump to last tab
- Alt+Left → back
- Alt+Right → forward
- Alt+Up → parent folder
- F5 → refresh
- Alt+D or Ctrl+L → focus address bar
- Ctrl+F → focus search
- F2 → rename
- Ctrl+C → copy
- Ctrl+X → cut
- Ctrl+V → paste
- Delete → delete
- Shift+Delete → permanent delete
- Ctrl+Z → undo if supported for implemented operations
- Ctrl+A → select all
- Alt+Enter → properties
- Up / Down → row navigation
- Left → one folder up
- Right → enter selected folder
- Enter → open selected item

All regular file explorer-style keyboard interactions should feel natural.

---

## 21. Runtime caching and persistence

## 21.1 Runtime caches

All runtime caching is **memory only**.

Allowed in-memory caches include:

- folder listings
- tab state
- search result state
- sibling folder lists
- flyout favourite expansions

## 21.2 Persistent settings

Persist only lightweight settings in files, such as:

- per-folder sort settings
- pane widths
- session restore data
- optional theme preference

Use dedicated CSV files for favourites:

- one CSV for flyout favourites
- one CSV for regular favourites

No SQLite. No database. No indexing service.

---

## 22. Performance and threading rules

These are hard constraints.

1. UI thread must never block on directory enumeration
2. UI thread must never block on recursive search
3. directory loads, search, and sibling flyouts must be asynchronous
4. stale work must be cancelled or ignored
5. updates must be batched
6. previous content should stay visible until replacement is ready enough
7. no blank-screen flash during navigation
8. if a value cannot be fetched quickly, show partial UI and continue
9. the application must stay responsive while navigating, searching, sorting, and refreshing
10. v1 must avoid expensive work that harms perceived speed

---

## 23. Technical design guidance

The app shall be implemented as:

- one UI thread
- background worker threads for folder enumeration and search
- generation-based cancellation or equivalent stale-result suppression
- wide-character Win32 APIs
- native list control or owner-data/virtualized list approach where useful
- RAII-based C++ ownership
- minimal dependencies
- portable release mindset from the beginning

The application shall not parse file contents in v1.

---

## 24. Out of scope for v1

Not part of current v1:

- inline preview pane
- full content search
- document metadata extraction
- project or revision columns
- editing metadata back into files
- plugin system
- FTP/cloud storage as first-class location
- bulk rename suite
- file tagging/rating
- second background engine process
- database-backed index

---

## 25. Acceptance summary

The product is correct when it delivers all of the following together:

- native Windows 11-feeling shell
- tabbed file browsing
- breadcrumb navigation with sibling flyouts
- compact dense details list
- restored file-type colors and list styling
- right sidebar with search + flyout favourites + regular favourites
- per-folder sort memory
- session restore
- familiar Explorer-like keyboard and mouse behavior
- extremely fast recursive wildcard search
- alphabetical favourites in the UI
- favourites persisted in CSV files
- no UI stalls
- portable deployment and lightweight architecture
