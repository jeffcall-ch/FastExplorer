# AGENTS.md — FileExplorer

> **Read this file first. Every session. No exceptions.**  
> Then read `DEVELOPMENT_JOURNAL.md`. Then read the phase you are working on in `IMPLEMENTATION_PLAN.md`.  
> Only then write code.

---

## What this project is

**FileExplorer** is a native Windows 11 file manager built as a portable `.exe`.  
It replaces Windows Explorer for power users who need fast tab-based folder navigation, recursive search, and favourites.

PRD: `FileExplorer_PRD_v2_5.md`  
Plan: `IMPLEMENTATION_PLAN.md`  
Journal: `DEVELOPMENT_JOURNAL.md`

---

## Development environment

| Item | Value |
|---|---|
| OS | Windows 11 x64 (VM) |
| IDE / editor | **VS Code** with AI coding assistant (writing code) |
| Build / compile | **Visual Studio 2026** (MSBuild, MSVC toolset) |
| Source control | **GitHub** |
| Build trigger from VS Code | `.vscode/tasks.json` calling MSBuild (see below) |
| Debugger | Visual Studio 2026 debugger (attach or launch from VS) |

### Workflow split

VS Code and Visual Studio 2026 are used for different things. Do not confuse them.

- **VS Code** — where you write and edit source files, run tasks, use the AI assistant
- **Visual Studio 2026** — where you compile (`Ctrl+Shift+B` in VS or via task), debug, and manage `.vcxproj` / `.sln` settings

Source files edited in VS Code are the same files compiled by Visual Studio 2026. There is no sync issue — they share the same filesystem.

### VS Code build task

A `.vscode/tasks.json` must exist at the repo root to allow building from VS Code without switching windows:

```json
{
  "version": "2.0.0",
  "tasks": [
    {
      "label": "Build Debug x64",
      "type": "shell",
      "command": "msbuild",
      "args": [
        "FileExplorer.sln",
        "/p:Configuration=Debug",
        "/p:Platform=x64",
        "/m"
      ],
      "group": { "kind": "build", "isDefault": true },
      "problemMatcher": "$msCompile"
    },
    {
      "label": "Build Release x64",
      "type": "shell",
      "command": "msbuild",
      "args": [
        "FileExplorer.sln",
        "/p:Configuration=Release",
        "/p:Platform=x64",
        "/m"
      ],
      "group": "build",
      "problemMatcher": "$msCompile"
    },
    {
      "label": "Run Tests",
      "type": "shell",
      "command": ".\\x64\\Debug\\FileExplorer.Tests.exe",
      "group": "test",
      "dependsOn": "Build Debug x64"
    }
  ]
}
```

`msbuild` must be on the PATH. If it isn't, use the full path from the Visual Studio 2026 installation, typically:  
`C:\Program Files\Microsoft Visual Studio\2026\Community\MSBuild\Current\Bin\MSBuild.exe`

### .gitignore requirements

The repo must have a `.gitignore` that excludes build artifacts before the first commit:

```
.vs/
x64/
Debug/
Release/
*.user
*.suo
*.VC.db
*.opendb
ipch/
```

Never commit compiled binaries, `.pdb` files, or intermediate build objects.

### VM note on Mica

The development VM may not expose GPU hardware acceleration to the compositor. `DwmSetWindowAttribute` calls for Mica (`DWMSBT_MAINWINDOW`) will likely return failure silently on the VM. This is expected — the fallback to solid dark `#1C1C1C` will activate. Do not treat this as a bug during VM development. Verify Mica on bare-metal Windows 11 hardware before release.

---

## The stack. This is not negotiable.

| What | Rule |
|---|---|
| Language | **C++20** |
| Compiler | **MSVC (Visual Studio 2026)** |
| Architecture | **x64 only** |
| UI framework | **Native Win32 API only** |
| Threading | **One UI thread + `std::thread` background workers** |
| IPC | **`PostMessageW` from workers to UI thread only** |
| Persistence | **Lightweight text files only** (CSV, INI) |
| Runtime cache | **In-memory only** |
| Process model | **Single process** |
| String type | **`std::wstring` / `WCHAR`** |
| Resource ownership | **RAII only** (`std::unique_ptr`, custom `Handle` wrappers) |
| Win32 API variant | **Always `W` (wide char) variants** |

### Permanently forbidden — do not introduce these under any circumstances

- C#, WPF, WinUI 3, Windows App SDK, Windows Runtime
- WebView2, Electron, Qt, MFC
- SQLite, any database, any indexing service
- Any second process or IPC architecture other than the single-instance named pipe
- File-content parsing libraries
- NuGet packages for runtime code (Google Test via local source copy is the only allowed exception)
- `using namespace std;` in any header file
- Raw `new` / `delete` for resource ownership
- `wcscpy`, `strcpy`, or any unsafe string copy — use `StringCchCopyW` or `std::wstring`
- `SendMessageW` from a worker thread (can deadlock — use `PostMessageW` only)
- Direct UI control access from any thread other than the UI thread

---

## Three documents, three roles

| Document | Role | When to read |
|---|---|---|
| `AGENTS.md` | Rules and orientation | Start of every session |
| `DEVELOPMENT_JOURNAL.md` | Current state of the project | Start of every session, after AGENTS.md |
| `IMPLEMENTATION_PLAN.md` | Goals, tasks, checkpoints | When starting or resuming a phase |

These three documents must never contradict each other. If they do, resolve in this priority order:
1. PRD wins on **product behavior**
2. IMPLEMENTATION_PLAN.md wins on **technical constraints**
3. AGENTS.md wins on **agent workflow rules**
4. DEVELOPMENT_JOURNAL.md is truth about **current state**

---

## How to start a session

```
1. Read AGENTS.md (this file) — re-read the stack rules above
2. Read DEVELOPMENT_JOURNAL.md — find "Current Status" block
3. Identify the current phase and target checkpoint
4. Read that phase section in IMPLEMENTATION_PLAN.md
5. Look at the checkpoint criteria for that phase
6. Begin work
```

Do not skip this sequence. Do not guess the current state from memory.

---

## How to work

### The loop

```
Write code → Compile → Fix errors → Run tests → Fix failures → Repeat
```

Never leave the loop with compile errors.  
Never leave the loop with failing tests in a completed phase.  
Never proceed to the next phase with a failing checkpoint.

### Checkpoints

Each phase ends with a `CHECKPOINT CP-NN` block in the plan.  
It is a list of binary criteria.  
You must verify each criterion before marking it complete.  
You mark it complete by updating `DEVELOPMENT_JOURNAL.md`.

A checkpoint is only complete when **every criterion is checked**.  
A partial checkpoint is not a completed checkpoint.

### When something breaks

1. Do not hide the failure behind a workaround that papers over the symptom
2. Understand the root cause
3. Fix it at the root
4. If you cannot fix it, log it in `DEVELOPMENT_JOURNAL.md` under "Known Issues" with full detail
5. Do not proceed past a blocker without logging it

---

## How to end a session

Before stopping, always:

```
1. Run all unit tests that exist so far
2. Check the current checkpoint criteria — how many are done?
3. Update DEVELOPMENT_JOURNAL.md:
   - Update "Current Status" block (phase, checkpoint, build status)
   - Add a new Session Log entry
   - Log any decisions made (AD-NNN entries)
   - Log any new known issues
4. Leave a clear "Next session should start with:" note in the session log entry
```

If you do not update the journal, the next session starts blind. That is a failure.

---

## Code rules

### Every source file
- Encoding: **UTF-8 with BOM**
- Include guard: `#pragma once`
- All path and UI strings: `L"..."` prefix
- All Win32 calls: check return value, log failure to `OutputDebugStringW`

### RAII for everything Win32

```cpp
// Correct — RAII wrapper
struct HandleDeleter { void operator()(HANDLE h) { if (h && h != INVALID_HANDLE_VALUE) CloseHandle(h); } };
using UniqueHandle = std::unique_ptr<void, HandleDeleter>;

// Wrong — raw handle
HANDLE h = CreateFileW(...);
// ... forgetting to call CloseHandle
```

Apply the same pattern for: `HICON`, `HFONT`, `HBITMAP`, `HGLOBAL`, `HMENU`, `HDC`, etc.

### Threading rules — enforced strictly

```
[UI Thread]         → only place where Win32 window/control calls happen
[Worker Thread]     → enumerate files, search, fetch siblings
[Communication]     → PostMessageW(hwndMain, WM_FE_*, wParam, lParam) only

Never:              SendMessageW from worker thread
Never:              ListView_SetItem from worker thread
Never:              CreateWindow from worker thread
```

Worker message IDs are defined in a shared header:
```cpp
#define WM_FE_FOLDER_LOADED   (WM_APP + 1)
#define WM_FE_SEARCH_RESULT   (WM_APP + 2)
#define WM_FE_SEARCH_DONE     (WM_APP + 3)
#define WM_FE_SIBLINGS_READY  (WM_APP + 4)
```

### Cancellation pattern — generation counter

```cpp
// In MainWindow or TabManager:
std::atomic<uint64_t> folderGeneration_{0};

// On new navigation:
uint64_t gen = ++folderGeneration_;
// Pass gen to worker

// In worker, before posting results:
if (gen != capturedGen) return; // stale — discard

// In UI message handler:
if (wParam != folderGeneration_) { delete results; return 0; } // stale
```

This pattern is mandatory for all worker results. There must be no path where stale results corrupt the UI.

### Sorting rule

Folders always sort before files, regardless of the active sort column and direction. This is a hard product rule from the PRD.

### Color values

Win32 `COLORREF` is `0x00BBGGRR` (BGR byte order, not RGB).  
The `Colors.h` header defines all colors.  
Do not hardcode color values outside `Colors.h`.

---

## Architecture map

```
MainWindow
  ├── TabManager          (tab state, history, per-tab path)
  ├── TabStrip            (custom-drawn UI, owner-draw)
  ├── NavBar
  │   └── BreadcrumbBar   (display mode + edit mode)
  ├── FileListView        (LVS_OWNERDATA ListView, NM_CUSTOMDRAW)
  ├── Sidebar
  │   ├── SearchBox
  │   └── FavouritesPanel (flyout favs + regular favs)
  └── StatusBar

Workers (background threads, communicate via PostMessageW):
  ├── FolderWorker        (enumerates one folder, posts WM_FE_FOLDER_LOADED)
  └── SearchWorker        (recursive wildcard search, posts WM_FE_SEARCH_RESULT)

Core logic (no Win32 UI, fully testable):
  ├── WildcardMatch       (pattern matching algorithm)
  ├── CsvParser           (read/write CSV)
  ├── SortSettings        (per-folder sort persistence)
  ├── SessionStore        (tab session save/load)
  ├── FavouritesStore     (flyout + regular favourites)
  ├── Settings            (general settings INI)
  ├── IconCache           (HICON cache keyed by extension)
  └── FileOps             (copy/cut/paste/delete/rename wrappers)
```

---

## Test project rules

The test project is `FileExplorer.Tests`.  
It links the core logic library (`FileExplorer.Core.lib`).  
It does **not** link any UI code.  
It uses Google Test, sourced locally in `/extern/googletest/`.

Run tests:
```bat
msbuild FileExplorer.Tests.vcxproj /p:Configuration=Debug /p:Platform=x64
.\x64\Debug\FileExplorer.Tests.exe
```

Expected: `[==========] All tests passed.`

Required test files and their coverage are specified in Phase 13 of `IMPLEMENTATION_PLAN.md`.  
Do not rename or reorganize test files. The plan references them by name.

---

## Persistence file locations

Preference order (portable-first):
1. Same directory as `FileExplorer.exe`
2. `%APPDATA%\FileExplorer\` (fallback if exe directory is read-only)

Files:
| File | Purpose |
|---|---|
| `settings.ini` | General settings (theme, show hidden, etc.) |
| `session.ini` | Tab session restore |
| `sort_settings.csv` | Per-folder sort column + direction |
| `flyout_favourites.csv` | Flyout favourites list |
| `regular_favourites.csv` | Regular favourites list |

No SQLite. No registry writes (optional, avoided for portability). No installer.

---

## PRD hard rules (enforced always)

These come directly from the PRD and must never be violated:

1. **UI thread must never block on directory enumeration or search**
2. **No blank screen flash during navigation** — old content stays until new is ready
3. **Folders always sort before files**
4. **File name color coding** in dark mode: PDF=red, Excel=green, Word=blue, archives=yellow, text/md=teal
5. **Both favourites systems are alphabetically sorted** by friendly name at all times
6. **Favourites persist in CSV files** — not registry, not database
7. **Single instance** — second launch focuses existing window, opens path as new tab
8. **Session restore** — tabs and folders restored on reopen
9. **Per-folder sort memory** — sort column and direction remembered per folder across sessions
10. **All Win32 wide-char (`W`) API variants** — no ANSI API calls

---

## What "done" means

A phase is done when:
- All code compiles with zero errors
- All unit tests that cover that phase pass
- All checkpoint criteria are checked
- `DEVELOPMENT_JOURNAL.md` is updated

A feature is done when:
- It works as the PRD describes
- It handles edge cases (empty folder, missing file, corrupt CSV, cancelled worker)
- It doesn't leak memory or handles
- It doesn't touch the UI from a worker thread

The project is done when CP-15 is signed off and the portable `.exe` runs on a clean Windows 11 machine.

---

*Keep this file at the project root alongside `IMPLEMENTATION_PLAN.md` and `DEVELOPMENT_JOURNAL.md`.*