#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>

#ifndef WM_FE_FOLDER_LOADED
#define WM_FE_FOLDER_LOADED (WM_APP + 1)
#endif

#ifndef WM_FE_SEARCH_RESULT
#define WM_FE_SEARCH_RESULT (WM_APP + 2)
#endif

#ifndef WM_FE_SEARCH_DONE
#define WM_FE_SEARCH_DONE (WM_APP + 3)
#endif

#ifndef WM_FE_SIBLINGS_READY
#define WM_FE_SIBLINGS_READY (WM_APP + 4)
#endif

namespace fileexplorer {

constexpr UINT WM_FE_ACTIVE_FOLDER_DIRTY = WM_APP + 5;
constexpr UINT WM_FE_IPC_OPEN_PATH = WM_APP + 6;

}  // namespace fileexplorer
