#include "App.h"

#include <CommCtrl.h>
#include <cwchar>

#include "MainWindow.h"

namespace {

void LogLastError(const wchar_t* context) {
    const DWORD error_code = GetLastError();
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (GetLastError=%lu).\r\n", context, error_code);
    OutputDebugStringW(buffer);
}

}  // namespace

namespace fileexplorer {

int App::Run(HINSTANCE instance, int show_command) {
    if (!SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)) {
        LogLastError(L"SetProcessDpiAwarenessContext");
    }

    INITCOMMONCONTROLSEX common_controls = {};
    common_controls.dwSize = sizeof(common_controls);
    common_controls.dwICC = ICC_STANDARD_CLASSES;
    if (!InitCommonControlsEx(&common_controls)) {
        LogLastError(L"InitCommonControlsEx");
    }

    MainWindow main_window;
    if (!main_window.Create(instance, show_command)) {
        return -1;
    }

    return main_window.MessageLoop();
}

}  // namespace fileexplorer
