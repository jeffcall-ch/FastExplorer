#include "App.h"

#include <CommCtrl.h>
#include <Objbase.h>
#include <cwchar>

#include "MainWindow.h"

namespace {

void LogLastError(const wchar_t* context) {
    const DWORD error_code = GetLastError();
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (GetLastError=%lu).\r\n", context, error_code);
    OutputDebugStringW(buffer);
}

void LogHResult(const wchar_t* context, HRESULT result) {
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (HRESULT=0x%08lX).\r\n", context, static_cast<unsigned long>(result));
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

    const HRESULT com_result = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (FAILED(com_result) && com_result != RPC_E_CHANGED_MODE) {
        LogHResult(L"CoInitializeEx", com_result);
    }

    MainWindow main_window;
    if (!main_window.Create(instance, show_command)) {
        if (SUCCEEDED(com_result)) {
            CoUninitialize();
        }
        return -1;
    }

    const int exit_code = main_window.MessageLoop();
    if (SUCCEEDED(com_result)) {
        CoUninitialize();
    }
    return exit_code;
}

}  // namespace fileexplorer
