#include "App.h"

#include <CommCtrl.h>
#include <Objbase.h>
#include <Shellapi.h>
#include <cwchar>
#include <algorithm>
#include <atomic>
#include <memory>
#include <string>
#include <thread>

#include "MainWindow.h"
#include "WorkerMessages.h"

namespace {

constexpr wchar_t kSingleInstanceMutexName[] = L"FileExplorerSingleInstanceMutex_v1";
constexpr wchar_t kSingleInstancePipeName[] = L"\\\\.\\pipe\\FileExplorerIPC_v1";
constexpr wchar_t kMainWindowClassName[] = L"FE_MainWindow";

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

class UniqueHandle final {
public:
    UniqueHandle() = default;
    explicit UniqueHandle(HANDLE handle) noexcept : handle_(handle) {}

    UniqueHandle(const UniqueHandle&) = delete;
    UniqueHandle& operator=(const UniqueHandle&) = delete;

    UniqueHandle(UniqueHandle&& other) noexcept : handle_(other.handle_) {
        other.handle_ = nullptr;
    }

    UniqueHandle& operator=(UniqueHandle&& other) noexcept {
        if (this == &other) {
            return *this;
        }
        Reset();
        handle_ = other.handle_;
        other.handle_ = nullptr;
        return *this;
    }

    ~UniqueHandle() {
        Reset();
    }

    HANDLE get() const noexcept {
        return handle_;
    }

    bool valid() const noexcept {
        return handle_ != nullptr && handle_ != INVALID_HANDLE_VALUE;
    }

    void Reset(HANDLE new_handle = nullptr) noexcept {
        if (valid()) {
            if (!CloseHandle(handle_)) {
                LogLastError(L"CloseHandle(UniqueHandle)");
            }
        }
        handle_ = new_handle;
    }

private:
    HANDLE handle_{nullptr};
};

std::wstring TrimWhitespace(const std::wstring& input) {
    const auto begin = input.find_first_not_of(L" \t\r\n");
    if (begin == std::wstring::npos) {
        return L"";
    }
    const auto end = input.find_last_not_of(L" \t\r\n");
    return input.substr(begin, end - begin + 1);
}

std::wstring StripWrappingQuotes(std::wstring input) {
    if (input.size() >= 2 && input.front() == L'"' && input.back() == L'"') {
        input = input.substr(1, input.size() - 2);
    }
    return input;
}

std::wstring NormalizePath(std::wstring path) {
    if (path.empty()) {
        return path;
    }

    std::replace(path.begin(), path.end(), L'/', L'\\');
    if (path.size() > 3) {
        while (!path.empty() && path.back() == L'\\') {
            if (path.size() == 3 && path[1] == L':') {
                break;
            }
            path.pop_back();
        }
    }
    if (path.size() == 2 && path[1] == L':') {
        path.push_back(L'\\');
    }
    return path;
}

std::wstring ExtractRequestedPath(const wchar_t* command_line) {
    std::wstring requested_path;

    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (argv != nullptr) {
        if (argc >= 2 && argv[1] != nullptr) {
            requested_path = argv[1];
        }
        LocalFree(argv);
    }

    if (requested_path.empty() && command_line != nullptr) {
        requested_path = command_line;
    }

    requested_path = StripWrappingQuotes(TrimWhitespace(requested_path));
    return NormalizePath(requested_path);
}

bool WritePathToPipe(const std::wstring& path) {
    const DWORD payload_bytes = static_cast<DWORD>((path.size() + 1) * sizeof(wchar_t));
    for (int attempt = 0; attempt < 25; ++attempt) {
        if (!WaitNamedPipeW(kSingleInstancePipeName, 120)) {
            const DWORD wait_error = GetLastError();
            if (wait_error != ERROR_FILE_NOT_FOUND && wait_error != ERROR_SEM_TIMEOUT) {
                LogLastError(L"WaitNamedPipeW");
            }
            Sleep(40);
            continue;
        }

        UniqueHandle pipe_handle(CreateFileW(
            kSingleInstancePipeName,
            GENERIC_WRITE,
            0,
            nullptr,
            OPEN_EXISTING,
            FILE_ATTRIBUTE_NORMAL,
            nullptr));
        if (!pipe_handle.valid()) {
            Sleep(40);
            continue;
        }

        DWORD bytes_written = 0;
        const bool write_ok =
            WriteFile(pipe_handle.get(), path.c_str(), payload_bytes, &bytes_written, nullptr) != FALSE &&
            bytes_written == payload_bytes;
        if (!write_ok) {
            LogLastError(L"WriteFile(SingleInstancePipe)");
            Sleep(20);
            continue;
        }

        FlushFileBuffers(pipe_handle.get());
        return true;
    }

    return false;
}

void FocusExistingWindowFallback() {
    HWND existing = FindWindowW(kMainWindowClassName, nullptr);
    if (existing == nullptr) {
        return;
    }

    ShowWindowAsync(existing, IsIconic(existing) ? SW_RESTORE : SW_SHOW);

    // Topmost toggle is a practical Win32 fallback to pull a window to the front
    // in cases where foreground activation heuristics are strict.
    SetWindowPos(
        existing,
        HWND_TOPMOST,
        0,
        0,
        0,
        0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOOWNERZORDER | SWP_SHOWWINDOW);
    SetWindowPos(
        existing,
        HWND_NOTOPMOST,
        0,
        0,
        0,
        0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOOWNERZORDER | SWP_SHOWWINDOW);

    SetForegroundWindow(existing);
    BringWindowToTop(existing);
}

void GrantForegroundPermissionToExistingInstance() {
    HWND existing = FindWindowW(kMainWindowClassName, nullptr);
    if (existing == nullptr) {
        return;
    }

    DWORD existing_process_id = 0;
    const DWORD existing_thread_id = GetWindowThreadProcessId(existing, &existing_process_id);
    if (existing_thread_id == 0 || existing_process_id == 0) {
        LogLastError(L"GetWindowThreadProcessId(ExistingMainWindow)");
        return;
    }

    SetLastError(ERROR_SUCCESS);
    if (!LockSetForegroundWindow(LSFW_UNLOCK)) {
        const DWORD error_code = GetLastError();
        if (error_code != ERROR_SUCCESS) {
            LogLastError(L"LockSetForegroundWindow(LSFW_UNLOCK)");
        }
    }

    SetLastError(ERROR_SUCCESS);
    if (!AllowSetForegroundWindow(ASFW_ANY)) {
        const DWORD error_code = GetLastError();
        if (error_code != ERROR_SUCCESS) {
            LogLastError(L"AllowSetForegroundWindow(ASFW_ANY)");
        }
    }

    SetLastError(ERROR_SUCCESS);
    if (!AllowSetForegroundWindow(existing_process_id)) {
        const DWORD error_code = GetLastError();
        if (error_code != ERROR_SUCCESS) {
            LogLastError(L"AllowSetForegroundWindow");
        }
    }
}

void PipeServerLoop(HWND hwnd_target, const std::atomic<bool>* stop_flag) {
    if (hwnd_target == nullptr || stop_flag == nullptr) {
        return;
    }

    while (!stop_flag->load(std::memory_order_acquire)) {
        UniqueHandle pipe_handle(CreateNamedPipeW(
            kSingleInstancePipeName,
            PIPE_ACCESS_INBOUND,
            PIPE_TYPE_MESSAGE | PIPE_READMODE_MESSAGE | PIPE_WAIT,
            1,
            4096,
            4096,
            0,
            nullptr));
        if (!pipe_handle.valid()) {
            LogLastError(L"CreateNamedPipeW");
            Sleep(80);
            continue;
        }

        BOOL connected = ConnectNamedPipe(pipe_handle.get(), nullptr);
        if (!connected) {
            const DWORD connect_error = GetLastError();
            if (connect_error == ERROR_PIPE_CONNECTED) {
                connected = TRUE;
            } else if (connect_error == ERROR_NO_DATA) {
                continue;
            } else {
                LogLastError(L"ConnectNamedPipe");
                continue;
            }
        }

        if (stop_flag->load(std::memory_order_acquire)) {
            DisconnectNamedPipe(pipe_handle.get());
            break;
        }

        wchar_t buffer[2048] = {};
        DWORD bytes_read = 0;
        const bool read_ok = ReadFile(
            pipe_handle.get(),
            buffer,
            static_cast<DWORD>(sizeof(buffer) - sizeof(wchar_t)),
            &bytes_read,
            nullptr) != FALSE;
        if (read_ok && bytes_read > 0) {
            buffer[bytes_read / sizeof(wchar_t)] = L'\0';
            std::wstring requested_path = NormalizePath(std::wstring(buffer));
            auto payload = std::make_unique<std::wstring>(std::move(requested_path));
            if (!PostMessageW(
                    hwnd_target,
                    fileexplorer::WM_FE_IPC_OPEN_PATH,
                    0,
                    reinterpret_cast<LPARAM>(payload.get()))) {
                LogLastError(L"PostMessageW(WM_FE_IPC_OPEN_PATH)");
            } else {
                payload.release();
            }
        }

        DisconnectNamedPipe(pipe_handle.get());
    }
}

}  // namespace

namespace fileexplorer {

int App::Run(HINSTANCE instance, int show_command, const wchar_t* command_line) {
    const std::wstring requested_path = ExtractRequestedPath(command_line);
    UniqueHandle single_instance_mutex(CreateMutexW(nullptr, TRUE, kSingleInstanceMutexName));
    if (!single_instance_mutex.valid()) {
        LogLastError(L"CreateMutexW(SingleInstance)");
    } else if (GetLastError() == ERROR_ALREADY_EXISTS) {
        GrantForegroundPermissionToExistingInstance();
        if (!WritePathToPipe(requested_path)) {
            LogLastError(L"WritePathToPipe(SecondInstance)");
        }
        FocusExistingWindowFallback();
        return 0;
    }

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

    std::atomic<bool> pipe_server_stop{false};
    std::thread pipe_server_thread(PipeServerLoop, main_window.hwnd(), &pipe_server_stop);

    const int exit_code = main_window.MessageLoop();
    pipe_server_stop.store(true, std::memory_order_release);
    WritePathToPipe(L"");
    if (pipe_server_thread.joinable()) {
        pipe_server_thread.join();
    }

    if (SUCCEEDED(com_result)) {
        CoUninitialize();
    }
    return exit_code;
}

}  // namespace fileexplorer
