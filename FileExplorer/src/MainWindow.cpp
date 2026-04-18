#include "MainWindow.h"

#include <CommCtrl.h>
#include <Windowsx.h>
#include <dwmapi.h>

#include <algorithm>
#include <cwchar>
#include <string>
#include <utility>

#include "WorkerMessages.h"

namespace {

constexpr wchar_t kMainWindowClassName[] = L"FE_MainWindow";
constexpr wchar_t kInitialWindowTitle[] = L"FileExplorer";

constexpr int kDefaultClientWidth = 1200;
constexpr int kDefaultClientHeight = 750;

constexpr int kControlIdTabStrip = 1001;
constexpr int kControlIdNavBar = 1002;
constexpr int kControlIdFileList = 1003;
constexpr int kControlIdSidebar = 1004;
constexpr int kControlIdStatusBar = 1005;

constexpr COLORREF kFallbackBackgroundColor = RGB(0x1C, 0x1C, 0x1C);
constexpr COLORREF kZoneTextColor = RGB(0xE8, 0xE8, 0xE8);

constexpr COLORREF kNavBarColor = RGB(0x27, 0x27, 0x27);
constexpr COLORREF kFileListColor = RGB(0x1F, 0x1F, 0x1F);
constexpr COLORREF kSidebarColor = RGB(0x25, 0x25, 0x25);
constexpr COLORREF kStatusBarColor = RGB(0x23, 0x23, 0x23);
constexpr UINT_PTR kFolderRefreshDebounceTimerId = 9201;
constexpr UINT kFolderRefreshDebounceMs = 500;

constexpr DWORD kDwmaUseImmersiveDarkMode = 20;
constexpr DWORD kDwmaSystemBackdropType = 38;
constexpr int kDwmSbtMainWindow = 2;

void LogLastError(const wchar_t* context) {
    const DWORD error_code = GetLastError();
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (GetLastError=%lu).\r\n", context, error_code);
    OutputDebugStringW(buffer);
}

void LogHResult(const wchar_t* context, HRESULT hr) {
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (HRESULT=0x%08lX).\r\n", context, static_cast<unsigned long>(hr));
    OutputDebugStringW(buffer);
}

RECT CenteredWindowRect(DWORD window_style, DWORD window_ex_style) {
    RECT window_rect = {0, 0, kDefaultClientWidth, kDefaultClientHeight};
    const UINT system_dpi = GetDpiForSystem();
    if (!AdjustWindowRectExForDpi(&window_rect, window_style, FALSE, window_ex_style, system_dpi)) {
        LogLastError(L"AdjustWindowRectExForDpi");
        if (!AdjustWindowRectEx(&window_rect, window_style, FALSE, window_ex_style)) {
            LogLastError(L"AdjustWindowRectEx");
        }
    }

    const int width = window_rect.right - window_rect.left;
    const int height = window_rect.bottom - window_rect.top;
    const int left = (GetSystemMetrics(SM_CXSCREEN) - width) / 2;
    const int top = (GetSystemMetrics(SM_CYSCREEN) - height) / 2;
    return {left, top, left + width, top + height};
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

bool PathsEqualInsensitive(const std::wstring& left, const std::wstring& right) {
    const std::wstring normalized_left = NormalizePath(left);
    const std::wstring normalized_right = NormalizePath(right);
    const int result = CompareStringOrdinal(
        normalized_left.c_str(),
        static_cast<int>(normalized_left.size()),
        normalized_right.c_str(),
        static_cast<int>(normalized_right.size()),
        TRUE);
    return result == CSTR_EQUAL;
}

std::wstring ParentPath(const std::wstring& path) {
    const std::wstring normalized = NormalizePath(path);
    if (normalized.empty()) {
        return L"";
    }
    if (normalized.size() == 3 && normalized[1] == L':') {
        return L"";
    }

    const std::wstring::size_type separator = normalized.find_last_of(L'\\');
    if (separator == std::wstring::npos) {
        return L"";
    }
    if (separator == 2 && normalized[1] == L':') {
        return normalized.substr(0, 3);
    }
    if (separator == 0) {
        return L"";
    }
    return normalized.substr(0, separator);
}

std::wstring LeafName(const std::wstring& path) {
    const std::wstring normalized = NormalizePath(path);
    if (normalized.empty()) {
        return L"";
    }
    if (normalized.size() == 3 && normalized[1] == L':') {
        return normalized.substr(0, 2);
    }

    const std::wstring::size_type separator = normalized.find_last_of(L'\\');
    if (separator == std::wstring::npos) {
        return normalized;
    }
    if (separator + 1 >= normalized.size()) {
        return normalized;
    }
    return normalized.substr(separator + 1);
}

}  // namespace

namespace fileexplorer {

void MainWindow::BrushDeleter::operator()(HBRUSH brush) const noexcept {
    if (brush != nullptr) {
        DeleteObject(brush);
    }
}

MainWindow::MainWindow()
    : metrics_(DefaultLayoutMetrics()),
      folder_generation_source_(std::make_shared<std::atomic<uint64_t>>(0)) {}

MainWindow::~MainWindow() {
    StopFolderWatcher();
    DestroyZoneBrushes();
}

bool MainWindow::Create(HINSTANCE instance, int show_command) {
    instance_ = instance;
    if (!RegisterWindowClass()) {
        return false;
    }

    if (!CreateMainWindow(show_command)) {
        return false;
    }

    return true;
}

int MainWindow::MessageLoop() {
    MSG message = {};
    while (true) {
        const BOOL result = GetMessageW(&message, nullptr, 0, 0);
        if (result == -1) {
            LogLastError(L"GetMessageW");
            return -1;
        }

        if (result == 0) {
            return static_cast<int>(message.wParam);
        }

        TranslateMessage(&message);
        DispatchMessageW(&message);
    }
}

bool MainWindow::RegisterWindowClass() {
    WNDCLASSEXW window_class = {};
    window_class.cbSize = sizeof(window_class);
    window_class.style = CS_HREDRAW | CS_VREDRAW;
    window_class.lpfnWndProc = &MainWindow::WindowProc;
    window_class.hInstance = instance_;
    window_class.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    if (window_class.hCursor == nullptr) {
        LogLastError(L"LoadCursorW");
    }
    window_class.hbrBackground = nullptr;
    window_class.lpszClassName = kMainWindowClassName;

    const ATOM class_id = RegisterClassExW(&window_class);
    if (class_id == 0) {
        if (GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
            LogLastError(L"RegisterClassExW");
            return false;
        }
    }

    return true;
}

bool MainWindow::CreateMainWindow(int show_command) {
    const DWORD window_style = WS_OVERLAPPEDWINDOW;
    const DWORD window_ex_style = 0;
    const RECT positioned_rect = CenteredWindowRect(window_style, window_ex_style);
    const int width = positioned_rect.right - positioned_rect.left;
    const int height = positioned_rect.bottom - positioned_rect.top;

    hwnd_ = CreateWindowExW(
        window_ex_style,
        kMainWindowClassName,
        kInitialWindowTitle,
        window_style,
        positioned_rect.left,
        positioned_rect.top,
        width,
        height,
        nullptr,
        nullptr,
        instance_,
        this);

    if (hwnd_ == nullptr) {
        LogLastError(L"CreateWindowExW(MainWindow)");
        return false;
    }

    ShowWindow(hwnd_, show_command);
    UpdateWindow(hwnd_);
    return true;
}

bool MainWindow::InitializeChildZones() {
    const DWORD child_style = WS_CHILD | WS_VISIBLE | WS_BORDER | SS_CENTER | SS_CENTERIMAGE;
    const DWORD child_ex_style = 0;

    if (!tab_strip_.Create(hwnd_, instance_, kControlIdTabStrip)) {
        return false;
    }
    tab_strip_.SetTabManager(&tab_manager_);
    tab_strip_.SetDpi(dpi_);
    tab_strip_hwnd_ = tab_strip_.hwnd();
    if (tab_strip_hwnd_ == nullptr) {
        LogLastError(L"TabStrip hwnd");
        return false;
    }

    if (!nav_bar_.Create(hwnd_, instance_, kControlIdNavBar)) {
        return false;
    }
    nav_bar_.SetDpi(dpi_);
    nav_bar_hwnd_ = nav_bar_.hwnd();
    if (nav_bar_hwnd_ == nullptr) {
        LogLastError(L"NavBar hwnd");
        return false;
    }

    if (!file_list_view_.Create(hwnd_, instance_, kControlIdFileList)) {
        return false;
    }
    file_list_view_.SetDpi(dpi_);
    file_list_hwnd_ = file_list_view_.hwnd();
    if (file_list_hwnd_ == nullptr) {
        LogLastError(L"FileListView hwnd");
        return false;
    }

    if (!sidebar_.Create(hwnd_, instance_, kControlIdSidebar)) {
        return false;
    }
    sidebar_.SetDpi(dpi_);
    sidebar_hwnd_ = sidebar_.hwnd();
    if (sidebar_hwnd_ == nullptr) {
        LogLastError(L"Sidebar hwnd");
        return false;
    }

    status_bar_hwnd_ = CreateWindowExW(
        child_ex_style,
        WC_STATICW,
        L"Status Bar Zone",
        child_style,
        0,
        0,
        0,
        0,
        hwnd_,
        reinterpret_cast<HMENU>(static_cast<INT_PTR>(kControlIdStatusBar)),
        instance_,
        nullptr);
    if (status_bar_hwnd_ == nullptr) {
        LogLastError(L"CreateWindowExW(StatusBarZone)");
        return false;
    }

    return true;
}

void MainWindow::CreateZoneBrushes() {
    nav_brush_.reset(CreateSolidBrush(kNavBarColor));
    if (!nav_brush_) {
        LogLastError(L"CreateSolidBrush(NavBar)");
    }

    file_list_brush_.reset(CreateSolidBrush(kFileListColor));
    if (!file_list_brush_) {
        LogLastError(L"CreateSolidBrush(FileList)");
    }

    sidebar_brush_.reset(CreateSolidBrush(kSidebarColor));
    if (!sidebar_brush_) {
        LogLastError(L"CreateSolidBrush(Sidebar)");
    }

    status_brush_.reset(CreateSolidBrush(kStatusBarColor));
    if (!status_brush_) {
        LogLastError(L"CreateSolidBrush(StatusBar)");
    }
}

void MainWindow::DestroyZoneBrushes() {
    nav_brush_.reset();
    file_list_brush_.reset();
    sidebar_brush_.reset();
    status_brush_.reset();
}

void MainWindow::ApplyWindowChrome() {
    BOOL dark_mode = TRUE;
    const HRESULT dark_mode_result =
        DwmSetWindowAttribute(hwnd_, kDwmaUseImmersiveDarkMode, &dark_mode, sizeof(dark_mode));
    if (FAILED(dark_mode_result)) {
        LogHResult(L"DwmSetWindowAttribute(DWMWA_USE_IMMERSIVE_DARK_MODE)", dark_mode_result);
    }

    int backdrop_type = kDwmSbtMainWindow;
    const HRESULT backdrop_result =
        DwmSetWindowAttribute(hwnd_, kDwmaSystemBackdropType, &backdrop_type, sizeof(backdrop_type));
    if (FAILED(backdrop_result)) {
        use_solid_fallback_background_ = true;
        LogHResult(L"DwmSetWindowAttribute(DWMWA_SYSTEMBACKDROP_TYPE)", backdrop_result);
    } else {
        use_solid_fallback_background_ = false;
    }

    if (!InvalidateRect(hwnd_, nullptr, TRUE)) {
        LogLastError(L"InvalidateRect");
    }
}

void MainWindow::RecalculateLayout() {
    metrics_ = ScaleLayoutMetrics(dpi_);
}

void MainWindow::LayoutChildZones() {
    if (hwnd_ == nullptr ||
        tab_strip_hwnd_ == nullptr ||
        nav_bar_hwnd_ == nullptr ||
        file_list_hwnd_ == nullptr ||
        sidebar_hwnd_ == nullptr ||
        status_bar_hwnd_ == nullptr) {
        return;
    }

    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        LogLastError(L"GetClientRect");
        return;
    }

    const int client_width = std::max(0, static_cast<int>(client_rect.right - client_rect.left));
    const int client_height = std::max(0, static_cast<int>(client_rect.bottom - client_rect.top));

    const int tab_height = std::min(metrics_.tabStripHeight, client_height);
    const int nav_height = std::min(metrics_.navBarHeight, std::max(0, client_height - tab_height));
    const int status_height =
        std::min(metrics_.statusBarHeight, std::max(0, client_height - tab_height - nav_height));

    const int content_top = tab_height + nav_height;
    const int content_bottom = client_height - status_height;
    const int content_height = std::max(0, content_bottom - content_top);

    const int sidebar_width = std::min(metrics_.sidebarWidth, client_width);
    const int file_list_width = std::max(0, client_width - sidebar_width);

    if (!MoveWindow(tab_strip_hwnd_, 0, 0, client_width, tab_height, TRUE)) {
        LogLastError(L"MoveWindow(TabStrip)");
    }
    if (!MoveWindow(nav_bar_hwnd_, 0, tab_height, client_width, nav_height, TRUE)) {
        LogLastError(L"MoveWindow(NavBarZone)");
    }
    if (!MoveWindow(file_list_hwnd_, 0, content_top, file_list_width, content_height, TRUE)) {
        LogLastError(L"MoveWindow(FileListZone)");
    }
    if (!MoveWindow(sidebar_hwnd_, file_list_width, content_top, sidebar_width, content_height, TRUE)) {
        LogLastError(L"MoveWindow(SidebarZone)");
    }
    if (!MoveWindow(status_bar_hwnd_, 0, client_height - status_height, client_width, status_height, TRUE)) {
        LogLastError(L"MoveWindow(StatusBarZone)");
    }
}

void MainWindow::OnDpiChanged(UINT new_dpi, RECT* suggested_rect) {
    dpi_ = (new_dpi == 0U) ? 96U : new_dpi;
    RecalculateLayout();
    tab_strip_.SetDpi(dpi_);
    nav_bar_.SetDpi(dpi_);
    file_list_view_.SetDpi(dpi_);
    sidebar_.SetDpi(dpi_);

    if (suggested_rect != nullptr) {
        if (!SetWindowPos(
                hwnd_,
                nullptr,
                suggested_rect->left,
                suggested_rect->top,
                suggested_rect->right - suggested_rect->left,
                suggested_rect->bottom - suggested_rect->top,
                SWP_NOZORDER | SWP_NOACTIVATE)) {
            LogLastError(L"SetWindowPos(WM_DPICHANGED)");
        }
    }

    LayoutChildZones();
}

void MainWindow::PaintFallbackBackground(HDC hdc) {
    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        LogLastError(L"GetClientRect(WM_ERASEBKGND)");
        return;
    }

    HBRUSH fallback_brush = CreateSolidBrush(kFallbackBackgroundColor);
    if (fallback_brush == nullptr) {
        LogLastError(L"CreateSolidBrush(FallbackBackground)");
        return;
    }

    if (FillRect(hdc, &client_rect, fallback_brush) == 0) {
        LogLastError(L"FillRect(FallbackBackground)");
    }

    if (!DeleteObject(fallback_brush)) {
        LogLastError(L"DeleteObject(FallbackBackgroundBrush)");
    }
}

void MainWindow::UpdateWindowTitle() {
    std::wstring title = kInitialWindowTitle;
    if (const TabState* active = tab_manager_.active_tab(); active != nullptr && !active->displayName.empty()) {
        title = active->displayName + L" - FileExplorer";
    }

    if (!SetWindowTextW(hwnd_, title.c_str())) {
        LogLastError(L"SetWindowTextW");
    }
}

void MainWindow::HandleTabStateChanged() {
    tab_strip_.Refresh();

    if (const TabState* active = tab_manager_.active_tab(); active != nullptr) {
        nav_bar_.SetPath(active->path);
        sidebar_.SetCurrentPath(active->path);
        const std::wstring post_load_selection_name = std::move(pending_post_load_selection_name_);
        pending_post_load_selection_name_.clear();
        StartFolderLoad(active->path, false, post_load_selection_name);
        RestartFolderWatcher(active->path);
    } else {
        sidebar_.SetCurrentPath(L"");
        pending_post_load_selection_name_.clear();
        StopFolderWatcher();
    }
    nav_bar_.SetNavigationState(
        tab_manager_.CanNavigateBack(),
        tab_manager_.CanNavigateForward(),
        tab_manager_.CanNavigateUp());

    UpdateWindowTitle();
}

void MainWindow::StartFolderLoad(
    const std::wstring& path,
    bool incremental_refresh,
    const std::wstring& post_load_select_name) {
    if (folder_generation_source_ == nullptr) {
        folder_generation_source_ = std::make_shared<std::atomic<uint64_t>>(0);
    }

    const std::wstring normalized_path = NormalizePath(path);
    pending_incremental_refresh_ = incremental_refresh;
    if (incremental_refresh) {
        file_list_view_.SetPostLoadSelectionName(L"");
    } else {
        file_list_view_.SetPostLoadSelectionName(post_load_select_name);
    }
    file_list_view_.BeginFolderLoad(normalized_path, incremental_refresh);

    const uint64_t generation = ++(*folder_generation_source_);
    FolderWorker::Request request = {};
    request.path = normalized_path;
    request.generation = generation;
    request.hwnd_target = hwnd_;
    request.generation_source = folder_generation_source_;
    FolderWorker::Start(std::move(request));
}

void MainWindow::PreparePostLoadSelectionForTransition(
    const std::wstring& from_path,
    const std::wstring& to_path) {
    pending_post_load_selection_name_.clear();

    const std::wstring normalized_from = NormalizePath(from_path);
    const std::wstring normalized_to = NormalizePath(to_path);
    if (normalized_from.empty() || normalized_to.empty()) {
        return;
    }

    if (NormalizePath(ParentPath(normalized_from)) != normalized_to) {
        return;
    }

    pending_post_load_selection_name_ = LeafName(normalized_from);
}

void MainWindow::StartDebouncedRefreshForActiveFolder(uint64_t watch_generation) {
    pending_debounced_refresh_ = true;
    pending_debounced_refresh_generation_ = watch_generation;
    if (SetTimer(hwnd_, kFolderRefreshDebounceTimerId, kFolderRefreshDebounceMs, nullptr) == 0) {
        LogLastError(L"SetTimer(FolderRefreshDebounce)");
    }
}

void MainWindow::RestartFolderWatcher(const std::wstring& path) {
    KillTimer(hwnd_, kFolderRefreshDebounceTimerId);
    pending_debounced_refresh_ = false;
    pending_debounced_refresh_generation_ = 0;
    StopFolderWatcher();

    const std::wstring normalized_path = NormalizePath(path);
    if (normalized_path.empty()) {
        return;
    }

    auto stop_flag = std::make_shared<std::atomic<bool>>(false);
    folder_watch_stop_ = stop_flag;
    watched_folder_path_ = normalized_path;
    const uint64_t watch_generation = ++folder_watch_generation_;
    const HWND target_hwnd = hwnd_;

    folder_watch_thread_ = std::thread([target_hwnd, normalized_path, watch_generation, stop_flag]() {
        HANDLE change_handle = FindFirstChangeNotificationW(
            normalized_path.c_str(),
            FALSE,
            FILE_NOTIFY_CHANGE_FILE_NAME |
                FILE_NOTIFY_CHANGE_DIR_NAME |
                FILE_NOTIFY_CHANGE_LAST_WRITE);
        if (change_handle == INVALID_HANDLE_VALUE) {
            LogLastError(L"FindFirstChangeNotificationW");
            return;
        }

        while (!stop_flag->load(std::memory_order_acquire)) {
            const DWORD wait_result = WaitForSingleObject(change_handle, 200);
            if (wait_result == WAIT_TIMEOUT) {
                continue;
            }

            if (wait_result == WAIT_OBJECT_0) {
                if (!stop_flag->load(std::memory_order_acquire)) {
                    if (!PostMessageW(
                            target_hwnd,
                            WM_FE_ACTIVE_FOLDER_DIRTY,
                            static_cast<WPARAM>(watch_generation),
                            0)) {
                        LogLastError(L"PostMessageW(WM_FE_ACTIVE_FOLDER_DIRTY)");
                        break;
                    }
                }

                if (!FindNextChangeNotification(change_handle)) {
                    LogLastError(L"FindNextChangeNotification");
                    break;
                }
                continue;
            }

            LogLastError(L"WaitForSingleObject(FolderWatcher)");
            break;
        }

        if (!FindCloseChangeNotification(change_handle)) {
            LogLastError(L"FindCloseChangeNotification");
        }
    });
}

void MainWindow::StopFolderWatcher() {
    if (folder_watch_stop_) {
        folder_watch_stop_->store(true, std::memory_order_release);
    }

    if (folder_watch_thread_.joinable()) {
        folder_watch_thread_.join();
    }

    folder_watch_stop_.reset();
    watched_folder_path_.clear();
    pending_debounced_refresh_ = false;
    pending_debounced_refresh_generation_ = 0;
}

bool MainWindow::HandleTabKeyboardShortcut(WPARAM key_code, bool ctrl_down, bool shift_down) {
    if (!ctrl_down) {
        return false;
    }

    bool handled = true;
    bool changed = false;

    switch (key_code) {
    case 'T':
        changed = tab_manager_.AddTab(TabManager::DefaultTabPath(), true, false) >= 0;
        break;

    case 'W':
        changed = tab_manager_.CloseActiveTab();
        break;

    case VK_TAB:
        changed = shift_down ? tab_manager_.ActivatePrevious() : tab_manager_.ActivateNext();
        break;

    case '1':
    case '2':
    case '3':
    case '4':
    case '5':
    case '6':
    case '7':
    case '8':
        changed = tab_manager_.JumpToIndex(static_cast<int>(key_code - '1'));
        break;

    case '9': {
        const int last_index = static_cast<int>(tab_manager_.tabs().size()) - 1;
        if (last_index >= 0) {
            changed = tab_manager_.JumpToIndex(last_index);
        }
        break;
    }

    case 'L':
        nav_bar_.ActivateAddressEditMode();
        break;

    case 'F':
        if (sidebar_hwnd_ != nullptr) {
            SetFocus(sidebar_hwnd_);
        }
        break;

    default:
        handled = false;
        break;
    }

    if (handled && changed) {
        HandleTabStateChanged();
    }
    return handled;
}

LRESULT CALLBACK MainWindow::WindowProc(HWND hwnd, UINT message, WPARAM w_param, LPARAM l_param) {
    MainWindow* self = nullptr;
    if (message == WM_NCCREATE) {
        auto* create_struct = reinterpret_cast<CREATESTRUCTW*>(l_param);
        self = reinterpret_cast<MainWindow*>(create_struct->lpCreateParams);
        if (self == nullptr) {
            return FALSE;
        }

        self->hwnd_ = hwnd;
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    } else {
        self = reinterpret_cast<MainWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    }

    if (self != nullptr) {
        return self->HandleMessage(message, w_param, l_param);
    }

    return DefWindowProcW(hwnd, message, w_param, l_param);
}

LRESULT MainWindow::HandleMessage(UINT message, WPARAM w_param, LPARAM l_param) {
    switch (message) {
    case WM_CREATE:
        dpi_ = GetDpiForWindow(hwnd_);
        if (dpi_ == 0U) {
            dpi_ = 96U;
        }
        RecalculateLayout();
        CreateZoneBrushes();
        if (!InitializeChildZones()) {
            return -1;
        }
        ApplyWindowChrome();
        UpdateWindowTitle();
        LayoutChildZones();
        HandleTabStateChanged();
        return 0;

    case WM_SIZE:
        LayoutChildZones();
        return 0;

    case WM_DPICHANGED:
        OnDpiChanged(HIWORD(w_param), reinterpret_cast<RECT*>(l_param));
        return 0;

    case WM_KEYDOWN: {
        const bool ctrl_down = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        const bool shift_down = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        if (HandleTabKeyboardShortcut(w_param, ctrl_down, shift_down)) {
            return 0;
        }
        if (w_param == VK_F5) {
            if (const TabState* active = tab_manager_.active_tab(); active != nullptr) {
                StartFolderLoad(active->path, true);
            }
            return 0;
        }
        break;
    }

    case WM_SYSKEYDOWN:
        if ((GetKeyState(VK_MENU) & 0x8000) != 0 && (w_param == 'D' || w_param == 'd')) {
            nav_bar_.ActivateAddressEditMode();
            return 0;
        }
        break;

    case WM_NCHITTEST: {
        const LRESULT hit = DefWindowProcW(hwnd_, message, w_param, l_param);
        if (hit == HTCLIENT) {
            POINT screen_point = {GET_X_LPARAM(l_param), GET_Y_LPARAM(l_param)};
            POINT client_point = screen_point;
            if (!ScreenToClient(hwnd_, &client_point)) {
                LogLastError(L"ScreenToClient(WM_NCHITTEST)");
                return hit;
            }
            if (tab_strip_.IsDraggableGap(client_point)) {
                return HTCAPTION;
            }
        }
        return hit;
    }

    case WM_FE_TAB_CHANGED:
        pending_post_load_selection_name_.clear();
        HandleTabStateChanged();
        return 0;

    case WM_FE_NAV_COMMAND: {
        const std::wstring previous_path =
            (tab_manager_.active_tab() != nullptr) ? tab_manager_.active_tab()->path : L"";
        bool changed = false;
        bool handled_refresh = false;
        switch (static_cast<NavCommand>(w_param)) {
        case NavCommand::Back:
            changed = tab_manager_.NavigateBack();
            break;
        case NavCommand::Forward:
            changed = tab_manager_.NavigateForward();
            break;
        case NavCommand::Up:
            changed = tab_manager_.NavigateUp();
            break;
        case NavCommand::Refresh:
            handled_refresh = tab_manager_.RefreshActive();
            if (handled_refresh) {
                if (const TabState* active = tab_manager_.active_tab(); active != nullptr) {
                    StartFolderLoad(active->path, true);
                    RestartFolderWatcher(active->path);
                }
                nav_bar_.SetNavigationState(
                    tab_manager_.CanNavigateBack(),
                    tab_manager_.CanNavigateForward(),
                    tab_manager_.CanNavigateUp());
                UpdateWindowTitle();
            }
            break;
        default:
            break;
        }

        if (!handled_refresh && changed) {
            if (const TabState* active = tab_manager_.active_tab(); active != nullptr) {
                PreparePostLoadSelectionForTransition(previous_path, active->path);
            }
            HandleTabStateChanged();
        }
        return 0;
    }

    case WM_FE_BREADCRUMB_NAVIGATE: {
        const std::wstring previous_path =
            (tab_manager_.active_tab() != nullptr) ? tab_manager_.active_tab()->path : L"";
        auto requested_path = std::unique_ptr<std::wstring>(reinterpret_cast<std::wstring*>(l_param));
        if (requested_path && tab_manager_.NavigateTo(*requested_path)) {
            if (const TabState* active = tab_manager_.active_tab(); active != nullptr) {
                PreparePostLoadSelectionForTransition(previous_path, active->path);
            }
            HandleTabStateChanged();
        }
        return 0;
    }

    case WM_FE_SIDEBAR_NAVIGATE: {
        auto requested_path = std::unique_ptr<std::wstring>(reinterpret_cast<std::wstring*>(l_param));
        if (requested_path) {
            const std::wstring target_path = NormalizePath(*requested_path);
            if (target_path.empty()) {
                return 0;
            }

            {
                wchar_t buffer[512] = {};
                swprintf_s(
                    buffer,
                    L"[FileExplorer] WM_FE_SIDEBAR_NAVIGATE target=%s\r\n",
                    target_path.c_str());
                OutputDebugStringW(buffer);
            }

            const auto& tabs = tab_manager_.tabs();
            int existing_index = -1;
            for (int i = 0; i < static_cast<int>(tabs.size()); ++i) {
                if (PathsEqualInsensitive(tabs[i].path, target_path)) {
                    existing_index = i;
                    break;
                }
            }

            if (existing_index >= 0) {
                if (tab_manager_.Activate(existing_index)) {
                    HandleTabStateChanged();
                }
            } else if (tab_manager_.AddTab(target_path, true, false) >= 0) {
                HandleTabStateChanged();
            }
        }
        return 0;
    }

    case WM_FE_FILELIST_NAVIGATE: {
        const std::wstring previous_path =
            (tab_manager_.active_tab() != nullptr) ? tab_manager_.active_tab()->path : L"";
        auto requested_path = std::unique_ptr<std::wstring>(reinterpret_cast<std::wstring*>(l_param));
        if (requested_path && tab_manager_.NavigateTo(*requested_path)) {
            if (const TabState* active = tab_manager_.active_tab(); active != nullptr) {
                PreparePostLoadSelectionForTransition(previous_path, active->path);
            }
            HandleTabStateChanged();
        }
        return 0;
    }

    case WM_FE_FILELIST_REFRESH:
        if (const TabState* active = tab_manager_.active_tab(); active != nullptr) {
            StartFolderLoad(active->path, true);
            RestartFolderWatcher(active->path);
        }
        return 0;

    case WM_FE_FILELIST_STATUS_UPDATE: {
        auto status_text = std::unique_ptr<std::wstring>(reinterpret_cast<std::wstring*>(l_param));
        if (status_text && status_bar_hwnd_ != nullptr) {
            if (!SetWindowTextW(status_bar_hwnd_, status_text->c_str())) {
                LogLastError(L"SetWindowTextW(StatusBar)");
            }
        }
        return 0;
    }

    case WM_FE_FOLDER_LOADED: {
        auto loaded_entries = std::unique_ptr<std::vector<FileEntry>>(reinterpret_cast<std::vector<FileEntry>*>(l_param));
        if (!loaded_entries || folder_generation_source_ == nullptr) {
            return 0;
        }

        const uint64_t generation = static_cast<uint64_t>(w_param);
        if (generation != folder_generation_source_->load(std::memory_order_acquire)) {
            return 0;
        }

        file_list_view_.ApplyLoadedFolderEntries(std::move(*loaded_entries), pending_incremental_refresh_);
        pending_incremental_refresh_ = false;
        return 0;
    }

    case WM_FE_ACTIVE_FOLDER_DIRTY: {
        const uint64_t watch_generation = static_cast<uint64_t>(w_param);
        if (watch_generation != folder_watch_generation_.load(std::memory_order_acquire)) {
            return 0;
        }
        StartDebouncedRefreshForActiveFolder(watch_generation);
        return 0;
    }

    case WM_TIMER:
        if (w_param == kFolderRefreshDebounceTimerId) {
            KillTimer(hwnd_, kFolderRefreshDebounceTimerId);
            if (pending_debounced_refresh_ &&
                pending_debounced_refresh_generation_ == folder_watch_generation_.load(std::memory_order_acquire)) {
                pending_debounced_refresh_ = false;
                pending_debounced_refresh_generation_ = 0;
                if (const TabState* active = tab_manager_.active_tab(); active != nullptr) {
                    StartFolderLoad(active->path, true);
                }
            } else {
                pending_debounced_refresh_ = false;
                pending_debounced_refresh_generation_ = 0;
            }
            return 0;
        }
        break;

    case WM_NOTIFY: {
        LRESULT notify_result = 0;
        if (file_list_view_.HandleNotify(l_param, &notify_result)) {
            return notify_result;
        }
        break;
    }

    case WM_ERASEBKGND:
        if (use_solid_fallback_background_) {
            PaintFallbackBackground(reinterpret_cast<HDC>(w_param));
            return 1;
        }
        break;

    case WM_CTLCOLORSTATIC: {
        const HDC hdc = reinterpret_cast<HDC>(w_param);
        const HWND control_hwnd = reinterpret_cast<HWND>(l_param);

        COLORREF background_color = kFileListColor;
        HBRUSH background_brush = file_list_brush_.get();

        if (control_hwnd == nav_bar_hwnd_) {
            background_color = kNavBarColor;
            background_brush = nav_brush_.get();
        } else if (control_hwnd == sidebar_hwnd_) {
            background_color = kSidebarColor;
            background_brush = sidebar_brush_.get();
        } else if (control_hwnd == status_bar_hwnd_) {
            background_color = kStatusBarColor;
            background_brush = status_brush_.get();
        }

        if (hdc != nullptr) {
            SetTextColor(hdc, kZoneTextColor);
            SetBkColor(hdc, background_color);
            SetBkMode(hdc, OPAQUE);
        }

        if (background_brush != nullptr) {
            return reinterpret_cast<LRESULT>(background_brush);
        }
        return reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_WINDOW));
    }

    case WM_CLOSE:
        if (!DestroyWindow(hwnd_)) {
            LogLastError(L"DestroyWindow(WM_CLOSE)");
        }
        return 0;

    case WM_DESTROY:
        KillTimer(hwnd_, kFolderRefreshDebounceTimerId);
        pending_debounced_refresh_ = false;
        pending_debounced_refresh_generation_ = 0;
        StopFolderWatcher();
        PostQuitMessage(0);
        return 0;

    case WM_NCDESTROY:
        SetWindowLongPtrW(hwnd_, GWLP_USERDATA, 0);
        hwnd_ = nullptr;
        break;

    default:
        break;
    }

    return DefWindowProcW(hwnd_, message, w_param, l_param);
}

}  // namespace fileexplorer
