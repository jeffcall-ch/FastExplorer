#include "NavBar.h"

#include <CommCtrl.h>

#include <algorithm>
#include <cwchar>

namespace {

constexpr wchar_t kNavBarClassName[] = L"FE_NavBar";

constexpr COLORREF kNavBarBackground = RGB(0x27, 0x27, 0x27);

void LogLastError(const wchar_t* context) {
    const DWORD error_code = GetLastError();
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (GetLastError=%lu).\r\n", context, error_code);
    OutputDebugStringW(buffer);
}

}  // namespace

namespace fileexplorer {

NavBar::NavBar() = default;

NavBar::~NavBar() = default;

bool NavBar::Create(HWND parent, HINSTANCE instance, int control_id) {
    parent_hwnd_ = parent;
    instance_ = instance;
    control_id_ = control_id;

    if (!RegisterWindowClass()) {
        return false;
    }

    hwnd_ = CreateWindowExW(
        0,
        kNavBarClassName,
        L"",
        WS_CHILD | WS_VISIBLE,
        0,
        0,
        0,
        0,
        parent,
        reinterpret_cast<HMENU>(static_cast<INT_PTR>(control_id)),
        instance,
        this);
    if (hwnd_ == nullptr) {
        LogLastError(L"CreateWindowExW(NavBar)");
        return false;
    }

    return true;
}

HWND NavBar::hwnd() const noexcept {
    return hwnd_;
}

void NavBar::SetDpi(UINT dpi) {
    dpi_ = (dpi == 0U) ? 96U : dpi;
    breadcrumb_bar_.SetDpi(dpi_);
    LayoutChildControls();
    Refresh();
}

void NavBar::SetPath(const std::wstring& path) {
    breadcrumb_bar_.SetPath(path);
}

void NavBar::SetNavigationState(bool can_back, bool can_forward, bool can_up) {
    if (back_button_hwnd_ != nullptr) {
        EnableWindow(back_button_hwnd_, can_back ? TRUE : FALSE);
    }
    if (forward_button_hwnd_ != nullptr) {
        EnableWindow(forward_button_hwnd_, can_forward ? TRUE : FALSE);
    }
    if (up_button_hwnd_ != nullptr) {
        EnableWindow(up_button_hwnd_, can_up ? TRUE : FALSE);
    }
}

void NavBar::ActivateAddressEditMode() {
    breadcrumb_bar_.ActivateEditMode();
}

void NavBar::Refresh() {
    if (hwnd_ != nullptr) {
        InvalidateRect(hwnd_, nullptr, TRUE);
    }
}

bool NavBar::RegisterWindowClass() {
    WNDCLASSEXW window_class = {};
    window_class.cbSize = sizeof(window_class);
    window_class.style = CS_HREDRAW | CS_VREDRAW;
    window_class.lpfnWndProc = &NavBar::WindowProc;
    window_class.hInstance = instance_;
    window_class.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    window_class.hbrBackground = nullptr;
    window_class.lpszClassName = kNavBarClassName;

    const ATOM class_id = RegisterClassExW(&window_class);
    if (class_id == 0 && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
        LogLastError(L"RegisterClassExW(FE_NavBar)");
        return false;
    }

    return true;
}

bool NavBar::CreateChildControls() {
    constexpr DWORD button_style = WS_CHILD | WS_VISIBLE | BS_FLAT | BS_PUSHBUTTON;

    back_button_hwnd_ = CreateWindowExW(
        0,
        WC_BUTTONW,
        L"<",
        button_style,
        0,
        0,
        0,
        0,
        hwnd_,
        reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_NAV_BACK)),
        instance_,
        nullptr);
    if (back_button_hwnd_ == nullptr) {
        LogLastError(L"CreateWindowExW(NavBack)");
        return false;
    }

    forward_button_hwnd_ = CreateWindowExW(
        0,
        WC_BUTTONW,
        L">",
        button_style,
        0,
        0,
        0,
        0,
        hwnd_,
        reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_NAV_FORWARD)),
        instance_,
        nullptr);
    if (forward_button_hwnd_ == nullptr) {
        LogLastError(L"CreateWindowExW(NavForward)");
        return false;
    }

    up_button_hwnd_ = CreateWindowExW(
        0,
        WC_BUTTONW,
        L"Up",
        button_style,
        0,
        0,
        0,
        0,
        hwnd_,
        reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_NAV_UP)),
        instance_,
        nullptr);
    if (up_button_hwnd_ == nullptr) {
        LogLastError(L"CreateWindowExW(NavUp)");
        return false;
    }

    refresh_button_hwnd_ = CreateWindowExW(
        0,
        WC_BUTTONW,
        L"R",
        button_style,
        0,
        0,
        0,
        0,
        hwnd_,
        reinterpret_cast<HMENU>(static_cast<INT_PTR>(ID_NAV_REFRESH)),
        instance_,
        nullptr);
    if (refresh_button_hwnd_ == nullptr) {
        LogLastError(L"CreateWindowExW(NavRefresh)");
        return false;
    }

    if (!breadcrumb_bar_.Create(hwnd_, instance_, 2100)) {
        return false;
    }
    breadcrumb_bar_.SetDpi(dpi_);

    return true;
}

void NavBar::LayoutChildControls() {
    if (hwnd_ == nullptr ||
        back_button_hwnd_ == nullptr ||
        forward_button_hwnd_ == nullptr ||
        up_button_hwnd_ == nullptr ||
        refresh_button_hwnd_ == nullptr ||
        breadcrumb_bar_.hwnd() == nullptr) {
        return;
    }

    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        LogLastError(L"GetClientRect(NavBar)");
        return;
    }

    const int width = client_rect.right - client_rect.left;
    const int height = client_rect.bottom - client_rect.top;
    if (width <= 0 || height <= 0) {
        return;
    }

    const int margin = MulDiv(6, static_cast<int>(dpi_), 96);
    const int gap = MulDiv(4, static_cast<int>(dpi_), 96);
    const int button_width = MulDiv(28, static_cast<int>(dpi_), 96);
    const int button_height = MulDiv(24, static_cast<int>(dpi_), 96);
    const int button_y = (std::max)(0, (height - button_height) / 2);

    int x = margin;
    if (!MoveWindow(back_button_hwnd_, x, button_y, button_width, button_height, TRUE)) {
        LogLastError(L"MoveWindow(NavBack)");
    }
    x += button_width + gap;

    if (!MoveWindow(forward_button_hwnd_, x, button_y, button_width, button_height, TRUE)) {
        LogLastError(L"MoveWindow(NavForward)");
    }
    x += button_width + gap;

    const int up_width = MulDiv(34, static_cast<int>(dpi_), 96);
    if (!MoveWindow(up_button_hwnd_, x, button_y, up_width, button_height, TRUE)) {
        LogLastError(L"MoveWindow(NavUp)");
    }
    x += up_width + gap;

    if (!MoveWindow(refresh_button_hwnd_, x, button_y, button_width, button_height, TRUE)) {
        LogLastError(L"MoveWindow(NavRefresh)");
    }
    x += button_width + gap;

    const int breadcrumb_width = (std::max)(0, width - x - margin);
    if (!MoveWindow(
            breadcrumb_bar_.hwnd(),
            x,
            margin / 2,
            breadcrumb_width,
            (std::max)(1, height - margin),
            TRUE)) {
        LogLastError(L"MoveWindow(BreadcrumbBar)");
    }
}

void NavBar::PostNavCommand(NavCommand command) const {
    if (parent_hwnd_ == nullptr) {
        return;
    }

    if (!PostMessageW(parent_hwnd_, WM_FE_NAV_COMMAND, static_cast<WPARAM>(command), 0)) {
        LogLastError(L"PostMessageW(WM_FE_NAV_COMMAND)");
    }
}

LRESULT CALLBACK NavBar::WindowProc(HWND hwnd, UINT message, WPARAM w_param, LPARAM l_param) {
    NavBar* self = nullptr;
    if (message == WM_NCCREATE) {
        auto* create_struct = reinterpret_cast<CREATESTRUCTW*>(l_param);
        self = reinterpret_cast<NavBar*>(create_struct->lpCreateParams);
        if (self == nullptr) {
            return FALSE;
        }
        self->hwnd_ = hwnd;
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    } else {
        self = reinterpret_cast<NavBar*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    }

    if (self != nullptr) {
        return self->HandleMessage(message, w_param, l_param);
    }
    return DefWindowProcW(hwnd, message, w_param, l_param);
}

LRESULT NavBar::HandleMessage(UINT message, WPARAM w_param, LPARAM l_param) {
    switch (message) {
    case WM_CREATE:
        if (!CreateChildControls()) {
            return -1;
        }
        LayoutChildControls();
        return 0;

    case WM_SIZE:
        LayoutChildControls();
        return 0;

    case WM_ERASEBKGND: {
        RECT client_rect = {};
        if (!GetClientRect(hwnd_, &client_rect)) {
            LogLastError(L"GetClientRect(WM_ERASEBKGND NavBar)");
            return 1;
        }

        HBRUSH brush = CreateSolidBrush(kNavBarBackground);
        if (brush != nullptr) {
            FillRect(reinterpret_cast<HDC>(w_param), &client_rect, brush);
            DeleteObject(brush);
        }
        return 1;
    }

    case WM_COMMAND:
        if (HIWORD(w_param) == BN_CLICKED) {
            switch (LOWORD(w_param)) {
            case ID_NAV_BACK:
                PostNavCommand(NavCommand::Back);
                return 0;
            case ID_NAV_FORWARD:
                PostNavCommand(NavCommand::Forward);
                return 0;
            case ID_NAV_UP:
                PostNavCommand(NavCommand::Up);
                return 0;
            case ID_NAV_REFRESH:
                PostNavCommand(NavCommand::Refresh);
                return 0;
            default:
                break;
            }
        }
        break;

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
