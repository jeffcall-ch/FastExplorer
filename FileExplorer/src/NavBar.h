#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>

#include <string>

#include "BreadcrumbBar.h"

namespace fileexplorer {

constexpr UINT WM_FE_NAV_COMMAND = WM_APP + 101;

enum class NavCommand : WPARAM {
    Back = 1,
    Forward = 2,
    Up = 3,
    Refresh = 4,
};

constexpr int ID_NAV_BACK = 2001;
constexpr int ID_NAV_FORWARD = 2002;
constexpr int ID_NAV_UP = 2003;
constexpr int ID_NAV_REFRESH = 2004;

class NavBar final {
public:
    NavBar();
    ~NavBar();

    NavBar(const NavBar&) = delete;
    NavBar& operator=(const NavBar&) = delete;

    bool Create(HWND parent, HINSTANCE instance, int control_id);
    HWND hwnd() const noexcept;

    void SetDpi(UINT dpi);
    void SetPath(const std::wstring& path);
    void SetNavigationState(bool can_back, bool can_forward, bool can_up);
    void ActivateAddressEditMode();
    void Refresh();

private:
    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM w_param, LPARAM l_param);
    LRESULT HandleMessage(UINT message, WPARAM w_param, LPARAM l_param);

    bool RegisterWindowClass();
    bool CreateChildControls();
    void LayoutChildControls();
    void PostNavCommand(NavCommand command) const;

    HWND parent_hwnd_{nullptr};
    HWND hwnd_{nullptr};
    HWND back_button_hwnd_{nullptr};
    HWND forward_button_hwnd_{nullptr};
    HWND up_button_hwnd_{nullptr};
    HWND refresh_button_hwnd_{nullptr};
    HINSTANCE instance_{nullptr};
    int control_id_{0};
    UINT dpi_{96U};

    BreadcrumbBar breadcrumb_bar_{};
};

}  // namespace fileexplorer

