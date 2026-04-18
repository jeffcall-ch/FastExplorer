#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>

#include <array>
#include <memory>
#include <string>
#include <type_traits>
#include <vector>

#include "TabManager.h"

namespace fileexplorer {

constexpr UINT WM_FE_TAB_CHANGED = WM_APP + 100;

class TabStrip final {
public:
    TabStrip();
    ~TabStrip();

    TabStrip(const TabStrip&) = delete;
    TabStrip& operator=(const TabStrip&) = delete;

    bool Create(HWND parent, HINSTANCE instance, int control_id);
    HWND hwnd() const noexcept;

    void SetTabManager(TabManager* manager);
    void SetDpi(UINT dpi);
    void Refresh();

    bool IsDraggableGap(POINT parent_client_point) const;

private:
    enum class HitPart {
        None,
        TabBody,
        CloseButton,
        PlusButton,
        LeftScroll,
        RightScroll,
    };

    struct HitTestResult {
        HitPart part = HitPart::None;
        int tabIndex = -1;
    };

    struct VisibleTab {
        int tabIndex = -1;
        RECT tabRect{};
        RECT closeRect{};
    };

    struct FontDeleter {
        void operator()(HFONT font) const noexcept;
    };

    using UniqueFont = std::unique_ptr<std::remove_pointer_t<HFONT>, FontDeleter>;

    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM w_param, LPARAM l_param);
    LRESULT HandleMessage(UINT message, WPARAM w_param, LPARAM l_param);

    bool RegisterWindowClass();
    bool CreateTooltip();
    void UpdateTooltipRegion();
    void RebuildLayout(HDC hdc);

    void HandleMouseMove(POINT point);
    void HandleMouseLeave();
    void HandleLeftButtonDown(POINT point);
    void HandleRightClick(POINT point_screen, POINT point_client);
    void ShowTabContextMenu(int tab_index, POINT point_screen);
    void NotifyParentTabsChanged() const;

    HitTestResult HitTest(POINT point) const;
    void DrawTabStrip(HDC hdc);
    int MeasureTextWidth(HDC hdc, const std::wstring& text) const;
    std::wstring TooltipTextForHoveredTab() const;
    std::wstring DefaultNewTabPath() const;

    HWND parent_hwnd_{nullptr};
    HWND hwnd_{nullptr};
    HWND tooltip_hwnd_{nullptr};
    HINSTANCE instance_{nullptr};
    int control_id_{0};
    UINT dpi_{96U};

    TabManager* tab_manager_{nullptr};

    std::vector<VisibleTab> visible_tabs_;
    RECT plus_button_rect_{};
    RECT left_scroll_rect_{};
    RECT right_scroll_rect_{};

    int first_visible_index_{0};
    bool show_scroll_buttons_{false};
    bool left_scroll_enabled_{false};
    bool right_scroll_enabled_{false};

    int hovered_tab_index_{-1};
    int hovered_close_tab_index_{-1};
    bool hovered_plus_button_{false};
    bool hovered_left_scroll_{false};
    bool hovered_right_scroll_{false};

    int tooltip_tab_index_{-1};
    std::array<wchar_t, 1024> tooltip_buffer_{};

    UniqueFont font_{nullptr};
    int font_pixel_height_{0};
};

}  // namespace fileexplorer
