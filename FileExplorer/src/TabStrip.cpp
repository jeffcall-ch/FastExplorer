#include "TabStrip.h"

#include <CommCtrl.h>
#include <Windowsx.h>
#include <strsafe.h>

#include <algorithm>
#include <cwchar>
#include <cstdlib>

namespace {

constexpr wchar_t kTabStripClassName[] = L"FE_TabStrip";
constexpr UINT_PTR kTooltipToolId = 1;

constexpr UINT kMenuCloseTab = 3001;
constexpr UINT kMenuCloseOthers = 3002;
constexpr UINT kMenuCloseRight = 3003;
constexpr UINT kMenuDuplicate = 3004;
constexpr UINT kMenuPinToggle = 3005;

constexpr COLORREF kStripBackground = RGB(0x15, 0x19, 0x1F);
constexpr COLORREF kStripBottomBorder = RGB(0x2D, 0x33, 0x3D);
constexpr COLORREF kTabActive = RGB(0x2A, 0x31, 0x3A);
constexpr COLORREF kTabHover = RGB(0x22, 0x28, 0x31);
constexpr COLORREF kTabInactive = RGB(0x1C, 0x22, 0x2A);
constexpr COLORREF kTabInactiveSeparator = RGB(0x2B, 0x32, 0x3B);
constexpr COLORREF kTabActiveTopAccent = RGB(0x67, 0x78, 0x8A);
constexpr COLORREF kTextActive = RGB(0xE8, 0xEE, 0xF6);
constexpr COLORREF kTextInactive = RGB(0xBD, 0xC6, 0xD1);
constexpr COLORREF kTextMuted = RGB(0x92, 0x9B, 0xA7);
constexpr COLORREF kButtonHover = RGB(0x2C, 0x34, 0x3E);
constexpr COLORREF kButtonInactive = RGB(0x21, 0x27, 0x30);
constexpr COLORREF kButtonPressed = RGB(0x34, 0x3D, 0x49);
constexpr COLORREF kButtonBorder = RGB(0x35, 0x3E, 0x49);
constexpr COLORREF kButtonBorderMuted = RGB(0x28, 0x2F, 0x38);

void FillRoundedRect(HDC hdc, const RECT& rect, COLORREF fill, COLORREF border, int radius) {
    HBRUSH brush = CreateSolidBrush(fill);
    HPEN pen = CreatePen(PS_SOLID, 1, border);
    if (brush == nullptr || pen == nullptr) {
        if (brush != nullptr) {
            DeleteObject(brush);
        }
        if (pen != nullptr) {
            DeleteObject(pen);
        }
        return;
    }

    HGDIOBJ old_brush = SelectObject(hdc, brush);
    HGDIOBJ old_pen = SelectObject(hdc, pen);
    RoundRect(hdc, rect.left, rect.top, rect.right, rect.bottom, radius, radius);
    SelectObject(hdc, old_pen);
    SelectObject(hdc, old_brush);
    DeleteObject(pen);
    DeleteObject(brush);
}

void LogLastError(const wchar_t* context) {
    const DWORD error_code = GetLastError();
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (GetLastError=%lu).\r\n", context, error_code);
    OutputDebugStringW(buffer);
}

bool RectContainsPoint(const RECT& rect, POINT point) {
    return point.x >= rect.left && point.x < rect.right && point.y >= rect.top && point.y < rect.bottom;
}

bool IsRectValid(const RECT& rect) {
    return rect.right > rect.left && rect.bottom > rect.top;
}

RECT MakeEmptyRect() {
    RECT rect = {0, 0, 0, 0};
    return rect;
}

}  // namespace

namespace fileexplorer {

void TabStrip::FontDeleter::operator()(HFONT font) const noexcept {
    if (font != nullptr) {
        DeleteObject(font);
    }
}

TabStrip::TabStrip() = default;

TabStrip::~TabStrip() {
    if (tooltip_hwnd_ != nullptr) {
        DestroyWindow(tooltip_hwnd_);
        tooltip_hwnd_ = nullptr;
    }
}

bool TabStrip::Create(HWND parent, HINSTANCE instance, int control_id) {
    parent_hwnd_ = parent;
    instance_ = instance;
    control_id_ = control_id;

    if (!RegisterWindowClass()) {
        return false;
    }

    hwnd_ = CreateWindowExW(
        0,
        kTabStripClassName,
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
        LogLastError(L"CreateWindowExW(TabStrip)");
        return false;
    }

    if (!CreateTooltip()) {
        return false;
    }

    return true;
}

HWND TabStrip::hwnd() const noexcept {
    return hwnd_;
}

void TabStrip::SetTabManager(TabManager* manager) {
    tab_manager_ = manager;
    Refresh();
}

void TabStrip::SetDpi(UINT dpi) {
    dpi_ = (dpi == 0U) ? 96U : dpi;
    font_.reset();
    font_pixel_height_ = 0;
    Refresh();
}

void TabStrip::Refresh() {
    if (hwnd_ != nullptr) {
        InvalidateRect(hwnd_, nullptr, TRUE);
    }
}

bool TabStrip::IsDraggableGap(POINT parent_client_point) const {
    if (hwnd_ == nullptr || parent_hwnd_ == nullptr) {
        return false;
    }

    POINT tab_point = parent_client_point;
    if (MapWindowPoints(parent_hwnd_, hwnd_, &tab_point, 1) == 0) {
        return false;
    }

    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        return false;
    }
    if (!RectContainsPoint(client_rect, tab_point)) {
        return false;
    }

    const HitTestResult hit = HitTest(tab_point);
    return hit.part == HitPart::None;
}

bool TabStrip::RegisterWindowClass() {
    WNDCLASSEXW window_class = {};
    window_class.cbSize = sizeof(window_class);
    window_class.style = CS_HREDRAW | CS_VREDRAW;
    window_class.lpfnWndProc = &TabStrip::WindowProc;
    window_class.hInstance = instance_;
    window_class.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    window_class.hbrBackground = nullptr;
    window_class.lpszClassName = kTabStripClassName;

    const ATOM class_id = RegisterClassExW(&window_class);
    if (class_id == 0 && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
        LogLastError(L"RegisterClassExW(FE_TabStrip)");
        return false;
    }

    return true;
}

bool TabStrip::CreateTooltip() {
    tooltip_hwnd_ = CreateWindowExW(
        WS_EX_TOPMOST,
        TOOLTIPS_CLASSW,
        nullptr,
        WS_POPUP | TTS_ALWAYSTIP | TTS_NOPREFIX,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        hwnd_,
        nullptr,
        instance_,
        nullptr);
    if (tooltip_hwnd_ == nullptr) {
        LogLastError(L"CreateWindowExW(Tooltip)");
        return false;
    }

    if (!SetWindowPos(
            tooltip_hwnd_,
            HWND_TOPMOST,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE)) {
        LogLastError(L"SetWindowPos(Tooltip)");
    }

    TOOLINFOW tool = {};
    tool.cbSize = sizeof(tool);
    tool.uFlags = TTF_SUBCLASS;
    tool.hwnd = hwnd_;
    tool.uId = kTooltipToolId;
    tool.lpszText = LPSTR_TEXTCALLBACKW;
    GetClientRect(hwnd_, &tool.rect);

    if (!SendMessageW(tooltip_hwnd_, TTM_ADDTOOLW, 0, reinterpret_cast<LPARAM>(&tool))) {
        LogLastError(L"TTM_ADDTOOLW");
        return false;
    }

    SendMessageW(tooltip_hwnd_, TTM_SETMAXTIPWIDTH, 0, 800);
    SendMessageW(tooltip_hwnd_, TTM_ACTIVATE, FALSE, 0);
    return true;
}

void TabStrip::UpdateTooltipRegion() {
    if (tooltip_hwnd_ == nullptr || hwnd_ == nullptr) {
        return;
    }

    TOOLINFOW tool = {};
    tool.cbSize = sizeof(tool);
    tool.hwnd = hwnd_;
    tool.uId = kTooltipToolId;
    GetClientRect(hwnd_, &tool.rect);
    SendMessageW(tooltip_hwnd_, TTM_NEWTOOLRECTW, 0, reinterpret_cast<LPARAM>(&tool));
}

void TabStrip::RebuildLayout(HDC hdc) {
    visible_tabs_.clear();
    plus_button_rect_ = MakeEmptyRect();
    left_scroll_rect_ = MakeEmptyRect();
    right_scroll_rect_ = MakeEmptyRect();
    show_scroll_buttons_ = false;
    left_scroll_enabled_ = false;
    right_scroll_enabled_ = false;

    if (tab_manager_ == nullptr || hdc == nullptr) {
        return;
    }

    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        return;
    }

    const int client_width = client_rect.right - client_rect.left;
    const int client_height = client_rect.bottom - client_rect.top;
    if (client_width <= 0 || client_height <= 0) {
        return;
    }

    const auto& tabs = tab_manager_->tabs();
    const int tab_count = static_cast<int>(tabs.size());

    const int margin = MulDiv(8, static_cast<int>(dpi_), 96);
    const int tab_gap = MulDiv(2, static_cast<int>(dpi_), 96);
    const int plus_width = MulDiv(28, static_cast<int>(dpi_), 96);
    const int arrow_width = MulDiv(24, static_cast<int>(dpi_), 96);
    const int tab_height = std::max(MulDiv(28, static_cast<int>(dpi_), 96), client_height - MulDiv(8, static_cast<int>(dpi_), 96));
    const int tab_top = std::max(0, (client_height - tab_height) / 2);
    const int close_size = MulDiv(12, static_cast<int>(dpi_), 96);
    const int pinned_min_width = MulDiv(96, static_cast<int>(dpi_), 96);
    const int pinned_max_width = MulDiv(180, static_cast<int>(dpi_), 96);
    const int tab_min_width = MulDiv(120, static_cast<int>(dpi_), 96);
    const int tab_max_width = MulDiv(230, static_cast<int>(dpi_), 96);
    const int pinned_prefix_width = MeasureTextWidth(hdc, L"[P] ");

    std::vector<int> tab_widths;
    tab_widths.reserve(tab_count);
    int total_tabs_width = 0;
    for (const auto& tab : tabs) {
        int width = 0;
        if (tab.pinned) {
            const int text_width = MeasureTextWidth(hdc, tab.displayName);
            width = std::clamp(
                pinned_prefix_width + text_width + MulDiv(28, static_cast<int>(dpi_), 96),
                pinned_min_width,
                pinned_max_width);
        } else {
            const int text_width = MeasureTextWidth(hdc, tab.displayName);
            width = std::clamp(text_width + MulDiv(36, static_cast<int>(dpi_), 96) + close_size, tab_min_width, tab_max_width);
        }
        tab_widths.push_back(width);
        total_tabs_width += width;
    }
    if (tab_count > 1) {
        total_tabs_width += (tab_count - 1) * tab_gap;
    }

    const int available_without_scroll = std::max(0, client_width - (2 * margin) - plus_width - tab_gap);
    show_scroll_buttons_ = total_tabs_width > available_without_scroll;

    int start_x = margin;
    int end_x = std::max(start_x, client_width - margin - plus_width - tab_gap);

    if (show_scroll_buttons_) {
        left_scroll_rect_ = {margin, tab_top, margin + arrow_width, tab_top + tab_height};
        right_scroll_rect_ = {
            std::max(margin, client_width - margin - plus_width - tab_gap - arrow_width),
            tab_top,
            std::max(margin, client_width - margin - plus_width - tab_gap),
            tab_top + tab_height,
        };

        start_x = left_scroll_rect_.right + tab_gap;
        end_x = right_scroll_rect_.left - tab_gap;
        if (end_x < start_x) {
            end_x = start_x;
        }
    }

    if (tab_count > 0) {
        first_visible_index_ = std::clamp(first_visible_index_, 0, tab_count - 1);
    } else {
        first_visible_index_ = 0;
    }

    for (int attempt = 0; attempt < 4; ++attempt) {
        visible_tabs_.clear();
        int x = start_x;

        for (int i = first_visible_index_; i < tab_count; ++i) {
            const int width = tab_widths[i];
            if (x + width > end_x) {
                break;
            }

            VisibleTab visible = {};
            visible.tabIndex = i;
            visible.tabRect = {x, tab_top, x + width, tab_top + tab_height};
            visible.closeRect = MakeEmptyRect();

            if (!tabs[i].pinned) {
                const int close_left = visible.tabRect.right - close_size - MulDiv(10, static_cast<int>(dpi_), 96);
                const int close_top = visible.tabRect.top + (tab_height - close_size) / 2;
                visible.closeRect = {close_left, close_top, close_left + close_size, close_top + close_size};
            }

            visible_tabs_.push_back(visible);
            x += width + tab_gap;
        }

        if (!show_scroll_buttons_ || tab_count == 0) {
            break;
        }

        const int active = tab_manager_->active_index();
        if (active < first_visible_index_) {
            first_visible_index_ = active;
            continue;
        }
        if (!visible_tabs_.empty() && active > visible_tabs_.back().tabIndex) {
            first_visible_index_ = active;
            continue;
        }
        if (visible_tabs_.empty() && first_visible_index_ > 0) {
            --first_visible_index_;
            continue;
        }
        break;
    }

    plus_button_rect_ = {
        std::max(margin, client_width - margin - plus_width),
        tab_top,
        std::max(margin, client_width - margin),
        tab_top + tab_height,
    };

    if (show_scroll_buttons_) {
        left_scroll_enabled_ = first_visible_index_ > 0;
        right_scroll_enabled_ =
            !visible_tabs_.empty() ? (visible_tabs_.back().tabIndex < (tab_count - 1)) : (first_visible_index_ < (tab_count - 1));
    }
}

void TabStrip::HandleMouseMove(POINT point) {
    if (drag_candidate_tab_index_ >= 0) {
        HandleTabDragMove(point);
        return;
    }

    TRACKMOUSEEVENT track = {};
    track.cbSize = sizeof(track);
    track.dwFlags = TME_LEAVE;
    track.hwndTrack = hwnd_;
    TrackMouseEvent(&track);

    const HitTestResult hit = HitTest(point);

    const int new_hover_tab =
        (hit.part == HitPart::TabBody || hit.part == HitPart::CloseButton) ? hit.tabIndex : -1;
    const int new_hover_close = (hit.part == HitPart::CloseButton) ? hit.tabIndex : -1;
    const bool new_hover_plus = hit.part == HitPart::PlusButton;
    const bool new_hover_left = hit.part == HitPart::LeftScroll;
    const bool new_hover_right = hit.part == HitPart::RightScroll;

    const bool changed =
        hovered_tab_index_ != new_hover_tab ||
        hovered_close_tab_index_ != new_hover_close ||
        hovered_plus_button_ != new_hover_plus ||
        hovered_left_scroll_ != new_hover_left ||
        hovered_right_scroll_ != new_hover_right;

    hovered_tab_index_ = new_hover_tab;
    hovered_close_tab_index_ = new_hover_close;
    hovered_plus_button_ = new_hover_plus;
    hovered_left_scroll_ = new_hover_left;
    hovered_right_scroll_ = new_hover_right;
    tooltip_tab_index_ = new_hover_tab;

    if (tooltip_hwnd_ != nullptr) {
        SendMessageW(tooltip_hwnd_, TTM_ACTIVATE, tooltip_tab_index_ >= 0 ? TRUE : FALSE, 0);
        if (tooltip_tab_index_ < 0) {
            SendMessageW(tooltip_hwnd_, TTM_POP, 0, 0);
        }
    }

    if (changed) {
        Refresh();
    }
}

void TabStrip::HandleMouseLeave() {
    if (drag_candidate_tab_index_ >= 0) {
        return;
    }

    hovered_tab_index_ = -1;
    hovered_close_tab_index_ = -1;
    hovered_plus_button_ = false;
    hovered_left_scroll_ = false;
    hovered_right_scroll_ = false;
    tooltip_tab_index_ = -1;
    if (tooltip_hwnd_ != nullptr) {
        SendMessageW(tooltip_hwnd_, TTM_ACTIVATE, FALSE, 0);
    }
    Refresh();
}

void TabStrip::HandleLeftButtonDown(POINT point) {
    if (tab_manager_ == nullptr) {
        return;
    }

    const HitTestResult hit = HitTest(point);
    bool tabs_changed = false;
    bool redraw_only = false;

    switch (hit.part) {
    case HitPart::CloseButton:
        tabs_changed = tab_manager_->CloseTab(hit.tabIndex);
        break;

    case HitPart::PlusButton:
        tabs_changed = tab_manager_->AddTab(DefaultNewTabPath(), true, false) >= 0;
        break;

    case HitPart::TabBody:
        if (tab_manager_->active_index() != hit.tabIndex) {
            tabs_changed = tab_manager_->Activate(hit.tabIndex);
        }
        drag_candidate_tab_index_ = hit.tabIndex;
        drag_tab_index_ = hit.tabIndex;
        drag_start_point_ = point;
        tab_dragging_ = false;
        tab_reordered_during_drag_ = false;
        SetCapture(hwnd_);
        break;

    case HitPart::LeftScroll:
        if (show_scroll_buttons_ && left_scroll_enabled_ && first_visible_index_ > 0) {
            --first_visible_index_;
            redraw_only = true;
        }
        break;

    case HitPart::RightScroll:
        if (show_scroll_buttons_ && right_scroll_enabled_) {
            ++first_visible_index_;
            redraw_only = true;
        }
        break;

    case HitPart::None:
        ResetTabDragState(true);
        break;
    }

    if (tabs_changed) {
        NotifyParentTabsChanged();
    }
    if (tabs_changed || redraw_only) {
        Refresh();
    }
}

void TabStrip::HandleLeftButtonUp(POINT point) {
    if (drag_candidate_tab_index_ >= 0) {
        HandleTabDragMove(point);
        if (tab_reordered_during_drag_) {
            Refresh();
        }
    }

    ResetTabDragState(true);
}

void TabStrip::HandleTabDragMove(POINT point) {
    if (tab_manager_ == nullptr || drag_candidate_tab_index_ < 0 || drag_tab_index_ < 0) {
        return;
    }

    if (!tab_dragging_) {
        const int drag_threshold_x = (std::max)(1, GetSystemMetrics(SM_CXDRAG));
        const int drag_threshold_y = (std::max)(1, GetSystemMetrics(SM_CYDRAG));
        const int delta_x = std::abs(point.x - drag_start_point_.x);
        const int delta_y = std::abs(point.y - drag_start_point_.y);
        if (delta_x < drag_threshold_x && delta_y < drag_threshold_y) {
            return;
        }
        tab_dragging_ = true;
    }

    const int target_index = TabIndexForDragPoint(point);
    if (target_index < 0 || target_index == drag_tab_index_) {
        return;
    }

    if (tab_manager_->MoveTab(drag_tab_index_, target_index)) {
        drag_tab_index_ = target_index;
        tab_reordered_during_drag_ = true;
        hovered_tab_index_ = drag_tab_index_;
        hovered_close_tab_index_ = -1;
        Refresh();
    }
}

void TabStrip::ResetTabDragState(bool release_capture) {
    if (release_capture && hwnd_ != nullptr && GetCapture() == hwnd_) {
        ReleaseCapture();
    }
    drag_candidate_tab_index_ = -1;
    drag_tab_index_ = -1;
    tab_dragging_ = false;
    tab_reordered_during_drag_ = false;
    drag_start_point_ = POINT{0, 0};
}

int TabStrip::TabIndexForDragPoint(POINT point) const {
    if (tab_manager_ == nullptr) {
        return -1;
    }
    if (visible_tabs_.empty()) {
        return -1;
    }

    for (const VisibleTab& tab : visible_tabs_) {
        if (RectContainsPoint(tab.tabRect, point)) {
            const int center_x = tab.tabRect.left + ((tab.tabRect.right - tab.tabRect.left) / 2);
            if (point.x >= center_x && tab.tabIndex < static_cast<int>(tab_manager_->tabs().size()) - 1) {
                return tab.tabIndex + 1;
            }
            return tab.tabIndex;
        }
    }

    if (point.x < visible_tabs_.front().tabRect.left) {
        return visible_tabs_.front().tabIndex;
    }
    if (point.x >= visible_tabs_.back().tabRect.right) {
        return visible_tabs_.back().tabIndex;
    }

    return -1;
}

void TabStrip::HandleRightClick(POINT point_screen, POINT point_client) {
    const HitTestResult hit = HitTest(point_client);
    if (hit.part == HitPart::TabBody || hit.part == HitPart::CloseButton) {
        ShowTabContextMenu(hit.tabIndex, point_screen);
    }
}

void TabStrip::ShowTabContextMenu(int tab_index, POINT point_screen) {
    if (tab_manager_ == nullptr) {
        return;
    }
    const auto& tabs = tab_manager_->tabs();
    if (tab_index < 0 || tab_index >= static_cast<int>(tabs.size())) {
        return;
    }

    HMENU menu = CreatePopupMenu();
    if (menu == nullptr) {
        LogLastError(L"CreatePopupMenu(TabStrip)");
        return;
    }

    AppendMenuW(menu, MF_STRING, kMenuCloseTab, L"Close tab");
    AppendMenuW(menu, MF_STRING, kMenuCloseOthers, L"Close other tabs");
    AppendMenuW(menu, MF_STRING, kMenuCloseRight, L"Close tabs to the right");
    AppendMenuW(menu, MF_STRING, kMenuDuplicate, L"Duplicate tab");
    AppendMenuW(menu, MF_STRING, kMenuPinToggle, tabs[tab_index].pinned ? L"Unpin tab" : L"Pin tab");

    if (!tab_manager_->CanCloseTab(tab_index)) {
        EnableMenuItem(menu, kMenuCloseTab, MF_BYCOMMAND | MF_GRAYED);
    }

    UINT command =
        static_cast<UINT>(TrackPopupMenuEx(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, point_screen.x, point_screen.y, hwnd_, nullptr));

    bool changed = false;
    switch (command) {
    case kMenuCloseTab:
        changed = tab_manager_->CloseTab(tab_index);
        break;
    case kMenuCloseOthers:
        changed = tab_manager_->CloseOtherTabs(tab_index);
        break;
    case kMenuCloseRight:
        changed = tab_manager_->CloseTabsToRight(tab_index);
        break;
    case kMenuDuplicate:
        changed = tab_manager_->DuplicateTab(tab_index);
        break;
    case kMenuPinToggle:
        changed = tab_manager_->TogglePin(tab_index);
        break;
    default:
        break;
    }

    DestroyMenu(menu);

    if (changed) {
        NotifyParentTabsChanged();
        Refresh();
    }
}

void TabStrip::NotifyParentTabsChanged() const {
    if (parent_hwnd_ != nullptr) {
        PostMessageW(parent_hwnd_, WM_FE_TAB_CHANGED, 0, 0);
    }
}

TabStrip::HitTestResult TabStrip::HitTest(POINT point) const {
    for (const auto& visible : visible_tabs_) {
        if (IsRectValid(visible.closeRect) && RectContainsPoint(visible.closeRect, point)) {
            return {HitPart::CloseButton, visible.tabIndex};
        }
    }

    for (const auto& visible : visible_tabs_) {
        if (RectContainsPoint(visible.tabRect, point)) {
            return {HitPart::TabBody, visible.tabIndex};
        }
    }

    if (RectContainsPoint(plus_button_rect_, point)) {
        return {HitPart::PlusButton, -1};
    }
    if (show_scroll_buttons_ && RectContainsPoint(left_scroll_rect_, point)) {
        return {HitPart::LeftScroll, -1};
    }
    if (show_scroll_buttons_ && RectContainsPoint(right_scroll_rect_, point)) {
        return {HitPart::RightScroll, -1};
    }

    return {};
}

void TabStrip::DrawTabStrip(HDC hdc) {
    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        return;
    }

    HBRUSH background = CreateSolidBrush(kStripBackground);
    if (background != nullptr) {
        FillRect(hdc, &client_rect, background);
        DeleteObject(background);
    }

    HPEN strip_border_pen = CreatePen(PS_SOLID, 1, kStripBottomBorder);
    if (strip_border_pen != nullptr) {
        HGDIOBJ old_pen = SelectObject(hdc, strip_border_pen);
        MoveToEx(hdc, client_rect.left, client_rect.bottom - 1, nullptr);
        LineTo(hdc, client_rect.right, client_rect.bottom - 1);
        SelectObject(hdc, old_pen);
        DeleteObject(strip_border_pen);
    }

    const int requested_pixel_height = -MulDiv(9, static_cast<int>(dpi_), 72);
    if (font_ == nullptr || font_pixel_height_ != requested_pixel_height) {
        font_.reset(CreateFontW(
            requested_pixel_height,
            0,
            0,
            0,
            FW_NORMAL,
            FALSE,
            FALSE,
            FALSE,
            DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS,
            CLIP_DEFAULT_PRECIS,
            CLEARTYPE_QUALITY,
            DEFAULT_PITCH | FF_DONTCARE,
            L"Segoe UI Variable Text"));
        font_pixel_height_ = requested_pixel_height;
    }

    HFONT previous_font = nullptr;
    if (font_ != nullptr) {
        previous_font = static_cast<HFONT>(SelectObject(hdc, font_.get()));
    }

    RebuildLayout(hdc);

    SetBkMode(hdc, TRANSPARENT);

    const int corner_radius = MulDiv(7, static_cast<int>(dpi_), 96);
    const auto draw_compact_button = [&](const RECT& rect, bool enabled, bool hovered, const wchar_t* symbol) {
        const bool pressed = hovered && (GetKeyState(VK_LBUTTON) & 0x8000) != 0;
        COLORREF fill = enabled ? (pressed ? kButtonPressed : (hovered ? kButtonHover : kButtonInactive)) : kButtonInactive;
        COLORREF border = enabled ? kButtonBorder : kButtonBorderMuted;
        FillRoundedRect(hdc, rect, fill, border, corner_radius);
        SetTextColor(hdc, enabled ? kTextInactive : kTextMuted);
        RECT text_rect = rect;
        DrawTextW(hdc, symbol, -1, &text_rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
    };

    if (show_scroll_buttons_) {
        draw_compact_button(left_scroll_rect_, left_scroll_enabled_, hovered_left_scroll_, L"\u2039");
        draw_compact_button(right_scroll_rect_, right_scroll_enabled_, hovered_right_scroll_, L"\u203A");
    }

    const int active_index = tab_manager_ != nullptr ? tab_manager_->active_index() : -1;
    const std::vector<TabState>* tabs = (tab_manager_ != nullptr) ? &tab_manager_->tabs() : nullptr;

    for (const auto& visible : visible_tabs_) {
        const bool is_active = visible.tabIndex == active_index;
        const bool is_hovered = visible.tabIndex == hovered_tab_index_;
        const bool show_close =
            tabs != nullptr &&
            !(*tabs)[visible.tabIndex].pinned &&
            (is_active || is_hovered);

        COLORREF fill = kTabInactive;
        if (is_active) {
            fill = kTabActive;
        } else if (is_hovered) {
            fill = kTabHover;
        }

        FillRoundedRect(hdc, visible.tabRect, fill, fill, corner_radius);

        if (!is_active) {
            HPEN separator_pen = CreatePen(PS_SOLID, 1, kTabInactiveSeparator);
            if (separator_pen != nullptr) {
                HGDIOBJ old_pen = SelectObject(hdc, separator_pen);
                MoveToEx(
                    hdc,
                    visible.tabRect.right - 1,
                    visible.tabRect.top + MulDiv(6, static_cast<int>(dpi_), 96),
                    nullptr);
                LineTo(
                    hdc,
                    visible.tabRect.right - 1,
                    visible.tabRect.bottom - MulDiv(6, static_cast<int>(dpi_), 96));
                SelectObject(hdc, old_pen);
                DeleteObject(separator_pen);
            }
        }

        if (is_active) {
            HPEN highlight_pen = CreatePen(PS_SOLID, 1, kTabActiveTopAccent);
            if (highlight_pen != nullptr) {
                HGDIOBJ old_pen = SelectObject(hdc, highlight_pen);
                MoveToEx(
                    hdc,
                    visible.tabRect.left + MulDiv(10, static_cast<int>(dpi_), 96),
                    visible.tabRect.top + MulDiv(1, static_cast<int>(dpi_), 96),
                    nullptr);
                LineTo(
                    hdc,
                    visible.tabRect.right - MulDiv(10, static_cast<int>(dpi_), 96),
                    visible.tabRect.top + MulDiv(1, static_cast<int>(dpi_), 96));
                SelectObject(hdc, old_pen);
                DeleteObject(highlight_pen);
            }

            RECT join_rect = visible.tabRect;
            join_rect.top = join_rect.bottom - 1;
            HBRUSH join_brush = CreateSolidBrush(fill);
            if (join_brush != nullptr) {
                FillRect(hdc, &join_rect, join_brush);
                DeleteObject(join_brush);
            }
        }

        RECT text_rect = visible.tabRect;
        text_rect.left += MulDiv(10, static_cast<int>(dpi_), 96);
        if (show_close && IsRectValid(visible.closeRect)) {
            text_rect.right = visible.closeRect.left - MulDiv(6, static_cast<int>(dpi_), 96);
        } else {
            text_rect.right -= MulDiv(8, static_cast<int>(dpi_), 96);
        }

        SetTextColor(hdc, is_active ? kTextActive : kTextInactive);
        if (tabs != nullptr && (*tabs)[visible.tabIndex].pinned) {
            const std::wstring pinned_label = L"[P] " + (*tabs)[visible.tabIndex].displayName;
            DrawTextW(
                hdc,
                pinned_label.c_str(),
                -1,
                &text_rect,
                DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
        } else {
            DrawTextW(
                hdc,
                (tabs != nullptr ? (*tabs)[visible.tabIndex].displayName.c_str() : L""),
                -1,
                &text_rect,
                DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
        }

        if (show_close && IsRectValid(visible.closeRect)) {
            const bool close_hovered = hovered_close_tab_index_ == visible.tabIndex;
            if (close_hovered) {
                FillRoundedRect(
                    hdc,
                    visible.closeRect,
                    kButtonHover,
                    kButtonHover,
                    MulDiv(4, static_cast<int>(dpi_), 96));
            }
            SetTextColor(hdc, close_hovered ? kTextActive : kTextMuted);
            RECT close_rect = visible.closeRect;
            DrawTextW(
                hdc,
                L"\u00D7",
                -1,
                &close_rect,
                DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
        }
    }

    draw_compact_button(plus_button_rect_, true, hovered_plus_button_, L"+");

    if (previous_font != nullptr) {
        SelectObject(hdc, previous_font);
    }
}

int TabStrip::MeasureTextWidth(HDC hdc, const std::wstring& text) const {
    if (text.empty()) {
        return 0;
    }
    SIZE size = {};
    if (!GetTextExtentPoint32W(hdc, text.c_str(), static_cast<int>(text.size()), &size)) {
        LogLastError(L"GetTextExtentPoint32W");
        return 0;
    }
    return size.cx;
}

std::wstring TabStrip::TooltipTextForHoveredTab() const {
    if (tab_manager_ == nullptr || tooltip_tab_index_ < 0) {
        return L"";
    }

    const auto& tabs = tab_manager_->tabs();
    if (tooltip_tab_index_ >= static_cast<int>(tabs.size())) {
        return L"";
    }

    return tabs[tooltip_tab_index_].path;
}

std::wstring TabStrip::DefaultNewTabPath() const {
    return TabManager::DefaultTabPath();
}

LRESULT CALLBACK TabStrip::WindowProc(HWND hwnd, UINT message, WPARAM w_param, LPARAM l_param) {
    TabStrip* self = nullptr;
    if (message == WM_NCCREATE) {
        auto* create_struct = reinterpret_cast<CREATESTRUCTW*>(l_param);
        self = reinterpret_cast<TabStrip*>(create_struct->lpCreateParams);
        if (self == nullptr) {
            return FALSE;
        }
        self->hwnd_ = hwnd;
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    } else {
        self = reinterpret_cast<TabStrip*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    }

    if (self != nullptr) {
        return self->HandleMessage(message, w_param, l_param);
    }
    return DefWindowProcW(hwnd, message, w_param, l_param);
}

LRESULT TabStrip::HandleMessage(UINT message, WPARAM w_param, LPARAM l_param) {
    switch (message) {
    case WM_SIZE:
        UpdateTooltipRegion();
        Refresh();
        return 0;

    case WM_ERASEBKGND:
        return 1;

    case WM_PAINT: {
        PAINTSTRUCT paint = {};
        HDC hdc = BeginPaint(hwnd_, &paint);
        DrawTabStrip(hdc);
        EndPaint(hwnd_, &paint);
        return 0;
    }

    case WM_MOUSEMOVE: {
        POINT point = {GET_X_LPARAM(l_param), GET_Y_LPARAM(l_param)};
        HandleMouseMove(point);
        return 0;
    }

    case WM_MOUSELEAVE:
        HandleMouseLeave();
        return 0;

    case WM_LBUTTONDOWN: {
        SetFocus(parent_hwnd_);
        POINT point = {GET_X_LPARAM(l_param), GET_Y_LPARAM(l_param)};
        HandleLeftButtonDown(point);
        return 0;
    }

    case WM_LBUTTONUP: {
        POINT point = {GET_X_LPARAM(l_param), GET_Y_LPARAM(l_param)};
        HandleLeftButtonUp(point);
        return 0;
    }

    case WM_RBUTTONUP: {
        POINT client_point = {GET_X_LPARAM(l_param), GET_Y_LPARAM(l_param)};
        POINT screen_point = client_point;
        ClientToScreen(hwnd_, &screen_point);
        HandleRightClick(screen_point, client_point);
        return 0;
    }

    case WM_CONTEXTMENU: {
        POINT screen_point = {GET_X_LPARAM(l_param), GET_Y_LPARAM(l_param)};
        POINT client_point = {};
        if (screen_point.x == -1 && screen_point.y == -1) {
            if (tab_manager_ == nullptr || tab_manager_->active_index() < 0) {
                return 0;
            }
            int active_index = tab_manager_->active_index();
            auto it = std::find_if(
                visible_tabs_.begin(),
                visible_tabs_.end(),
                [active_index](const VisibleTab& visible) { return visible.tabIndex == active_index; });
            if (it == visible_tabs_.end()) {
                return 0;
            }
            client_point.x = (it->tabRect.left + it->tabRect.right) / 2;
            client_point.y = (it->tabRect.top + it->tabRect.bottom) / 2;
            screen_point = client_point;
            ClientToScreen(hwnd_, &screen_point);
        } else {
            client_point = screen_point;
            ScreenToClient(hwnd_, &client_point);
        }
        HandleRightClick(screen_point, client_point);
        return 0;
    }

    case WM_NOTIFY: {
        auto* header = reinterpret_cast<NMHDR*>(l_param);
        if (header != nullptr && header->hwndFrom == tooltip_hwnd_ && header->code == TTN_GETDISPINFOW) {
            auto* tooltip = reinterpret_cast<NMTTDISPINFOW*>(l_param);
            const std::wstring text = TooltipTextForHoveredTab();
            StringCchCopyW(tooltip_buffer_.data(), tooltip_buffer_.size(), text.c_str());
            tooltip->lpszText = tooltip_buffer_.data();
            return TRUE;
        }
        break;
    }

    case WM_CAPTURECHANGED:
        if (reinterpret_cast<HWND>(l_param) != hwnd_) {
            ResetTabDragState(false);
        }
        return 0;

    case WM_NCDESTROY:
        ResetTabDragState(false);
        SetWindowLongPtrW(hwnd_, GWLP_USERDATA, 0);
        hwnd_ = nullptr;
        break;

    default:
        break;
    }

    return DefWindowProcW(hwnd_, message, w_param, l_param);
}

}  // namespace fileexplorer
