#include "Sidebar.h"

#include <Shellapi.h>
#include <Windowsx.h>
#include <strsafe.h>

#include <algorithm>
#include <cwchar>
#include <cwctype>

namespace {

constexpr wchar_t kSidebarClassName[] = L"FE_Sidebar";

constexpr COLORREF kSidebarBackgroundColor = RGB(0x21, 0x21, 0x21);
constexpr COLORREF kSidebarBorderColor = RGB(0x33, 0x33, 0x33);
constexpr COLORREF kSearchBackgroundColor = RGB(0x2A, 0x2A, 0x2A);
constexpr COLORREF kSearchBorderColor = RGB(0x3A, 0x3A, 0x3A);
constexpr COLORREF kSearchTextColor = RGB(0xA8, 0xA8, 0xA8);
constexpr COLORREF kHeaderTextColor = RGB(0xD7, 0xD7, 0xD7);
constexpr COLORREF kItemTextColor = RGB(0xE6, 0xE6, 0xE6);
constexpr COLORREF kItemHoverColor = RGB(0x2A, 0x2A, 0x2A);
constexpr COLORREF kItemActiveColor = RGB(0x30, 0x39, 0x34);
constexpr COLORREF kActiveUnderlineColor = RGB(0x6F, 0xB2, 0x7C);

constexpr UINT kMenuRemove = 8101;
constexpr UINT kMenuRename = 8102;
constexpr UINT kFlyoutMenuFirstCommand = 8200;
constexpr UINT kFlyoutMenuLastCommand = 0xEFFF;

#ifndef MNS_NOTIFYBYPOS
#define MNS_NOTIFYBYPOS 0x08000000
#endif

void LogLastError(const wchar_t* context) {
    const DWORD error_code = GetLastError();
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (GetLastError=%lu).\r\n", context, error_code);
    OutputDebugStringW(buffer);
}

int ScaleForDpi(int value, UINT dpi) {
    return MulDiv(value, static_cast<int>(dpi), 96);
}

bool QueryMessageFont(LOGFONTW* out_font) {
    if (out_font == nullptr) {
        return false;
    }

    NONCLIENTMETRICSW metrics = {};
    metrics.cbSize = sizeof(metrics);
    if (!SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(metrics), &metrics, 0)) {
        return false;
    }

    *out_font = metrics.lfMessageFont;
    return true;
}

void ConfigureMenuForNotifications(HMENU menu) {
    if (menu == nullptr) {
        return;
    }

    MENUINFO menu_info = {};
    menu_info.cbSize = sizeof(menu_info);
    menu_info.fMask = MIM_STYLE;
    if (!GetMenuInfo(menu, &menu_info)) {
        LogLastError(L"GetMenuInfo(Sidebar Flyout)");
        return;
    }

    menu_info.dwStyle |= MNS_NOTIFYBYPOS;
    if (!SetMenuInfo(menu, &menu_info)) {
        LogLastError(L"SetMenuInfo(Sidebar Flyout)");
    }
}

}  // namespace

namespace fileexplorer {

thread_local Sidebar* g_active_flyout_sidebar = nullptr;

void Sidebar::FontDeleter::operator()(HFONT font) const noexcept {
    if (font != nullptr) {
        DeleteObject(font);
    }
}

Sidebar::Sidebar() = default;

Sidebar::~Sidebar() = default;

bool Sidebar::Create(HWND parent, HINSTANCE instance, int control_id) {
    parent_hwnd_ = parent;
    instance_ = instance;
    control_id_ = control_id;

    if (!RegisterWindowClass()) {
        return false;
    }

    hwnd_ = CreateWindowExW(
        0,
        kSidebarClassName,
        L"",
        WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VSCROLL,
        0,
        0,
        0,
        0,
        parent_hwnd_,
        reinterpret_cast<HMENU>(static_cast<INT_PTR>(control_id_)),
        instance_,
        this);
    if (hwnd_ == nullptr) {
        LogLastError(L"CreateWindowExW(FE_Sidebar)");
        return false;
    }

    BuildDefaultItems();
    EnsureFonts();
    UpdateScrollInfo();
    return true;
}

HWND Sidebar::hwnd() const noexcept {
    return hwnd_;
}

void Sidebar::SetDpi(UINT dpi) {
    dpi_ = (dpi == 0U) ? 96U : dpi;
    header_font_.reset();
    item_font_.reset();
    header_font_height_ = 0;
    item_font_height_ = 0;
    EnsureFonts();
    UpdateScrollInfo();
    InvalidateRect(hwnd_, nullptr, TRUE);
}

void Sidebar::SetCurrentPath(const std::wstring& path) {
    current_path_normalized_ = ComparePathNormalized(path);
    InvalidateRect(hwnd_, nullptr, FALSE);
}

LRESULT CALLBACK Sidebar::WindowProc(HWND hwnd, UINT message, WPARAM w_param, LPARAM l_param) {
    Sidebar* self = nullptr;
    if (message == WM_NCCREATE) {
        auto* create_struct = reinterpret_cast<CREATESTRUCTW*>(l_param);
        self = reinterpret_cast<Sidebar*>(create_struct->lpCreateParams);
        if (self == nullptr) {
            return FALSE;
        }

        self->hwnd_ = hwnd;
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    } else {
        self = reinterpret_cast<Sidebar*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    }

    if (self != nullptr) {
        return self->HandleMessage(message, w_param, l_param);
    }

    return DefWindowProcW(hwnd, message, w_param, l_param);
}

LRESULT Sidebar::HandleMessage(UINT message, WPARAM w_param, LPARAM l_param) {
    switch (message) {
    case WM_SIZE:
        UpdateScrollInfo();
        InvalidateRect(hwnd_, nullptr, TRUE);
        return 0;

    case WM_MOUSEWHEEL: {
        const int delta = GET_WHEEL_DELTA_WPARAM(w_param);
        RECT client_rect = {};
        if (!GetClientRect(hwnd_, &client_rect)) {
            LogLastError(L"GetClientRect(Sidebar WM_MOUSEWHEEL)");
            return 0;
        }
        const LayoutInfo layout = BuildLayout(client_rect);
        const int row_step = (std::max)(layout.row_height, 8);
        const int next_offset = scroll_offset_ - ((delta / WHEEL_DELTA) * row_step * 3);
        SetScrollOffset(next_offset);
        return 0;
    }

    case WM_VSCROLL: {
        RECT client_rect = {};
        if (!GetClientRect(hwnd_, &client_rect)) {
            LogLastError(L"GetClientRect(Sidebar WM_VSCROLL)");
            return 0;
        }
        const LayoutInfo layout = BuildLayout(client_rect);

        int next_offset = scroll_offset_;
        switch (LOWORD(w_param)) {
        case SB_LINEUP:
            next_offset -= (std::max)(layout.row_height, 8);
            break;
        case SB_LINEDOWN:
            next_offset += (std::max)(layout.row_height, 8);
            break;
        case SB_PAGEUP:
            next_offset -= (std::max)(1, layout.content_bottom - layout.content_top);
            break;
        case SB_PAGEDOWN:
            next_offset += (std::max)(1, layout.content_bottom - layout.content_top);
            break;
        case SB_THUMBPOSITION:
        case SB_THUMBTRACK: {
            SCROLLINFO scroll_info = {};
            scroll_info.cbSize = sizeof(scroll_info);
            scroll_info.fMask = SIF_TRACKPOS;
            if (GetScrollInfo(hwnd_, SB_VERT, &scroll_info)) {
                next_offset = scroll_info.nTrackPos;
            }
            break;
        }
        case SB_TOP:
            next_offset = 0;
            break;
        case SB_BOTTOM:
            next_offset = MaxScrollOffset(layout);
            break;
        default:
            break;
        }

        SetScrollOffset(next_offset);
        return 0;
    }

    case WM_MOUSEMOVE: {
        if (!tracking_mouse_) {
            TRACKMOUSEEVENT track = {};
            track.cbSize = sizeof(track);
            track.dwFlags = TME_LEAVE;
            track.hwndTrack = hwnd_;
            if (TrackMouseEvent(&track)) {
                tracking_mouse_ = true;
            } else {
                LogLastError(L"TrackMouseEvent(Sidebar)");
            }
        }

        POINT point = {GET_X_LPARAM(l_param), GET_Y_LPARAM(l_param)};
        Section section = Section::None;
        int item_index = -1;
        if (HitTestItem(point, &section, &item_index)) {
            if (section != hovered_section_ || item_index != hovered_index_) {
                hovered_section_ = section;
                hovered_index_ = item_index;
                InvalidateRect(hwnd_, nullptr, FALSE);
            }
        } else if (hovered_section_ != Section::None || hovered_index_ != -1) {
            hovered_section_ = Section::None;
            hovered_index_ = -1;
            InvalidateRect(hwnd_, nullptr, FALSE);
        }
        return 0;
    }

    case WM_MOUSELEAVE:
        tracking_mouse_ = false;
        if (hovered_section_ != Section::None || hovered_index_ != -1) {
            hovered_section_ = Section::None;
            hovered_index_ = -1;
            InvalidateRect(hwnd_, nullptr, FALSE);
        }
        return 0;

    case WM_LBUTTONDOWN: {
        POINT point = {GET_X_LPARAM(l_param), GET_Y_LPARAM(l_param)};
        Section section = Section::None;
        int item_index = -1;
        if (!HitTestItem(point, &section, &item_index) || section == Section::None || item_index < 0) {
            return 0;
        }

        POINT screen_point = point;
        ClientToScreen(hwnd_, &screen_point);
        HandleItemInvoke(section, item_index, screen_point);
        return 0;
    }

    case WM_MENUSELECT: {
        if (!flyout_popup_active_) {
            break;
        }

        const bool left_down = (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0;
        const bool left_down_edge = left_down && !flyout_menu_left_button_was_down_;
        flyout_menu_left_button_was_down_ = left_down;

        if (!left_down_edge) {
            return 0;
        }

        const UINT item_position = LOWORD(w_param);
        const UINT select_flags = HIWORD(w_param);
        if (item_position == 0xFFFFU || (select_flags & MF_POPUP) == 0U) {
            return 0;
        }

        HMENU source_menu = reinterpret_cast<HMENU>(l_param);
        if (source_menu == nullptr) {
            return 0;
        }

        HMENU submenu = GetSubMenu(source_menu, static_cast<int>(item_position));
        if (submenu == nullptr) {
            return 0;
        }

        auto menu_path_it = flyout_menu_paths_.find(submenu);
        if (menu_path_it == flyout_menu_paths_.end()) {
            return 0;
        }

        flyout_pending_selection_ = true;
        flyout_pending_target_ = {menu_path_it->second, true};
        if (!EndMenu()) {
            LogLastError(L"EndMenu(Sidebar WM_MENUSELECT)");
        }
        return 0;
    }

    case WM_MENUCOMMAND: {
        if (!flyout_popup_active_) {
            break;
        }

        HMENU source_menu = reinterpret_cast<HMENU>(l_param);
        if (source_menu == nullptr) {
            break;
        }

        MENUITEMINFOW item_info = {};
        item_info.cbSize = sizeof(item_info);
        item_info.fMask = MIIM_ID;
        if (!GetMenuItemInfoW(source_menu, static_cast<UINT>(w_param), TRUE, &item_info)) {
            LogLastError(L"GetMenuItemInfoW(Sidebar WM_MENUCOMMAND)");
            return 0;
        }

        auto target_it = flyout_command_targets_.find(item_info.wID);
        if (target_it == flyout_command_targets_.end()) {
            const UINT64 position_key = MakeFlyoutMenuPositionKey(source_menu, static_cast<UINT>(w_param));
            auto position_it = flyout_position_targets_.find(position_key);
            if (position_it == flyout_position_targets_.end()) {
                return 0;
            }

            flyout_pending_selection_ = true;
            flyout_pending_target_ = position_it->second;
            if (!EndMenu()) {
                LogLastError(L"EndMenu(Sidebar WM_MENUCOMMAND)");
            }
            return 0;
        }

        flyout_pending_selection_ = true;
        flyout_pending_target_ = target_it->second;
        if (!EndMenu()) {
            LogLastError(L"EndMenu(Sidebar WM_MENUCOMMAND)");
        }
        return 0;
    }

    case WM_COMMAND: {
        if (!flyout_popup_active_) {
            break;
        }

        const UINT command_id = LOWORD(w_param);
        auto target_it = flyout_command_targets_.find(command_id);
        if (target_it == flyout_command_targets_.end()) {
            break;
        }

        flyout_pending_selection_ = true;
        flyout_pending_target_ = target_it->second;
        return 0;
    }

    case WM_INITMENUPOPUP: {
        if (!flyout_popup_active_) {
            break;
        }

        HMENU popup_menu = reinterpret_cast<HMENU>(w_param);
        if (popup_menu == nullptr) {
            break;
        }

        auto path_it = flyout_menu_paths_.find(popup_menu);
        if (path_it == flyout_menu_paths_.end()) {
            break;
        }
        if (flyout_populated_menus_.find(popup_menu) != flyout_populated_menus_.end()) {
            return 0;
        }

        PopulateFlyoutMenu(popup_menu, path_it->second, false, L"");
        return 0;
    }

    case WM_CONTEXTMENU: {
        POINT screen_point = {GET_X_LPARAM(l_param), GET_Y_LPARAM(l_param)};
        POINT client_point = screen_point;

        Section section = Section::None;
        int item_index = -1;
        if (screen_point.x == -1 && screen_point.y == -1) {
            if (hovered_section_ != Section::None && hovered_index_ >= 0) {
                section = hovered_section_;
                item_index = hovered_index_;
            }
        } else {
            if (ScreenToClient(hwnd_, &client_point)) {
                HitTestItem(client_point, &section, &item_index);
            } else {
                LogLastError(L"ScreenToClient(Sidebar WM_CONTEXTMENU)");
            }
        }

        if (section == Section::None || item_index < 0) {
            return 0;
        }

        if (screen_point.x == -1 && screen_point.y == -1) {
            RECT client_rect = {};
            if (GetClientRect(hwnd_, &client_rect)) {
                const LayoutInfo layout = BuildLayout(client_rect);
                const int row_height = layout.row_height;
                const int logical_start =
                    (section == Section::Flyout) ? layout.flyout_items_logical_top : layout.favourites_items_logical_top;
                const int row_top = layout.content_top + (logical_start - scroll_offset_) + (item_index * row_height);
                screen_point.x = client_rect.left + ScaleForDpi(24, dpi_);
                screen_point.y = row_top + (row_height / 2);
                ClientToScreen(hwnd_, &screen_point);
            }
        }

        ShowItemContextMenu(screen_point, section, item_index);
        return 0;
    }

    case WM_PAINT: {
        PAINTSTRUCT ps = {};
        HDC hdc = BeginPaint(hwnd_, &ps);
        if (hdc != nullptr) {
            Paint(hdc);
        }
        EndPaint(hwnd_, &ps);
        return 0;
    }

    case WM_ERASEBKGND:
        return 1;

    case WM_NCDESTROY:
        SetWindowLongPtrW(hwnd_, GWLP_USERDATA, 0);
        hwnd_ = nullptr;
        break;

    default:
        break;
    }

    return DefWindowProcW(hwnd_, message, w_param, l_param);
}

bool Sidebar::RegisterWindowClass() {
    WNDCLASSEXW window_class = {};
    window_class.cbSize = sizeof(window_class);
    window_class.style = CS_HREDRAW | CS_VREDRAW;
    window_class.lpfnWndProc = &Sidebar::WindowProc;
    window_class.hInstance = instance_;
    window_class.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    if (window_class.hCursor == nullptr) {
        LogLastError(L"LoadCursorW(Sidebar)");
    }
    window_class.hbrBackground = nullptr;
    window_class.lpszClassName = kSidebarClassName;

    const ATOM atom = RegisterClassExW(&window_class);
    if (atom == 0 && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
        LogLastError(L"RegisterClassExW(FE_Sidebar)");
        return false;
    }
    return true;
}

void Sidebar::EnsureFonts() {
    LOGFONTW base_font = {};
    const bool have_message_font = QueryMessageFont(&base_font);
    if (!have_message_font) {
        base_font.lfHeight = -MulDiv(9, static_cast<int>(dpi_), 72);
        base_font.lfWeight = FW_NORMAL;
        StringCchCopyW(base_font.lfFaceName, LF_FACESIZE, L"Segoe UI");
    }

    // Keep sidebar text density close to details-list typography.
    base_font.lfHeight = MulDiv(base_font.lfHeight, static_cast<int>(dpi_), 96);
    const int desired_item = base_font.lfHeight;
    const int desired_header = desired_item;

    if (header_font_ == nullptr || header_font_height_ != desired_header) {
        LOGFONTW header_lf = base_font;
        header_lf.lfWeight = FW_BOLD;
        header_font_.reset(CreateFontIndirectW(&header_lf));
        if (header_font_ == nullptr) {
            LogLastError(L"CreateFontIndirectW(SidebarHeader)");
        }
        header_font_height_ = desired_header;
    }

    if (item_font_ == nullptr || item_font_height_ != desired_item) {
        LOGFONTW item_lf = base_font;
        item_lf.lfWeight = FW_NORMAL;
        item_font_.reset(CreateFontIndirectW(&item_lf));
        if (item_font_ == nullptr) {
            LogLastError(L"CreateFontIndirectW(SidebarItem)");
        }
        item_font_height_ = desired_item;
    }
}

void Sidebar::BuildDefaultItems() {
    flyout_items_.clear();
    favourite_items_.clear();

    const std::wstring user_home = UserProfilePath();
    flyout_items_.push_back({L"Home", user_home, true});
    flyout_items_.push_back({L"System Drive", L"C:\\", true});
    flyout_items_.push_back({L"Windows", L"C:\\Windows", true});

    favourite_items_.push_back({L"Documents", BuildUserChildPath(L"Documents"), false});
    favourite_items_.push_back({L"Pictures", BuildUserChildPath(L"Pictures"), false});
    favourite_items_.push_back({L"Music", BuildUserChildPath(L"Music"), false});

    auto by_name = [](const SidebarItem& left, const SidebarItem& right) {
        return CompareTextInsensitive(left.name, right.name) < 0;
    };
    std::sort(flyout_items_.begin(), flyout_items_.end(), by_name);
    std::sort(favourite_items_.begin(), favourite_items_.end(), by_name);
}

void Sidebar::Paint(HDC hdc) {
    if (hdc == nullptr || hwnd_ == nullptr) {
        return;
    }

    EnsureFonts();

    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        LogLastError(L"GetClientRect(Sidebar Paint)");
        return;
    }

    HBRUSH background_brush = CreateSolidBrush(kSidebarBackgroundColor);
    if (background_brush != nullptr) {
        FillRect(hdc, &client_rect, background_brush);
        DeleteObject(background_brush);
    }

    HPEN border_pen = CreatePen(PS_SOLID, 1, kSidebarBorderColor);
    if (border_pen != nullptr) {
        HGDIOBJ old_pen = SelectObject(hdc, border_pen);
        MoveToEx(hdc, client_rect.left, client_rect.top, nullptr);
        LineTo(hdc, client_rect.left, client_rect.bottom);
        SelectObject(hdc, old_pen);
        DeleteObject(border_pen);
    }

    const LayoutInfo layout = BuildLayout(client_rect);
    const int search_left = ScaleForDpi(12, dpi_);
    const int search_right = (std::max)(
        search_left + ScaleForDpi(40, dpi_),
        static_cast<int>(client_rect.right) - ScaleForDpi(12, dpi_));

    HBRUSH search_brush = CreateSolidBrush(kSearchBackgroundColor);
    if (search_brush != nullptr) {
        FillRect(hdc, &layout.search_rect, search_brush);
        DeleteObject(search_brush);
    }

    HPEN search_border = CreatePen(PS_SOLID, 1, kSearchBorderColor);
    if (search_border != nullptr) {
        HGDIOBJ old_pen = SelectObject(hdc, search_border);
        HGDIOBJ old_brush = SelectObject(hdc, GetStockObject(NULL_BRUSH));
        Rectangle(hdc, layout.search_rect.left, layout.search_rect.top, layout.search_rect.right, layout.search_rect.bottom);
        SelectObject(hdc, old_brush);
        SelectObject(hdc, old_pen);
        DeleteObject(search_border);
    }

    if (item_font_ != nullptr) {
        SelectObject(hdc, item_font_.get());
    }
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, kSearchTextColor);
    RECT search_text_rect = layout.search_rect;
    search_text_rect.left += ScaleForDpi(10, dpi_);
    DrawTextW(hdc, L"Search (Phase 7)", -1, &search_text_rect, DT_SINGLELINE | DT_VCENTER | DT_LEFT);

    if (layout.content_bottom <= layout.content_top) {
        return;
    }

    if (header_font_ != nullptr) {
        SelectObject(hdc, header_font_.get());
    }
    SetTextColor(hdc, kHeaderTextColor);

    RECT flyout_header_rect = {};
    flyout_header_rect.left = search_left;
    flyout_header_rect.right = search_right;
    flyout_header_rect.top = layout.content_top + (layout.flyout_header_logical_top - scroll_offset_);
    flyout_header_rect.bottom = flyout_header_rect.top + ScaleForDpi(18, dpi_);
    DrawTextW(hdc, L"Flyout Favourites", -1, &flyout_header_rect, DT_SINGLELINE | DT_VCENTER | DT_LEFT);

    RECT favourites_header_rect = {};
    favourites_header_rect.left = search_left;
    favourites_header_rect.right = search_right;
    favourites_header_rect.top = layout.content_top + (layout.favourites_header_logical_top - scroll_offset_);
    favourites_header_rect.bottom = favourites_header_rect.top + ScaleForDpi(18, dpi_);
    DrawTextW(hdc, L"Favourites", -1, &favourites_header_rect, DT_SINGLELINE | DT_VCENTER | DT_LEFT);

    const int item_left = ScaleForDpi(16, dpi_);
    const int item_text_offset = ScaleForDpi(10, dpi_);
    const int right_padding = ScaleForDpi(12, dpi_);
    const int row_height = layout.row_height;

    if (item_font_ != nullptr) {
        SelectObject(hdc, item_font_.get());
    }
    SetTextColor(hdc, kItemTextColor);

    auto draw_item_row = [&](const SidebarItem& item, Section section, int index, int logical_top) {
        RECT row_rect = {};
        row_rect.left = item_left;
        row_rect.right = search_right;
        row_rect.top = layout.content_top + (logical_top - scroll_offset_);
        row_rect.bottom = row_rect.top + row_height;

        if (row_rect.bottom < layout.content_top || row_rect.top > layout.content_bottom) {
            return;
        }

        const bool hovered = (hovered_section_ == section && hovered_index_ == index);

        if (hovered) {
            HBRUSH row_brush = CreateSolidBrush(kItemHoverColor);
            if (row_brush != nullptr) {
                FillRect(hdc, &row_rect, row_brush);
                DeleteObject(row_brush);
            }
        }

        RECT text_rect = row_rect;
        text_rect.left += item_text_offset;
        text_rect.right -= right_padding;
        DrawTextW(hdc, item.name.c_str(), -1, &text_rect, DT_SINGLELINE | DT_VCENTER | DT_LEFT | DT_END_ELLIPSIS);

        if (item.show_arrow) {
            const int triangle_size = ScaleForDpi(7, dpi_);
            const int center_y = row_rect.top + (row_height / 2);
            const int right_x = row_rect.right - ScaleForDpi(10, dpi_);
            POINT points[3] = {
                {right_x - triangle_size, center_y - (triangle_size / 2)},
                {right_x - triangle_size, center_y + (triangle_size / 2)},
                {right_x, center_y},
            };

            HBRUSH arrow_brush = CreateSolidBrush(kItemTextColor);
            if (arrow_brush != nullptr) {
                HGDIOBJ old_pen = SelectObject(hdc, GetStockObject(NULL_PEN));
                HGDIOBJ old_brush = SelectObject(hdc, arrow_brush);
                Polygon(hdc, points, 3);
                SelectObject(hdc, old_brush);
                SelectObject(hdc, old_pen);
                DeleteObject(arrow_brush);
            }
        }

    };

    for (int i = 0; i < static_cast<int>(flyout_items_.size()); ++i) {
        const int logical_top = layout.flyout_items_logical_top + (i * row_height);
        draw_item_row(flyout_items_[i], Section::Flyout, i, logical_top);
    }

    for (int i = 0; i < static_cast<int>(favourite_items_.size()); ++i) {
        const int logical_top = layout.favourites_items_logical_top + (i * row_height);
        draw_item_row(favourite_items_[i], Section::Favourites, i, logical_top);
    }

    const int divider_y = layout.content_top + (layout.divider_logical_y - scroll_offset_);
    if (divider_y >= layout.content_top && divider_y <= layout.content_bottom) {
        HPEN divider_pen = CreatePen(PS_SOLID, 1, kSidebarBorderColor);
        if (divider_pen != nullptr) {
            HGDIOBJ old_pen = SelectObject(hdc, divider_pen);
            MoveToEx(hdc, search_left, divider_y, nullptr);
            LineTo(hdc, search_right, divider_y);
            SelectObject(hdc, old_pen);
            DeleteObject(divider_pen);
        }
    }
}

Sidebar::LayoutInfo Sidebar::BuildLayout(const RECT& client_rect) const {
    LayoutInfo layout = {};

    const int outer_padding = ScaleForDpi(12, dpi_);
    const int search_height = ScaleForDpi(32, dpi_);
    const int section_spacing = ScaleForDpi(6, dpi_);
    const int header_height = ScaleForDpi(16, dpi_);
    const int row_height = (std::max)(ScaleForDpi(22, dpi_), MeasuredItemTextHeight() + ScaleForDpi(8, dpi_));
    const int header_item_gap = ScaleForDpi(3, dpi_);
    const int divider_margin = ScaleForDpi(5, dpi_);

    layout.search_rect.left = outer_padding;
    layout.search_rect.right = (std::max)(layout.search_rect.left + ScaleForDpi(40, dpi_), client_rect.right - outer_padding);
    layout.search_rect.top = outer_padding;
    layout.search_rect.bottom = layout.search_rect.top + search_height;

    layout.content_top = layout.search_rect.bottom + section_spacing;
    layout.content_bottom = client_rect.bottom - outer_padding;
    layout.row_height = row_height;

    int logical = 0;
    layout.flyout_header_logical_top = logical + section_spacing;
    logical = layout.flyout_header_logical_top + header_height + header_item_gap;

    layout.flyout_items_logical_top = logical;
    logical += static_cast<int>(flyout_items_.size()) * row_height;

    logical += divider_margin;
    layout.divider_logical_y = logical;
    logical += 1 + divider_margin;

    layout.favourites_header_logical_top = logical;
    logical += header_height + header_item_gap;

    layout.favourites_items_logical_top = logical;
    logical += static_cast<int>(favourite_items_.size()) * row_height;

    logical += section_spacing;
    layout.total_logical_height = logical;

    return layout;
}

void Sidebar::UpdateScrollInfo() {
    if (hwnd_ == nullptr) {
        return;
    }

    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        LogLastError(L"GetClientRect(Sidebar Scroll)");
        return;
    }

    const LayoutInfo layout = BuildLayout(client_rect);
    const int max_scroll = MaxScrollOffset(layout);
    scroll_offset_ = (std::max)(0, (std::min)(scroll_offset_, max_scroll));

    SCROLLINFO scroll_info = {};
    scroll_info.cbSize = sizeof(scroll_info);
    scroll_info.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
    scroll_info.nMin = 0;
    scroll_info.nMax = (std::max)(0, layout.total_logical_height - 1);
    scroll_info.nPage = static_cast<UINT>((std::max)(0, layout.content_bottom - layout.content_top));
    scroll_info.nPos = scroll_offset_;
    SetScrollInfo(hwnd_, SB_VERT, &scroll_info, TRUE);
}

int Sidebar::MaxScrollOffset(const LayoutInfo& layout) const {
    const int viewport_height = (std::max)(0, layout.content_bottom - layout.content_top);
    return (std::max)(0, layout.total_logical_height - viewport_height);
}

void Sidebar::SetScrollOffset(int offset) {
    if (hwnd_ == nullptr) {
        return;
    }

    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        LogLastError(L"GetClientRect(Sidebar SetScrollOffset)");
        return;
    }

    const LayoutInfo layout = BuildLayout(client_rect);
    const int clamped = (std::max)(0, (std::min)(offset, MaxScrollOffset(layout)));
    if (clamped == scroll_offset_) {
        return;
    }

    scroll_offset_ = clamped;
    UpdateScrollInfo();
    InvalidateRect(hwnd_, nullptr, TRUE);
}

bool Sidebar::HitTestItem(POINT point, Section* section, int* item_index) const {
    if (section == nullptr || item_index == nullptr || hwnd_ == nullptr) {
        return false;
    }

    *section = Section::None;
    *item_index = -1;

    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        return false;
    }

    const LayoutInfo layout = BuildLayout(client_rect);
    if (point.y < layout.content_top || point.y >= layout.content_bottom) {
        return false;
    }

    const int logical_y = (point.y - layout.content_top) + scroll_offset_;
    const int flyout_start = layout.flyout_items_logical_top;
    const int flyout_end = flyout_start + static_cast<int>(flyout_items_.size()) * layout.row_height;
    if (logical_y >= flyout_start && logical_y < flyout_end) {
        *section = Section::Flyout;
        *item_index = (logical_y - flyout_start) / layout.row_height;
        return true;
    }

    const int fav_start = layout.favourites_items_logical_top;
    const int fav_end = fav_start + static_cast<int>(favourite_items_.size()) * layout.row_height;
    if (logical_y >= fav_start && logical_y < fav_end) {
        *section = Section::Favourites;
        *item_index = (logical_y - fav_start) / layout.row_height;
        return true;
    }

    return false;
}

bool Sidebar::IsItemActive(const SidebarItem& item) const {
    if (current_path_normalized_.empty()) {
        return false;
    }
    return CompareTextInsensitive(current_path_normalized_, ComparePathNormalized(item.path)) == 0;
}

void Sidebar::ShowItemContextMenu(POINT screen_point, Section section, int item_index) {
    const std::vector<SidebarItem>* source = nullptr;
    if (section == Section::Flyout) {
        source = &flyout_items_;
    } else if (section == Section::Favourites) {
        source = &favourite_items_;
    }
    if (source == nullptr || item_index < 0 || item_index >= static_cast<int>(source->size())) {
        return;
    }

    HMENU menu = CreatePopupMenu();
    if (menu == nullptr) {
        LogLastError(L"CreatePopupMenu(Sidebar)");
        return;
    }

    if (!AppendMenuW(menu, MF_STRING, kMenuRemove, L"Remove")) {
        LogLastError(L"AppendMenuW(Sidebar Remove)");
    }
    if (!AppendMenuW(menu, MF_STRING, kMenuRename, L"Rename")) {
        LogLastError(L"AppendMenuW(Sidebar Rename)");
    }

    const UINT command = static_cast<UINT>(TrackPopupMenuEx(
        menu,
        TPM_RIGHTBUTTON | TPM_RETURNCMD,
        screen_point.x,
        screen_point.y,
        hwnd_,
        nullptr));
    (void)command;

    if (!DestroyMenu(menu)) {
        LogLastError(L"DestroyMenu(Sidebar)");
    }
}

void Sidebar::HandleItemInvoke(Section section, int item_index, POINT screen_point) {
    const SidebarItem* item = nullptr;
    if (section == Section::Flyout) {
        if (item_index >= 0 && item_index < static_cast<int>(flyout_items_.size())) {
            item = &flyout_items_[item_index];
        }
    } else if (section == Section::Favourites) {
        if (item_index >= 0 && item_index < static_cast<int>(favourite_items_.size())) {
            item = &favourite_items_[item_index];
        }
    }

    if (item == nullptr) {
        return;
    }

    if (section == Section::Flyout) {
        ShowFlyoutPopup(*item, screen_point);
        return;
    }

    PostNavigatePath(item->path);
}

void Sidebar::PostNavigatePath(const std::wstring& path) const {
    if (parent_hwnd_ == nullptr || path.empty()) {
        return;
    }

    auto payload = std::make_unique<std::wstring>(NormalizePath(path));
    std::wstring* raw_payload = payload.get();

    const DWORD parent_thread_id = GetWindowThreadProcessId(parent_hwnd_, nullptr);
    const DWORD current_thread_id = GetCurrentThreadId();
    if (parent_thread_id != 0 && parent_thread_id == current_thread_id) {
        // We are on the same UI thread as MainWindow; send synchronously to avoid
        // popup-loop/queue timing edge cases dropping apparent navigation.
        payload.release();
        SendMessageW(
            parent_hwnd_,
            WM_FE_SIDEBAR_NAVIGATE,
            0,
            reinterpret_cast<LPARAM>(raw_payload));
        return;
    }

    if (!PostMessageW(
            parent_hwnd_,
            WM_FE_SIDEBAR_NAVIGATE,
            0,
            reinterpret_cast<LPARAM>(raw_payload))) {
        LogLastError(L"PostMessageW(WM_FE_SIDEBAR_NAVIGATE)");
        return;
    }
    payload.release();
}

void Sidebar::OpenPathOrNavigate(const std::wstring& path, bool is_folder) const {
    if (path.empty()) {
        return;
    }

    const std::wstring normalized = NormalizePath(path);
    if (is_folder) {
        PostNavigatePath(normalized);
        return;
    }

    const HINSTANCE open_result =
        ShellExecuteW(hwnd_, L"open", normalized.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    const INT_PTR shell_status = reinterpret_cast<INT_PTR>(open_result);
    if (shell_status <= 32) {
        wchar_t buffer[256] = {};
        swprintf_s(
            buffer,
            L"[FileExplorer] ShellExecuteW(Sidebar OpenPath) failed (status=%lld path=%s).\r\n",
            static_cast<long long>(shell_status),
            normalized.c_str());
        OutputDebugStringW(buffer);
    }
}

void Sidebar::InstallFlyoutMenuHook() {
    RemoveFlyoutMenuHook();
    g_active_flyout_sidebar = this;
    flyout_menu_hook_ = SetWindowsHookExW(
        WH_MSGFILTER,
        &Sidebar::MenuMsgFilterProc,
        nullptr,
        GetCurrentThreadId());
    if (flyout_menu_hook_ == nullptr) {
        LogLastError(L"SetWindowsHookExW(Sidebar Flyout)");
        if (g_active_flyout_sidebar == this) {
            g_active_flyout_sidebar = nullptr;
        }
    }
}

void Sidebar::RemoveFlyoutMenuHook() {
    if (flyout_menu_hook_ != nullptr) {
        if (!UnhookWindowsHookEx(flyout_menu_hook_)) {
            LogLastError(L"UnhookWindowsHookEx(Sidebar Flyout)");
        }
        flyout_menu_hook_ = nullptr;
    }
    if (g_active_flyout_sidebar == this) {
        g_active_flyout_sidebar = nullptr;
    }
}

bool Sidebar::HandleFlyoutMenuHookMessage(const MSG& msg) {
    if (!flyout_popup_active_) {
        return false;
    }

    if (msg.message != WM_LBUTTONUP) {
        return false;
    }

    FlyoutCommandTarget target = {};
    if (!ResolveFlyoutPopupParentFromPoint(msg.pt, &target)) {
        return false;
    }

    flyout_pending_selection_ = true;
    flyout_pending_target_ = target;
    if (!EndMenu()) {
        LogLastError(L"EndMenu(Sidebar Hook)");
    }
    return true;
}

bool Sidebar::ResolveFlyoutPopupParentFromPoint(POINT screen_point, FlyoutCommandTarget* target) const {
    if (target == nullptr) {
        return false;
    }

    for (HMENU menu : flyout_populated_menus_) {
        if (menu == nullptr) {
            continue;
        }

        const int item_position = MenuItemFromPoint(hwnd_, menu, screen_point);
        if (item_position < 0) {
            continue;
        }

        const UINT64 position_key = MakeFlyoutMenuPositionKey(menu, static_cast<UINT>(item_position));
        auto it = flyout_position_targets_.find(position_key);
        if (it == flyout_position_targets_.end()) {
            continue;
        }

        *target = it->second;
        return true;
    }

    return false;
}

LRESULT CALLBACK Sidebar::MenuMsgFilterProc(int code, WPARAM w_param, LPARAM l_param) {
    if (code == MSGF_MENU && g_active_flyout_sidebar != nullptr && l_param != 0) {
        const MSG* msg = reinterpret_cast<const MSG*>(l_param);
        if (msg != nullptr && g_active_flyout_sidebar->HandleFlyoutMenuHookMessage(*msg)) {
            return 1;
        }
    }

    return CallNextHookEx(nullptr, code, w_param, l_param);
}

void Sidebar::ShowFlyoutPopup(const SidebarItem& item, POINT screen_point) {
    ClearFlyoutPopupState();

    HMENU menu = CreatePopupMenu();
    if (menu == nullptr) {
        LogLastError(L"CreatePopupMenu(Sidebar Flyout)");
        return;
    }
    ConfigureMenuForNotifications(menu);

    std::wstring top_open_label = L"Open ";
    top_open_label.append(item.name.empty() ? item.path : item.name);
    top_open_label.append(L" folder");

    if (!PopulateFlyoutMenu(menu, item.path, true, top_open_label)) {
        if (!DestroyMenu(menu)) {
            LogLastError(L"DestroyMenu(Sidebar Flyout Empty)");
        }
        PostNavigatePath(item.path);
        ClearFlyoutPopupState();
        return;
    }

    flyout_popup_active_ = true;
    flyout_menu_left_button_was_down_ = (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0;
    InstallFlyoutMenuHook();

    const UINT selected_command = static_cast<UINT>(TrackPopupMenuEx(
        menu,
        TPM_RIGHTBUTTON | TPM_RETURNCMD,
        screen_point.x,
        screen_point.y,
        hwnd_,
        nullptr));
    RemoveFlyoutMenuHook();

    FlyoutCommandTarget selected_target = {};
    bool have_selected_target = false;
    if (selected_command != 0U) {
        auto selected_it = flyout_command_targets_.find(selected_command);
        if (selected_it != flyout_command_targets_.end()) {
            selected_target = selected_it->second;
            have_selected_target = true;
        }
    }

    if (have_selected_target) {
        wchar_t buffer[512] = {};
        swprintf_s(
            buffer,
            L"[FileExplorer] Flyout selection resolved (cmd=%u folder=%d path=%s).\r\n",
            static_cast<unsigned>(selected_command),
            selected_target.is_folder ? 1 : 0,
            selected_target.path.c_str());
        OutputDebugStringW(buffer);
    } else {
        wchar_t buffer[256] = {};
        swprintf_s(
            buffer,
            L"[FileExplorer] Flyout selection not resolved (cmd=%u).\r\n",
            static_cast<unsigned>(selected_command));
        OutputDebugStringW(buffer);
    }

    flyout_popup_active_ = false;

    if (!DestroyMenu(menu)) {
        LogLastError(L"DestroyMenu(Sidebar Flyout)");
    }

    if (have_selected_target) {
        OpenPathOrNavigate(selected_target.path, selected_target.is_folder);
        ClearFlyoutPopupState();
        return;
    }

    if (flyout_pending_selection_) {
        OpenPathOrNavigate(flyout_pending_target_.path, flyout_pending_target_.is_folder);
    }

    ClearFlyoutPopupState();
}

void Sidebar::ClearFlyoutPopupState() {
    RemoveFlyoutMenuHook();
    flyout_popup_active_ = false;
    next_flyout_command_id_ = kFlyoutMenuFirstCommand;
    flyout_command_targets_.clear();
    flyout_position_targets_.clear();
    flyout_menu_paths_.clear();
    flyout_populated_menus_.clear();
    flyout_pending_selection_ = false;
    flyout_pending_target_ = {};
    flyout_menu_left_button_was_down_ = false;
}

UINT64 Sidebar::MakeFlyoutMenuPositionKey(HMENU menu, UINT item_position) {
    const UINT_PTR menu_bits = reinterpret_cast<UINT_PTR>(menu);
    return (static_cast<UINT64>(menu_bits) << 32) | static_cast<UINT64>(item_position);
}

void Sidebar::RememberFlyoutMenuPositionTarget(
    HMENU menu,
    UINT item_position,
    const FlyoutCommandTarget& target) {
    if (menu == nullptr) {
        return;
    }
    flyout_position_targets_[MakeFlyoutMenuPositionKey(menu, item_position)] = target;
}

bool Sidebar::PopulateFlyoutMenu(
    HMENU menu,
    const std::wstring& base_path,
    bool include_top_open_row,
    const std::wstring& top_open_label) {
    if (menu == nullptr) {
        return false;
    }

    while (GetMenuItemCount(menu) > 0) {
        if (!DeleteMenu(menu, 0, MF_BYPOSITION)) {
            LogLastError(L"DeleteMenu(Sidebar Flyout Clear)");
            break;
        }
    }

    bool appended_any = false;
    if (include_top_open_row) {
        const UINT command_id = NextFlyoutCommandId();
        const FlyoutCommandTarget target = {NormalizePath(base_path), true};
        flyout_command_targets_[command_id] = target;
        const std::wstring label = top_open_label.empty() ? L"Open folder" : top_open_label;
        if (!AppendMenuW(menu, MF_STRING, command_id, label.c_str())) {
            LogLastError(L"AppendMenuW(Sidebar Flyout OpenRoot)");
        } else {
            const int count = GetMenuItemCount(menu);
            if (count > 0) {
                RememberFlyoutMenuPositionTarget(menu, static_cast<UINT>(count - 1), target);
            }
            appended_any = true;
        }
    }

    const std::vector<FlyoutEntry> entries = EnumerateFlyoutEntries(base_path);
    if (entries.empty()) {
        flyout_populated_menus_.insert(menu);
        return appended_any;
    }

    if (include_top_open_row) {
        if (!AppendMenuW(menu, MF_SEPARATOR, 0, nullptr)) {
            LogLastError(L"AppendMenuW(Sidebar Flyout Separator)");
        }
    }

    for (const FlyoutEntry& entry : entries) {
        if (entry.is_folder) {
            const std::vector<FlyoutEntry> child_entries = EnumerateFlyoutEntries(entry.path);
            if (!child_entries.empty()) {
                HMENU submenu = CreatePopupMenu();
                if (submenu != nullptr) {
                    ConfigureMenuForNotifications(submenu);
                    flyout_menu_paths_[submenu] = entry.path;
                    if (!AppendMenuW(submenu, MF_STRING | MF_GRAYED, 0, L"...")) {
                        LogLastError(L"AppendMenuW(Sidebar Flyout Placeholder)");
                    }
                    const UINT command_id = NextFlyoutCommandId();
                    const FlyoutCommandTarget target = {entry.path, true};
                    flyout_command_targets_[command_id] = target;

                    MENUITEMINFOW menu_item = {};
                    menu_item.cbSize = sizeof(menu_item);
                    menu_item.fMask = MIIM_ID | MIIM_SUBMENU | MIIM_STRING;
                    menu_item.wID = command_id;
                    menu_item.hSubMenu = submenu;
                    menu_item.dwTypeData = const_cast<LPWSTR>(entry.name.c_str());

                    if (!InsertMenuItemW(menu, GetMenuItemCount(menu), TRUE, &menu_item)) {
                        LogLastError(L"InsertMenuItemW(Sidebar Flyout Submenu)");
                        flyout_command_targets_.erase(command_id);
                        flyout_menu_paths_.erase(submenu);
                        if (!DestroyMenu(submenu)) {
                            LogLastError(L"DestroyMenu(Sidebar Flyout Submenu)");
                        }
                    } else {
                        const int count = GetMenuItemCount(menu);
                        if (count > 0) {
                            RememberFlyoutMenuPositionTarget(menu, static_cast<UINT>(count - 1), target);
                        }
                        appended_any = true;
                    }
                    continue;
                }
                LogLastError(L"CreatePopupMenu(Sidebar Flyout Submenu)");
            }
        }

        const UINT command_id = NextFlyoutCommandId();
        const FlyoutCommandTarget target = {entry.path, entry.is_folder};
        flyout_command_targets_[command_id] = target;
        if (!AppendMenuW(menu, MF_STRING, command_id, entry.name.c_str())) {
            LogLastError(L"AppendMenuW(Sidebar Flyout Entry)");
            continue;
        }
        const int count = GetMenuItemCount(menu);
        if (count > 0) {
            RememberFlyoutMenuPositionTarget(menu, static_cast<UINT>(count - 1), target);
        }
        appended_any = true;
    }

    flyout_populated_menus_.insert(menu);
    return appended_any;
}

UINT Sidebar::NextFlyoutCommandId() {
    UINT candidate = next_flyout_command_id_;
    if (candidate < kFlyoutMenuFirstCommand || candidate > kFlyoutMenuLastCommand) {
        candidate = kFlyoutMenuFirstCommand;
    }

    while (flyout_command_targets_.contains(candidate)) {
        ++candidate;
        if (candidate > kFlyoutMenuLastCommand) {
            candidate = kFlyoutMenuFirstCommand;
        }
    }

    next_flyout_command_id_ = candidate + 1;
    if (next_flyout_command_id_ > kFlyoutMenuLastCommand) {
        next_flyout_command_id_ = kFlyoutMenuFirstCommand;
    }
    return candidate;
}

std::wstring Sidebar::JoinPath(const std::wstring& base_path, const std::wstring& child_name) const {
    std::wstring combined = NormalizePath(base_path);
    if (!combined.empty() && combined.back() != L'\\') {
        combined.push_back(L'\\');
    }
    combined.append(child_name);
    return NormalizePath(std::move(combined));
}

std::vector<Sidebar::FlyoutEntry> Sidebar::EnumerateFlyoutEntries(const std::wstring& base_path) const {
    std::vector<FlyoutEntry> entries;

    std::wstring wildcard = NormalizePath(base_path);
    if (wildcard.empty()) {
        return entries;
    }
    if (wildcard.back() != L'\\') {
        wildcard.push_back(L'\\');
    }
    wildcard.push_back(L'*');

    WIN32_FIND_DATAW find_data = {};
    HANDLE find_handle = FindFirstFileW(wildcard.c_str(), &find_data);
    if (find_handle == INVALID_HANDLE_VALUE) {
        return entries;
    }

    do {
        const wchar_t* name = find_data.cFileName;
        if (name[0] == L'.' && (name[1] == L'\0' || (name[1] == L'.' && name[2] == L'\0'))) {
            continue;
        }
        if ((find_data.dwFileAttributes & (FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM)) != 0) {
            continue;
        }

        const bool is_folder = (find_data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        entries.push_back({std::wstring(name), JoinPath(base_path, name), is_folder});
    } while (FindNextFileW(find_handle, &find_data));

    if (!FindClose(find_handle)) {
        LogLastError(L"FindClose(Sidebar Flyout)");
    }

    std::sort(
        entries.begin(),
        entries.end(),
        [](const FlyoutEntry& left, const FlyoutEntry& right) {
            if (left.is_folder != right.is_folder) {
                return left.is_folder && !right.is_folder;
            }
            const int by_name = Sidebar::CompareTextInsensitive(left.name, right.name);
            if (by_name != 0) {
                return by_name < 0;
            }
            return Sidebar::CompareTextInsensitive(left.path, right.path) < 0;
        });
    return entries;
}

int Sidebar::MeasuredItemTextHeight() const {
    if (item_font_ == nullptr) {
        return ScaleForDpi(14, dpi_);
    }

    HDC hdc = GetDC(hwnd_);
    if (hdc == nullptr) {
        return ScaleForDpi(14, dpi_);
    }

    HGDIOBJ old_font = SelectObject(hdc, item_font_.get());
    TEXTMETRICW tm = {};
    const BOOL ok = GetTextMetricsW(hdc, &tm);
    SelectObject(hdc, old_font);
    ReleaseDC(hwnd_, hdc);

    if (!ok) {
        return ScaleForDpi(14, dpi_);
    }

    return static_cast<int>(tm.tmHeight);
}

std::wstring Sidebar::NormalizePath(std::wstring path) {
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

std::wstring Sidebar::ComparePathNormalized(std::wstring path) {
    path = NormalizePath(std::move(path));
    std::transform(
        path.begin(),
        path.end(),
        path.begin(),
        [](wchar_t ch) { return static_cast<wchar_t>(towlower(ch)); });
    return path;
}

std::wstring Sidebar::UserProfilePath() {
    wchar_t buffer[MAX_PATH] = {};
    const DWORD length = GetEnvironmentVariableW(L"USERPROFILE", buffer, static_cast<DWORD>(std::size(buffer)));
    if (length > 0 && length < std::size(buffer)) {
        return NormalizePath(std::wstring(buffer));
    }
    return L"C:\\";
}

std::wstring Sidebar::BuildUserChildPath(const wchar_t* child_name) {
    std::wstring result = UserProfilePath();
    if (!result.empty() && result.back() != L'\\') {
        result.push_back(L'\\');
    }
    result.append(child_name);
    return NormalizePath(std::move(result));
}

int Sidebar::CompareTextInsensitive(const std::wstring& left, const std::wstring& right) {
    const int result = CompareStringOrdinal(
        left.c_str(),
        static_cast<int>(left.size()),
        right.c_str(),
        static_cast<int>(right.size()),
        TRUE);
    if (result == CSTR_LESS_THAN) {
        return -1;
    }
    if (result == CSTR_GREATER_THAN) {
        return 1;
    }
    return 0;
}

}  // namespace fileexplorer
