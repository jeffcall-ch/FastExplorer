#include "Sidebar.h"

#include <CommCtrl.h>
#include <Shellapi.h>
#include <Shobjidl.h>
#include <Windowsx.h>
#include <strsafe.h>
#include <uxtheme.h>

#include <algorithm>
#include <cwchar>
#include <cwctype>
#include <thread>
#include <utility>
#include <vector>

namespace {

constexpr wchar_t kSidebarClassName[] = L"FE_Sidebar";

constexpr COLORREF kSidebarBackgroundColor = RGB(0x1C, 0x22, 0x2A);
constexpr COLORREF kSidebarBorderColor = RGB(0x2E, 0x37, 0x42);
constexpr COLORREF kSidebarSectionSeparatorColor = RGB(0x32, 0x3A, 0x46);
constexpr COLORREF kSearchBackgroundColor = RGB(0x21, 0x28, 0x31);
constexpr COLORREF kSearchBorderColor = RGB(0x39, 0x44, 0x50);
constexpr COLORREF kSearchBorderFocusedColor = RGB(0x52, 0x60, 0x71);
constexpr COLORREF kSearchTextColor = RGB(0xA9, 0xB4, 0xC1);
constexpr COLORREF kHeaderTextColor = RGB(0xD5, 0xDE, 0xE8);
constexpr COLORREF kItemTextColor = RGB(0xE2, 0xE8, 0xF0);
constexpr COLORREF kItemHoverColor = RGB(0x2A, 0x33, 0x3E);

constexpr UINT kMenuRemove = 8101;
constexpr UINT kMenuRename = 8102;
constexpr UINT kFlyoutMenuFirstCommand = 8200;
constexpr UINT kFlyoutMenuLastCommand = 0xEFFF;
constexpr UINT_PTR kSearchEditSubclassId = 2;
constexpr UINT_PTR kRenameEditSubclassId = 3;
constexpr UINT kSidebarMessageFlyoutLoaded = WM_APP + 201;

#ifndef MNS_NOTIFYBYPOS
#define MNS_NOTIFYBYPOS 0x08000000
#endif

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

int ScaleForDpi(int value, UINT dpi) {
    return MulDiv(value, static_cast<int>(dpi), 96);
}

bool EndsWithInsensitive(const std::wstring& text, const wchar_t* suffix) {
    if (suffix == nullptr) {
        return false;
    }

    const size_t suffix_length = wcslen(suffix);
    if (suffix_length == 0 || text.size() < suffix_length) {
        return false;
    }

    const wchar_t* tail = text.c_str() + (text.size() - suffix_length);
    const int compare = CompareStringOrdinal(
        tail,
        static_cast<int>(suffix_length),
        suffix,
        static_cast<int>(suffix_length),
        TRUE);
    return compare == CSTR_EQUAL;
}

bool TryResolveShortcutTarget(
    const std::wstring& shortcut_path,
    std::wstring* target_path,
    bool* target_is_folder) {
    if (target_path == nullptr || target_is_folder == nullptr) {
        return false;
    }
    if (!EndsWithInsensitive(shortcut_path, L".lnk")) {
        return false;
    }

    IShellLinkW* shell_link = nullptr;
    HRESULT hr = CoCreateInstance(
        CLSID_ShellLink,
        nullptr,
        CLSCTX_INPROC_SERVER,
        IID_PPV_ARGS(&shell_link));
    if (FAILED(hr) || shell_link == nullptr) {
        return false;
    }

    IPersistFile* persist_file = nullptr;
    hr = shell_link->QueryInterface(IID_PPV_ARGS(&persist_file));
    if (FAILED(hr) || persist_file == nullptr) {
        shell_link->Release();
        return false;
    }

    hr = persist_file->Load(shortcut_path.c_str(), STGM_READ);
    if (SUCCEEDED(hr)) {
        (void)shell_link->Resolve(nullptr, SLR_NO_UI | SLR_NOUPDATE | SLR_NOTRACK);

        WIN32_FIND_DATAW find_data = {};
        wchar_t resolved_path[MAX_PATH] = {};
        hr = shell_link->GetPath(resolved_path, static_cast<int>(std::size(resolved_path)), &find_data, SLGP_RAWPATH);
        if (SUCCEEDED(hr) && resolved_path[0] != L'\0') {
            *target_path = resolved_path;
            DWORD attributes = find_data.dwFileAttributes;
            if (attributes == 0 || attributes == INVALID_FILE_ATTRIBUTES) {
                attributes = GetFileAttributesW(resolved_path);
            }
            *target_is_folder =
                attributes != INVALID_FILE_ATTRIBUTES &&
                (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
            persist_file->Release();
            shell_link->Release();
            return true;
        }
    }

    persist_file->Release();
    shell_link->Release();
    return false;
}

std::wstring TrimWhitespace(std::wstring text) {
    const auto not_space = [](wchar_t ch) { return !iswspace(ch); };
    const auto begin = std::find_if(text.begin(), text.end(), not_space);
    if (begin == text.end()) {
        return L"";
    }
    const auto reverse_begin = std::find_if(text.rbegin(), text.rend(), not_space).base();
    return std::wstring(begin, reverse_begin);
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

}  // namespace

namespace fileexplorer {

thread_local Sidebar* g_active_flyout_sidebar = nullptr;

void Sidebar::FontDeleter::operator()(HFONT font) const noexcept {
    if (font != nullptr) {
        DeleteObject(font);
    }
}

void Sidebar::BrushDeleter::operator()(HBRUSH brush) const noexcept {
    if (brush != nullptr) {
        DeleteObject(brush);
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

    const HRESULT theme_hr = SetWindowTheme(hwnd_, L"DarkMode_Explorer", nullptr);
    if (FAILED(theme_hr)) {
        LogHResult(L"SetWindowTheme(FE_Sidebar)", theme_hr);
    }

    BuildDefaultItems();
    EnsureFonts();
    if (!CreateSearchEditControl()) {
        return false;
    }
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

void Sidebar::FocusSearch() {
    if (search_edit_hwnd_ == nullptr) {
        return;
    }

    SetFocus(search_edit_hwnd_);
    SendMessageW(search_edit_hwnd_, EM_SETSEL, 0, -1);
}

void Sidebar::SetSearchText(const std::wstring& text) {
    if (search_edit_hwnd_ == nullptr) {
        return;
    }

    SetWindowTextW(search_edit_hwnd_, text.c_str());
    UpdateScrollInfo();
    InvalidateRect(hwnd_, nullptr, FALSE);
}

void Sidebar::ClearSearchText(bool notify_parent) {
    if (search_edit_hwnd_ != nullptr) {
        SetWindowTextW(search_edit_hwnd_, L"");
    }

    if (notify_parent) {
        DispatchSearchClear();
    }

    UpdateScrollInfo();
    InvalidateRect(hwnd_, nullptr, FALSE);
}

bool Sidebar::AddRegularFavourite(const std::wstring& path) {
    if (!favourites_store_.AddRegular(path)) {
        return false;
    }
    RefreshItemsFromStore();
    UpdateScrollInfo();
    InvalidateRect(hwnd_, nullptr, FALSE);
    return true;
}

bool Sidebar::AddFlyoutFavourite(const std::wstring& path) {
    if (!favourites_store_.AddFlyout(path)) {
        return false;
    }
    RefreshItemsFromStore();
    UpdateScrollInfo();
    InvalidateRect(hwnd_, nullptr, FALSE);
    return true;
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

LRESULT CALLBACK Sidebar::SearchEditSubclassProc(
    HWND hwnd,
    UINT message,
    WPARAM w_param,
    LPARAM l_param,
    UINT_PTR subclass_id,
    DWORD_PTR ref_data) {
    (void)subclass_id;
    auto* self = reinterpret_cast<Sidebar*>(ref_data);
    if (self == nullptr) {
        return DefSubclassProc(hwnd, message, w_param, l_param);
    }

    switch (message) {
    case WM_KEYDOWN:
        if (w_param == VK_RETURN) {
            self->DispatchSearchRequest();
            return 0;
        }
        if (w_param == VK_ESCAPE) {
            self->ClearSearchText(true);
            return 0;
        }
        break;

    case WM_SETFOCUS:
    case WM_KILLFOCUS:
        if (self->hwnd_ != nullptr) {
            InvalidateRect(self->hwnd_, nullptr, FALSE);
        }
        break;

    default:
        break;
    }

    return DefSubclassProc(hwnd, message, w_param, l_param);
}

LRESULT CALLBACK Sidebar::RenameEditSubclassProc(
    HWND hwnd,
    UINT message,
    WPARAM w_param,
    LPARAM l_param,
    UINT_PTR subclass_id,
    DWORD_PTR ref_data) {
    (void)subclass_id;
    auto* self = reinterpret_cast<Sidebar*>(ref_data);
    if (self == nullptr) {
        return DefSubclassProc(hwnd, message, w_param, l_param);
    }

    switch (message) {
    case WM_KEYDOWN:
        if (w_param == VK_RETURN) {
            self->EndRenameItemEdit(true);
            return 0;
        }
        if (w_param == VK_ESCAPE) {
            self->EndRenameItemEdit(false);
            return 0;
        }
        break;

    case WM_KILLFOCUS:
        self->EndRenameItemEdit(true);
        return 0;

    default:
        break;
    }

    return DefSubclassProc(hwnd, message, w_param, l_param);
}

LRESULT Sidebar::HandleMessage(UINT message, WPARAM w_param, LPARAM l_param) {
    switch (message) {
    case WM_SIZE:
        EndRenameItemEdit(true);
        UpdateScrollInfo();
        InvalidateRect(hwnd_, nullptr, TRUE);
        return 0;

    case WM_MOUSEWHEEL: {
        EndRenameItemEdit(true);
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
        EndRenameItemEdit(true);
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
        RECT client_rect = {};
        if (GetClientRect(hwnd_, &client_rect)) {
            const LayoutInfo layout = BuildLayout(client_rect);
            if (PtInRect(&layout.search_rect, point)) {
                if (clear_button_visible_ && PtInRect(&clear_button_rect_, point)) {
                    ClearSearchText(true);
                } else {
                    FocusSearch();
                }
                return 0;
            }
        }

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
        if (flyout_ignore_initial_mouse_release_) {
            flyout_menu_left_button_was_down_ = left_down;
            if (!left_down) {
                flyout_ignore_initial_mouse_release_ = false;
            }
            return 0;
        }

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

        auto metadata_it = flyout_menu_metadata_.find(submenu);
        if (metadata_it == flyout_menu_metadata_.end()) {
            return 0;
        }

        flyout_pending_selection_ = true;
        flyout_pending_target_ = {metadata_it->second.path, true};
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
        const HWND command_hwnd = reinterpret_cast<HWND>(l_param);
        if (command_hwnd == search_edit_hwnd_) {
            if (HIWORD(w_param) == EN_CHANGE) {
                UpdateScrollInfo();
                InvalidateRect(hwnd_, nullptr, FALSE);
            }
            return 0;
        }

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

        auto metadata_it = flyout_menu_metadata_.find(popup_menu);
        if (metadata_it == flyout_menu_metadata_.end()) {
            break;
        }
        if (metadata_it->second.loaded || metadata_it->second.loading) {
            return 0;
        }

        BeginFlyoutMenuLoad(popup_menu);
        return 0;
    }

    case kSidebarMessageFlyoutLoaded: {
        auto result = std::unique_ptr<FlyoutAsyncResult>(reinterpret_cast<FlyoutAsyncResult*>(l_param));
        if (!result) {
            return 0;
        }

        auto request_it = flyout_request_menus_.find(result->request_id);
        if (request_it == flyout_request_menus_.end()) {
            return 0;
        }

        const HMENU menu = request_it->second;
        flyout_request_menus_.erase(request_it);

        auto metadata_it = flyout_menu_metadata_.find(menu);
        if (metadata_it == flyout_menu_metadata_.end() || metadata_it->second.request_id != result->request_id) {
            return 0;
        }

        FlyoutMenuMetadata metadata = metadata_it->second;
        metadata_it->second.loading = false;
        metadata_it->second.loaded = true;
        metadata_it->second.request_id = 0;
        ApplyFlyoutMenuEntries(menu, metadata, result->entries);
        return 0;
    }

    case WM_CONTEXTMENU: {
        EndRenameItemEdit(true);
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

    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORSTATIC:
        if (reinterpret_cast<HWND>(l_param) == search_edit_hwnd_ ||
            reinterpret_cast<HWND>(l_param) == rename_edit_hwnd_) {
            HDC hdc = reinterpret_cast<HDC>(w_param);
            if (hdc != nullptr) {
                SetTextColor(hdc, kItemTextColor);
                SetBkColor(hdc, kSearchBackgroundColor);
                SetBkMode(hdc, OPAQUE);
            }
            if (search_edit_brush_ != nullptr) {
                return reinterpret_cast<LRESULT>(search_edit_brush_.get());
            }
        }
        break;

    case WM_NCDESTROY:
        EndRenameItemEdit(false);
        if (search_edit_hwnd_ != nullptr) {
            RemoveWindowSubclass(search_edit_hwnd_, &Sidebar::SearchEditSubclassProc, kSearchEditSubclassId);
            search_edit_hwnd_ = nullptr;
        }
        if (rename_edit_hwnd_ != nullptr) {
            RemoveWindowSubclass(rename_edit_hwnd_, &Sidebar::RenameEditSubclassProc, kRenameEditSubclassId);
            rename_edit_hwnd_ = nullptr;
        }
        ClearFlyoutPopupState();
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
    const int desired_item = -MulDiv(9, static_cast<int>(dpi_), 72);
    const int desired_header = desired_item;

    auto create_ui_font = [](int pixel_height, LONG weight) -> HFONT {
        LOGFONTW lf = {};
        lf.lfHeight = pixel_height;
        lf.lfWeight = weight;
        lf.lfCharSet = DEFAULT_CHARSET;
        lf.lfOutPrecision = OUT_DEFAULT_PRECIS;
        lf.lfClipPrecision = CLIP_DEFAULT_PRECIS;
        lf.lfQuality = CLEARTYPE_QUALITY;
        lf.lfPitchAndFamily = DEFAULT_PITCH | FF_DONTCARE;

        StringCchCopyW(lf.lfFaceName, LF_FACESIZE, L"Segoe UI Variable Text");
        HFONT font = CreateFontIndirectW(&lf);
        if (font != nullptr) {
            return font;
        }

        StringCchCopyW(lf.lfFaceName, LF_FACESIZE, L"Segoe UI");
        return CreateFontIndirectW(&lf);
    };

    if (header_font_ == nullptr || header_font_height_ != desired_header) {
        header_font_.reset(create_ui_font(desired_header, FW_SEMIBOLD));
        if (header_font_ == nullptr) {
            LogLastError(L"CreateFontIndirectW(SidebarHeader)");
        }
        header_font_height_ = desired_header;
    }

    if (item_font_ == nullptr || item_font_height_ != desired_item) {
        item_font_.reset(create_ui_font(desired_item, FW_NORMAL));
        if (item_font_ == nullptr) {
            LogLastError(L"CreateFontIndirectW(SidebarItem)");
        }
        item_font_height_ = desired_item;
    }

    if (search_edit_hwnd_ != nullptr && item_font_ != nullptr) {
        SendMessageW(search_edit_hwnd_, WM_SETFONT, reinterpret_cast<WPARAM>(item_font_.get()), FALSE);
    }
}

bool Sidebar::CreateSearchEditControl() {
    if (hwnd_ == nullptr) {
        return false;
    }

    search_edit_brush_.reset(CreateSolidBrush(kSearchBackgroundColor));
    if (search_edit_brush_ == nullptr) {
        LogLastError(L"CreateSolidBrush(Sidebar SearchEdit)");
    }

    search_edit_hwnd_ = CreateWindowExW(
        0,
        WC_EDITW,
        L"",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL,
        0,
        0,
        0,
        0,
        hwnd_,
        nullptr,
        instance_,
        nullptr);
    if (search_edit_hwnd_ == nullptr) {
        LogLastError(L"CreateWindowExW(Sidebar SearchEdit)");
        return false;
    }

    const HRESULT theme_hr = SetWindowTheme(search_edit_hwnd_, L"", L"");
    (void)theme_hr;

    if (item_font_ != nullptr) {
        SendMessageW(search_edit_hwnd_, WM_SETFONT, reinterpret_cast<WPARAM>(item_font_.get()), FALSE);
    }
    if (!SetWindowSubclass(
            search_edit_hwnd_,
            &Sidebar::SearchEditSubclassProc,
            kSearchEditSubclassId,
            reinterpret_cast<DWORD_PTR>(this))) {
        LogLastError(L"SetWindowSubclass(Sidebar SearchEdit)");
        return false;
    }

    SendMessageW(search_edit_hwnd_, EM_SETCUEBANNER, FALSE, reinterpret_cast<LPARAM>(L"Search"));
    return true;
}

void Sidebar::UpdateSearchEditLayout(const LayoutInfo& layout) {
    if (search_edit_hwnd_ == nullptr) {
        return;
    }

    RECT edit_rect = layout.search_rect;
    edit_rect.left += ScaleForDpi(10, dpi_);
    edit_rect.top += ScaleForDpi(6, dpi_);
    edit_rect.bottom -= ScaleForDpi(6, dpi_);
    edit_rect.right -= ScaleForDpi(SearchHasText() ? 26 : 10, dpi_);
    if (edit_rect.right <= edit_rect.left) {
        edit_rect.right = edit_rect.left + ScaleForDpi(24, dpi_);
    }

    if (!MoveWindow(
            search_edit_hwnd_,
            edit_rect.left,
            edit_rect.top,
            edit_rect.right - edit_rect.left,
            edit_rect.bottom - edit_rect.top,
            TRUE)) {
        LogLastError(L"MoveWindow(Sidebar SearchEdit)");
    }
}

void Sidebar::BuildDefaultItems() {
    favourites_store_.UseDefaultStoragePaths();
    if (!favourites_store_.Load()) {
        return;
    }

    // Keep the previous Phase 6 default quick-start favourites only on first run.
    if (!favourites_store_.had_existing_files_on_last_load() &&
        favourites_store_.GetFlyouts().empty() &&
        favourites_store_.GetRegulars().empty()) {
        favourites_store_.AddFlyout(UserProfilePath(), L"Home");
        favourites_store_.AddFlyout(L"C:\\", L"System Drive");
        favourites_store_.AddFlyout(L"C:\\Windows", L"Windows");

        favourites_store_.AddRegular(BuildUserChildPath(L"Documents"), L"Documents");
        favourites_store_.AddRegular(BuildUserChildPath(L"Pictures"), L"Pictures");
        favourites_store_.AddRegular(BuildUserChildPath(L"Music"), L"Music");
    }

    RefreshItemsFromStore();
}

void Sidebar::RefreshItemsFromStore() {
    flyout_items_.clear();
    favourite_items_.clear();

    flyout_items_.reserve(favourites_store_.GetFlyouts().size());
    for (const FavouriteEntry& entry : favourites_store_.GetFlyouts()) {
        flyout_items_.push_back({entry.friendly_name, entry.path, true});
    }

    favourite_items_.reserve(favourites_store_.GetRegulars().size());
    for (const FavouriteEntry& entry : favourites_store_.GetRegulars()) {
        favourite_items_.push_back({entry.friendly_name, entry.path, false});
    }
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

    const bool search_focused = (search_edit_hwnd_ != nullptr && GetFocus() == search_edit_hwnd_);
    const COLORREF search_border = search_focused ? kSearchBorderFocusedColor : kSearchBorderColor;
    FillRoundedRect(hdc, layout.search_rect, kSearchBackgroundColor, search_border, ScaleForDpi(6, dpi_));
    UpdateSearchEditLayout(layout);

    clear_button_visible_ = false;
    SetRectEmpty(&clear_button_rect_);
    if (SearchHasText()) {
        clear_button_visible_ = true;
        const int button_size = ScaleForDpi(12, dpi_);
        clear_button_rect_.right = layout.search_rect.right - ScaleForDpi(9, dpi_);
        clear_button_rect_.left = clear_button_rect_.right - button_size;
        clear_button_rect_.top =
            layout.search_rect.top + ((layout.search_rect.bottom - layout.search_rect.top) - button_size) / 2;
        clear_button_rect_.bottom = clear_button_rect_.top + button_size;

        HPEN clear_pen = CreatePen(PS_SOLID, 1, kSearchTextColor);
        if (clear_pen != nullptr) {
            HGDIOBJ old_pen = SelectObject(hdc, clear_pen);
            MoveToEx(hdc, clear_button_rect_.left, clear_button_rect_.top, nullptr);
            LineTo(hdc, clear_button_rect_.right, clear_button_rect_.bottom);
            MoveToEx(hdc, clear_button_rect_.left, clear_button_rect_.bottom, nullptr);
            LineTo(hdc, clear_button_rect_.right, clear_button_rect_.top);
            SelectObject(hdc, old_pen);
            DeleteObject(clear_pen);
        }
    }

    if (layout.content_bottom <= layout.content_top) {
        return;
    }

    if (header_font_ != nullptr) {
        SelectObject(hdc, header_font_.get());
    }
    SetBkMode(hdc, TRANSPARENT);
    SetBkColor(hdc, kSidebarBackgroundColor);
    SetTextColor(hdc, kHeaderTextColor);

    RECT flyout_header_rect = {};
    flyout_header_rect.left = search_left;
    flyout_header_rect.right = search_right;
    flyout_header_rect.top = layout.content_top + (layout.flyout_header_logical_top - scroll_offset_);
    flyout_header_rect.bottom = flyout_header_rect.top + ScaleForDpi(18, dpi_);
    flyout_header_rect.left += ScaleForDpi(2, dpi_);
    DrawTextW(hdc, L"Flyout Favourites", -1, &flyout_header_rect, DT_SINGLELINE | DT_VCENTER | DT_LEFT);

    RECT favourites_header_rect = {};
    favourites_header_rect.left = search_left;
    favourites_header_rect.right = search_right;
    favourites_header_rect.top = layout.content_top + (layout.favourites_header_logical_top - scroll_offset_);
    favourites_header_rect.bottom = favourites_header_rect.top + ScaleForDpi(18, dpi_);
    favourites_header_rect.left += ScaleForDpi(2, dpi_);
    DrawTextW(hdc, L"Favourites", -1, &favourites_header_rect, DT_SINGLELINE | DT_VCENTER | DT_LEFT);

    const int item_left = ScaleForDpi(16, dpi_);
    const int item_text_offset = ScaleForDpi(10, dpi_);
    const int right_padding = ScaleForDpi(12, dpi_);
    const int row_height = layout.row_height;

    if (item_font_ != nullptr) {
        SelectObject(hdc, item_font_.get());
    }
    SetBkMode(hdc, TRANSPARENT);
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
            FillRoundedRect(hdc, row_rect, kItemHoverColor, kItemHoverColor, ScaleForDpi(6, dpi_));
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
        HPEN divider_pen = CreatePen(PS_SOLID, 1, kSidebarSectionSeparatorColor);
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
    const int section_spacing = ScaleForDpi(8, dpi_);
    const int header_height = ScaleForDpi(18, dpi_);
    const int row_height = (std::max)(ScaleForDpi(22, dpi_), MeasuredItemTextHeight() + ScaleForDpi(8, dpi_));
    const int header_item_gap = ScaleForDpi(4, dpi_);
    const int divider_margin = ScaleForDpi(8, dpi_);

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
    UpdateSearchEditLayout(layout);
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

    if (!DestroyMenu(menu)) {
        LogLastError(L"DestroyMenu(Sidebar)");
    }

    switch (command) {
    case kMenuRemove:
        RemoveItem(section, item_index);
        break;
    case kMenuRename:
        BeginRenameItem(section, item_index);
        break;
    default:
        break;
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

bool Sidebar::RemoveItem(Section section, int item_index) {
    const FavouriteType type =
        (section == Section::Flyout) ? FavouriteType::Flyout : FavouriteType::Regular;
    if (section != Section::Flyout && section != Section::Favourites) {
        return false;
    }
    if (item_index < 0) {
        return false;
    }

    if (!favourites_store_.Remove(type, static_cast<size_t>(item_index))) {
        return false;
    }

    RefreshItemsFromStore();
    EndRenameItemEdit(false);
    UpdateScrollInfo();
    InvalidateRect(hwnd_, nullptr, FALSE);
    return true;
}

bool Sidebar::BeginRenameItem(Section section, int item_index) {
    if (section != Section::Flyout && section != Section::Favourites) {
        return false;
    }
    if (item_index < 0) {
        return false;
    }

    const std::vector<SidebarItem>* items = (section == Section::Flyout) ? &flyout_items_ : &favourite_items_;
    if (item_index >= static_cast<int>(items->size())) {
        return false;
    }

    EndRenameItemEdit(false);

    RECT edit_rect = {};
    if (!ResolveItemTextRect(section, item_index, &edit_rect)) {
        return false;
    }

    rename_edit_hwnd_ = CreateWindowExW(
        0,
        WC_EDITW,
        (*items)[item_index].name.c_str(),
        WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
        edit_rect.left,
        edit_rect.top,
        edit_rect.right - edit_rect.left,
        edit_rect.bottom - edit_rect.top,
        hwnd_,
        nullptr,
        instance_,
        nullptr);
    if (rename_edit_hwnd_ == nullptr) {
        LogLastError(L"CreateWindowExW(Sidebar RenameEdit)");
        return false;
    }

    const HRESULT theme_hr = SetWindowTheme(rename_edit_hwnd_, L"", L"");
    (void)theme_hr;
    if (item_font_ != nullptr) {
        SendMessageW(rename_edit_hwnd_, WM_SETFONT, reinterpret_cast<WPARAM>(item_font_.get()), FALSE);
    }
    if (!SetWindowSubclass(
            rename_edit_hwnd_,
            &Sidebar::RenameEditSubclassProc,
            kRenameEditSubclassId,
            reinterpret_cast<DWORD_PTR>(this))) {
        LogLastError(L"SetWindowSubclass(Sidebar RenameEdit)");
        DestroyWindow(rename_edit_hwnd_);
        rename_edit_hwnd_ = nullptr;
        return false;
    }

    rename_section_ = section;
    rename_item_index_ = item_index;
    SetFocus(rename_edit_hwnd_);
    SendMessageW(rename_edit_hwnd_, EM_SETSEL, 0, -1);
    InvalidateRect(hwnd_, nullptr, FALSE);
    return true;
}

void Sidebar::EndRenameItemEdit(bool commit) {
    if (rename_edit_closing_ || rename_edit_hwnd_ == nullptr) {
        return;
    }

    rename_edit_closing_ = true;

    if (commit) {
        const int text_length = GetWindowTextLengthW(rename_edit_hwnd_);
        std::vector<wchar_t> buffer(static_cast<size_t>(text_length) + 1, L'\0');
        if (GetWindowTextW(rename_edit_hwnd_, buffer.data(), static_cast<int>(buffer.size())) > 0) {
            const FavouriteType type =
                (rename_section_ == Section::Flyout) ? FavouriteType::Flyout : FavouriteType::Regular;
            const std::wstring new_name = TrimWhitespace(buffer.data());
            if (!new_name.empty() && rename_item_index_ >= 0) {
                favourites_store_.Rename(type, static_cast<size_t>(rename_item_index_), new_name);
                RefreshItemsFromStore();
            }
        }
    }

    RemoveWindowSubclass(rename_edit_hwnd_, &Sidebar::RenameEditSubclassProc, kRenameEditSubclassId);
    DestroyWindow(rename_edit_hwnd_);
    rename_edit_hwnd_ = nullptr;
    rename_section_ = Section::None;
    rename_item_index_ = -1;
    rename_edit_closing_ = false;
    UpdateScrollInfo();
    InvalidateRect(hwnd_, nullptr, FALSE);
}

bool Sidebar::ResolveItemTextRect(Section section, int item_index, RECT* rect) const {
    if (rect == nullptr || hwnd_ == nullptr) {
        return false;
    }

    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        return false;
    }
    const LayoutInfo layout = BuildLayout(client_rect);

    const int row_height = layout.row_height;
    const int logical_start =
        (section == Section::Flyout) ? layout.flyout_items_logical_top : layout.favourites_items_logical_top;
    const int row_top = layout.content_top + (logical_start - scroll_offset_) + (item_index * row_height);
    const int row_bottom = row_top + row_height;
    if (row_bottom <= layout.content_top || row_top >= layout.content_bottom) {
        return false;
    }

    const int item_left = ScaleForDpi(16, dpi_);
    const int item_text_offset = ScaleForDpi(10, dpi_);
    const int right_padding = ScaleForDpi(12, dpi_);
    const int arrow_reserve = (section == Section::Flyout) ? ScaleForDpi(18, dpi_) : 0;

    rect->left = item_left + item_text_offset - ScaleForDpi(2, dpi_);
    rect->right = layout.search_rect.right - right_padding - arrow_reserve;
    rect->top = row_top + ScaleForDpi(2, dpi_);
    rect->bottom = row_bottom - ScaleForDpi(2, dpi_);
    if (rect->right <= rect->left) {
        rect->right = rect->left + ScaleForDpi(32, dpi_);
    }
    return true;
}

void Sidebar::DispatchSearchRequest() {
    if (parent_hwnd_ == nullptr) {
        return;
    }

    std::wstring query = TrimWhitespace(SearchText());
    if (query.empty()) {
        DispatchSearchClear();
        return;
    }

    auto payload = std::make_unique<std::wstring>(std::move(query));
    if (!PostMessageW(
            parent_hwnd_,
            WM_FE_SIDEBAR_SEARCH_REQUEST,
            0,
            reinterpret_cast<LPARAM>(payload.get()))) {
        LogLastError(L"PostMessageW(WM_FE_SIDEBAR_SEARCH_REQUEST)");
        return;
    }
    payload.release();
}

void Sidebar::DispatchSearchClear() {
    if (parent_hwnd_ == nullptr) {
        return;
    }

    if (!PostMessageW(parent_hwnd_, WM_FE_SIDEBAR_SEARCH_CLEAR, 0, 0)) {
        LogLastError(L"PostMessageW(WM_FE_SIDEBAR_SEARCH_CLEAR)");
    }
}

std::wstring Sidebar::SearchText() const {
    if (search_edit_hwnd_ == nullptr) {
        return L"";
    }

    const int text_length = GetWindowTextLengthW(search_edit_hwnd_);
    if (text_length <= 0) {
        return L"";
    }

    std::vector<wchar_t> buffer(static_cast<size_t>(text_length) + 1, L'\0');
    if (GetWindowTextW(search_edit_hwnd_, buffer.data(), text_length + 1) <= 0) {
        return L"";
    }
    return std::wstring(buffer.data());
}

bool Sidebar::SearchHasText() const {
    return !SearchText().empty();
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

    std::wstring launch_target = normalized;
    bool launch_target_is_folder = false;
    if (TryResolveShortcutTarget(normalized, &launch_target, &launch_target_is_folder)) {
        launch_target = NormalizePath(std::move(launch_target));
        if (launch_target_is_folder) {
            PostNavigatePath(launch_target);
            return;
        }
    }

    const HINSTANCE open_result =
        ShellExecuteW(hwnd_, L"open", launch_target.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    const INT_PTR shell_status = reinterpret_cast<INT_PTR>(open_result);
    if (shell_status <= 32) {
        wchar_t buffer[256] = {};
        swprintf_s(
            buffer,
            L"[FileExplorer] ShellExecuteW(Sidebar OpenPath) failed (status=%lld path=%s).\r\n",
            static_cast<long long>(shell_status),
            launch_target.c_str());
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

    if (flyout_ignore_initial_mouse_release_) {
        flyout_ignore_initial_mouse_release_ = false;
        flyout_menu_left_button_was_down_ = false;
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
    flyout_ignore_initial_mouse_release_ = flyout_menu_left_button_was_down_;
    InstallFlyoutMenuHook();

    POINT popup_point = screen_point;
    MONITORINFO monitor_info = {};
    monitor_info.cbSize = sizeof(monitor_info);
    const HMONITOR monitor = MonitorFromPoint(screen_point, MONITOR_DEFAULTTONEAREST);
    if (monitor != nullptr && GetMonitorInfoW(monitor, &monitor_info)) {
        const int reserved_nested_width = ScaleForDpi(560, dpi_);
        popup_point.x = (std::min)(popup_point.x, monitor_info.rcWork.right - reserved_nested_width);
        popup_point.x = (std::max)(popup_point.x, monitor_info.rcWork.left + ScaleForDpi(8, dpi_));
        popup_point.y = (std::max)(popup_point.y, monitor_info.rcWork.top + ScaleForDpi(8, dpi_));
        popup_point.y = (std::min)(popup_point.y, monitor_info.rcWork.bottom - ScaleForDpi(8, dpi_));
    }

    const UINT selected_command = static_cast<UINT>(TrackPopupMenuEx(
        menu,
        TPM_RIGHTBUTTON | TPM_RETURNCMD,
        popup_point.x,
        popup_point.y,
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
    next_flyout_request_id_ = 1;
    flyout_command_targets_.clear();
    flyout_position_targets_.clear();
    flyout_menu_metadata_.clear();
    flyout_request_menus_.clear();
    flyout_populated_menus_.clear();
    flyout_pending_selection_ = false;
    flyout_pending_target_ = {};
    flyout_menu_left_button_was_down_ = false;
    flyout_ignore_initial_mouse_release_ = false;
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

    FlyoutMenuMetadata metadata = {};
    metadata.path = NormalizePath(base_path);
    metadata.include_top_open_row = include_top_open_row;
    metadata.top_open_label = top_open_label;
    metadata.loading = false;
    metadata.loaded = false;
    metadata.request_id = 0;
    flyout_menu_metadata_[menu] = metadata;

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

    if (include_top_open_row && !AppendMenuW(menu, MF_SEPARATOR, 0, nullptr)) {
        LogLastError(L"AppendMenuW(Sidebar Flyout Separator)");
    }
    if (!AppendMenuW(menu, MF_STRING | MF_GRAYED, 0, L"Loading...")) {
        LogLastError(L"AppendMenuW(Sidebar Flyout Loading)");
    } else {
        appended_any = true;
    }

    flyout_populated_menus_.insert(menu);
    BeginFlyoutMenuLoad(menu);
    return appended_any;
}

bool Sidebar::BeginFlyoutMenuLoad(HMENU menu) {
    auto metadata_it = flyout_menu_metadata_.find(menu);
    if (metadata_it == flyout_menu_metadata_.end()) {
        return false;
    }
    if (metadata_it->second.loading || metadata_it->second.loaded) {
        return true;
    }
    if (hwnd_ == nullptr) {
        return false;
    }

    const uint64_t request_id = next_flyout_request_id_++;
    metadata_it->second.loading = true;
    metadata_it->second.request_id = request_id;
    flyout_request_menus_[request_id] = menu;

    const std::wstring request_path = metadata_it->second.path;
    const HWND target_hwnd = hwnd_;
    std::thread(
        [target_hwnd, request_id, request_path]() {
            auto payload = std::make_unique<Sidebar::FlyoutAsyncResult>();
            payload->request_id = request_id;
            payload->entries = Sidebar::EnumerateFlyoutEntries(request_path);
            if (!PostMessageW(
                    target_hwnd,
                    kSidebarMessageFlyoutLoaded,
                    0,
                    reinterpret_cast<LPARAM>(payload.get()))) {
                LogLastError(L"PostMessageW(Sidebar FlyoutLoaded)");
                return;
            }
            payload.release();
        })
        .detach();
    return true;
}

void Sidebar::ApplyFlyoutMenuEntries(
    HMENU menu,
    const FlyoutMenuMetadata& metadata,
    const std::vector<FlyoutEntry>& entries) {
    if (menu == nullptr) {
        return;
    }

    while (GetMenuItemCount(menu) > 0) {
        if (!DeleteMenu(menu, 0, MF_BYPOSITION)) {
            LogLastError(L"DeleteMenu(Sidebar Flyout Replace)");
            break;
        }
    }

    bool appended_any = false;
    if (metadata.include_top_open_row) {
        const UINT command_id = NextFlyoutCommandId();
        const FlyoutCommandTarget target = {NormalizePath(metadata.path), true};
        flyout_command_targets_[command_id] = target;
        const std::wstring label = metadata.top_open_label.empty() ? L"Open folder" : metadata.top_open_label;
        if (AppendMenuW(menu, MF_STRING, command_id, label.c_str())) {
            const int count = GetMenuItemCount(menu);
            if (count > 0) {
                RememberFlyoutMenuPositionTarget(menu, static_cast<UINT>(count - 1), target);
            }
            appended_any = true;
        } else {
            LogLastError(L"AppendMenuW(Sidebar Flyout OpenRoot Async)");
        }
    }

    if (!entries.empty() && metadata.include_top_open_row) {
        if (!AppendMenuW(menu, MF_SEPARATOR, 0, nullptr)) {
            LogLastError(L"AppendMenuW(Sidebar Flyout Separator Async)");
        }
    }

    if (entries.empty()) {
        if (!metadata.include_top_open_row) {
            if (!AppendMenuW(menu, MF_STRING | MF_GRAYED, 0, L"(Empty)")) {
                LogLastError(L"AppendMenuW(Sidebar Flyout Empty)");
            }
        }
        flyout_populated_menus_.insert(menu);
        return;
    }

    for (const FlyoutEntry& entry : entries) {
        if (entry.is_folder) {
            HMENU submenu = CreatePopupMenu();
            if (submenu == nullptr) {
                LogLastError(L"CreatePopupMenu(Sidebar Flyout Submenu Async)");
                continue;
            }
            ConfigureMenuForNotifications(submenu);
            if (!AppendMenuW(submenu, MF_STRING | MF_GRAYED, 0, L"Loading...")) {
                LogLastError(L"AppendMenuW(Sidebar Flyout Submenu Loading)");
            }

            FlyoutMenuMetadata submenu_metadata = {};
            submenu_metadata.path = entry.path;
            submenu_metadata.include_top_open_row = false;
            submenu_metadata.loading = false;
            submenu_metadata.loaded = false;
            submenu_metadata.request_id = 0;
            flyout_menu_metadata_[submenu] = std::move(submenu_metadata);

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
                LogLastError(L"InsertMenuItemW(Sidebar Flyout Submenu Async)");
                flyout_command_targets_.erase(command_id);
                flyout_menu_metadata_.erase(submenu);
                if (!DestroyMenu(submenu)) {
                    LogLastError(L"DestroyMenu(Sidebar Flyout Submenu Async)");
                }
                continue;
            }

            const int count = GetMenuItemCount(menu);
            if (count > 0) {
                RememberFlyoutMenuPositionTarget(menu, static_cast<UINT>(count - 1), target);
            }
            flyout_populated_menus_.insert(submenu);
            appended_any = true;
            continue;
        }

        const UINT command_id = NextFlyoutCommandId();
        const FlyoutCommandTarget target = {entry.path, false};
        flyout_command_targets_[command_id] = target;
        if (!AppendMenuW(menu, MF_STRING, command_id, entry.name.c_str())) {
            LogLastError(L"AppendMenuW(Sidebar Flyout Entry Async)");
            continue;
        }
        const int count = GetMenuItemCount(menu);
        if (count > 0) {
            RememberFlyoutMenuPositionTarget(menu, static_cast<UINT>(count - 1), target);
        }
        appended_any = true;
    }

    if (!appended_any && !metadata.include_top_open_row) {
        if (!AppendMenuW(menu, MF_STRING | MF_GRAYED, 0, L"(Empty)")) {
            LogLastError(L"AppendMenuW(Sidebar Flyout Empty Async)");
        }
    }
    flyout_populated_menus_.insert(menu);
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

std::wstring Sidebar::JoinPath(const std::wstring& base_path, const std::wstring& child_name) {
    std::wstring combined = NormalizePath(base_path);
    if (!combined.empty() && combined.back() != L'\\') {
        combined.push_back(L'\\');
    }
    combined.append(child_name);
    return NormalizePath(std::move(combined));
}

std::vector<Sidebar::FlyoutEntry> Sidebar::EnumerateFlyoutEntries(const std::wstring& base_path) {
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
