#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>
#include <CommCtrl.h>

#include <cstdint>
#include <memory>
#include <string>
#include <type_traits>
#include <unordered_map>
#include <unordered_set>
#include <vector>

#include "FavouritesStore.h"

namespace fileexplorer {

constexpr UINT WM_FE_SIDEBAR_NAVIGATE = WM_APP + 106;
constexpr UINT WM_FE_SIDEBAR_SEARCH_REQUEST = WM_APP + 107;
constexpr UINT WM_FE_SIDEBAR_SEARCH_CLEAR = WM_APP + 108;

class Sidebar final {
public:
    Sidebar();
    ~Sidebar();

    Sidebar(const Sidebar&) = delete;
    Sidebar& operator=(const Sidebar&) = delete;

    bool Create(HWND parent, HINSTANCE instance, int control_id);
    HWND hwnd() const noexcept;

    void SetDpi(UINT dpi);
    void SetCurrentPath(const std::wstring& path);
    void FocusSearch();
    void SetSearchText(const std::wstring& text);
    void ClearSearchText(bool notify_parent);
    bool AddRegularFavourite(const std::wstring& path);
    bool AddFlyoutFavourite(const std::wstring& path);

private:
    enum class Section {
        None,
        Flyout,
        Favourites,
    };

    struct SidebarItem {
        std::wstring name;
        std::wstring path;
        bool show_arrow = false;
    };

    struct FlyoutEntry {
        std::wstring name;
        std::wstring path;
        bool is_folder = false;
    };

    struct FlyoutCommandTarget {
        std::wstring path;
        bool is_folder = false;
    };

    struct FlyoutMenuMetadata {
        std::wstring path;
        bool include_top_open_row = false;
        std::wstring top_open_label;
        bool loading = false;
        bool loaded = false;
        uint64_t request_id = 0;
    };

    struct FlyoutAsyncResult {
        uint64_t request_id = 0;
        std::vector<FlyoutEntry> entries;
    };

    struct LayoutInfo {
        RECT search_rect{};
        int content_top = 0;
        int content_bottom = 0;
        int flyout_header_logical_top = 0;
        int flyout_items_logical_top = 0;
        int divider_logical_y = 0;
        int favourites_header_logical_top = 0;
        int favourites_items_logical_top = 0;
        int total_logical_height = 0;
        int row_height = 0;
    };

    struct FontDeleter {
        void operator()(HFONT font) const noexcept;
    };

    struct BrushDeleter {
        void operator()(HBRUSH brush) const noexcept;
    };

    using UniqueFont = std::unique_ptr<std::remove_pointer_t<HFONT>, FontDeleter>;
    using UniqueBrush = std::unique_ptr<std::remove_pointer_t<HBRUSH>, BrushDeleter>;

    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM w_param, LPARAM l_param);
    static LRESULT CALLBACK SearchEditSubclassProc(
        HWND hwnd,
        UINT message,
        WPARAM w_param,
        LPARAM l_param,
        UINT_PTR subclass_id,
        DWORD_PTR ref_data);
    static LRESULT CALLBACK RenameEditSubclassProc(
        HWND hwnd,
        UINT message,
        WPARAM w_param,
        LPARAM l_param,
        UINT_PTR subclass_id,
        DWORD_PTR ref_data);
    LRESULT HandleMessage(UINT message, WPARAM w_param, LPARAM l_param);

    bool RegisterWindowClass();
    void EnsureFonts();
    bool CreateSearchEditControl();
    void UpdateSearchEditLayout(const LayoutInfo& layout);
    void BuildDefaultItems();
    void RefreshItemsFromStore();
    void Paint(HDC hdc);
    LayoutInfo BuildLayout(const RECT& client_rect) const;
    void UpdateScrollInfo();
    int MaxScrollOffset(const LayoutInfo& layout) const;
    void SetScrollOffset(int offset);
    bool HitTestItem(POINT point, Section* section, int* item_index) const;
    bool IsItemActive(const SidebarItem& item) const;
    void ShowItemContextMenu(POINT screen_point, Section section, int item_index);
    void HandleItemInvoke(Section section, int item_index, POINT screen_point);
    bool RemoveItem(Section section, int item_index);
    bool BeginRenameItem(Section section, int item_index);
    void EndRenameItemEdit(bool commit);
    bool ResolveItemTextRect(Section section, int item_index, RECT* rect) const;
    void DispatchSearchRequest();
    void DispatchSearchClear();
    std::wstring SearchText() const;
    bool SearchHasText() const;
    void PostNavigatePath(const std::wstring& path) const;
    void OpenPathOrNavigate(const std::wstring& path, bool is_folder) const;
    void ShowFlyoutPopup(const SidebarItem& item, POINT screen_point);
    void ClearFlyoutPopupState();
    void InstallFlyoutMenuHook();
    void RemoveFlyoutMenuHook();
    bool HandleFlyoutMenuHookMessage(const MSG& msg);
    bool ResolveFlyoutPopupParentFromPoint(POINT screen_point, FlyoutCommandTarget* target) const;
    static LRESULT CALLBACK MenuMsgFilterProc(int code, WPARAM w_param, LPARAM l_param);
    static UINT64 MakeFlyoutMenuPositionKey(HMENU menu, UINT item_position);
    void RememberFlyoutMenuPositionTarget(HMENU menu, UINT item_position, const FlyoutCommandTarget& target);
    bool PopulateFlyoutMenu(
        HMENU menu,
        const std::wstring& base_path,
        bool include_top_open_row,
        const std::wstring& top_open_label);
    void ApplyFlyoutMenuEntries(HMENU menu, const FlyoutMenuMetadata& metadata, const std::vector<FlyoutEntry>& entries);
    bool BeginFlyoutMenuLoad(HMENU menu);
    UINT NextFlyoutCommandId();
    static std::wstring JoinPath(const std::wstring& base_path, const std::wstring& child_name);
    static std::vector<FlyoutEntry> EnumerateFlyoutEntries(const std::wstring& base_path);
    int MeasuredItemTextHeight() const;

    static std::wstring NormalizePath(std::wstring path);
    static std::wstring ComparePathNormalized(std::wstring path);
    static std::wstring UserProfilePath();
    static std::wstring BuildUserChildPath(const wchar_t* child_name);
    static int CompareTextInsensitive(const std::wstring& left, const std::wstring& right);

    HWND parent_hwnd_{nullptr};
    HWND hwnd_{nullptr};
    HWND search_edit_hwnd_{nullptr};
    HWND rename_edit_hwnd_{nullptr};
    HINSTANCE instance_{nullptr};
    int control_id_{0};
    UINT dpi_{96U};

    FavouritesStore favourites_store_{};
    std::wstring current_path_normalized_{};
    std::vector<SidebarItem> flyout_items_{};
    std::vector<SidebarItem> favourite_items_{};

    Section hovered_section_{Section::None};
    int hovered_index_{-1};
    bool tracking_mouse_{false};
    int scroll_offset_{0};

    Section rename_section_{Section::None};
    int rename_item_index_{-1};
    bool rename_edit_closing_{false};

    bool flyout_popup_active_{false};
    UINT next_flyout_command_id_{9000};
    uint64_t next_flyout_request_id_{1};
    std::unordered_map<UINT, FlyoutCommandTarget> flyout_command_targets_{};
    std::unordered_map<UINT64, FlyoutCommandTarget> flyout_position_targets_{};
    std::unordered_map<HMENU, FlyoutMenuMetadata> flyout_menu_metadata_{};
    std::unordered_map<uint64_t, HMENU> flyout_request_menus_{};
    std::unordered_set<HMENU> flyout_populated_menus_{};
    bool flyout_pending_selection_{false};
    FlyoutCommandTarget flyout_pending_target_{};
    bool flyout_menu_left_button_was_down_{false};
    bool flyout_ignore_initial_mouse_release_{false};
    HHOOK flyout_menu_hook_{nullptr};
    RECT clear_button_rect_{};
    bool clear_button_visible_{false};

    UniqueFont header_font_{nullptr};
    UniqueFont item_font_{nullptr};
    UniqueBrush search_edit_brush_{nullptr};
    int header_font_height_{0};
    int item_font_height_{0};
};

}  // namespace fileexplorer
