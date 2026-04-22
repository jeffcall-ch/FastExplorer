#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>
#include <CommCtrl.h>

#include <array>
#include <memory>
#include <string>
#include <type_traits>
#include <unordered_set>
#include <vector>

struct IContextMenu;

namespace fileexplorer {

constexpr UINT WM_FE_FILELIST_NAVIGATE = WM_APP + 103;
constexpr UINT WM_FE_FILELIST_STATUS_UPDATE = WM_APP + 104;
constexpr UINT WM_FE_FILELIST_REFRESH = WM_APP + 105;
constexpr UINT WM_FE_FILELIST_OPEN_LOCATION_NEW_TAB = WM_APP + 109;
constexpr UINT WM_FE_FILELIST_ADD_REGULAR_FAVOURITE = WM_APP + 110;
constexpr UINT WM_FE_FILELIST_ADD_FLYOUT_FAVOURITE = WM_APP + 111;
constexpr UINT WM_FE_FILELIST_SORT_CHANGED = WM_APP + 112;
constexpr UINT WM_FE_FILELIST_TOGGLE_SHOW_HIDDEN = WM_APP + 113;
constexpr UINT WM_FE_FILELIST_TOGGLE_SHOW_EXTENSIONS = WM_APP + 114;

enum class SortColumn {
    Name = 0,
    Extension = 1,
    DateModified = 2,
    Size = 3,
    Path = 4,
};

enum class SortDirection {
    Ascending,
    Descending,
};

struct FileEntry {
    std::wstring name;
    std::wstring extension;
    std::wstring full_path;
    FILETIME modified_time{};
    ULONGLONG size_bytes = 0;
    DWORD attributes = 0;
    bool is_folder = false;
    int icon_index = 0;

    bool is_group_divider = false;
    std::wstring group_label;
};

class FileListView final {
public:
    static constexpr int kColumnCount = 5;
    using ColumnWidthsLogical = std::array<int, kColumnCount>;

    struct ViewStateSnapshot {
        std::vector<std::wstring> selected_names;
        std::wstring focused_name;
        int top_index = 0;
    };

    struct SearchSnapshot {
        std::wstring root_path;
        std::wstring pattern;
        ULONGLONG elapsed_ms = 0;
        std::vector<FileEntry> entries;
        SortColumn sort_column = SortColumn::Name;
        SortDirection sort_direction = SortDirection::Ascending;
    };

    FileListView();
    ~FileListView();

    FileListView(const FileListView&) = delete;
    FileListView& operator=(const FileListView&) = delete;

    bool Create(HWND parent, HINSTANCE instance, int control_id);
    HWND hwnd() const noexcept;
    void SetDpi(UINT dpi);
    void SetPostLoadSelectionName(std::wstring name);

    void BeginFolderLoad(const std::wstring& path, bool incremental_refresh);
    void ApplyLoadedFolderEntries(std::vector<FileEntry> entries, bool incremental_refresh);
    void BeginSearch(const std::wstring& root_path, const std::wstring& pattern);
    void AppendSearchResults(std::vector<FileEntry> entries);
    void SetSearchElapsedMs(ULONGLONG elapsed_ms);
    void CompleteSearch();
    void LeaveSearchMode();
    bool IsSearchMode() const noexcept;
    bool CaptureSearchSnapshot(SearchSnapshot* snapshot) const;
    bool RestoreSearchSnapshot(const SearchSnapshot& snapshot);
    bool CaptureViewStateSnapshot(ViewStateSnapshot* snapshot) const;
    void SetPendingViewStateSnapshot(const ViewStateSnapshot* snapshot);
    void ClearPendingViewStateSnapshot();

    bool LoadFolder(const std::wstring& path);
    const std::wstring& current_path() const noexcept;
    void SetShowHiddenFiles(bool show_hidden_files);
    bool show_hidden_files() const noexcept;
    void SetShowExtensions(bool show_extensions);
    bool show_extensions() const noexcept;
    void SetSortOrder(SortColumn sort_column, SortDirection sort_direction);
    SortColumn sort_column() const noexcept;
    SortDirection sort_direction() const noexcept;
    void SetColumnWidthsLogical(const ColumnWidthsLogical& widths);
    ColumnWidthsLogical GetColumnWidthsLogical() const;
    static ColumnWidthsLogical DefaultColumnWidthsLogical();

    bool HandleNotify(LPARAM l_param, LRESULT* result);

private:
    enum class DateBucket {
        Today,
        Yesterday,
        ThisWeek,
        ThisMonth,
        Older,
    };

    struct FontDeleter {
        void operator()(HFONT font) const noexcept;
    };

    using UniqueFont = std::unique_ptr<std::remove_pointer_t<HFONT>, FontDeleter>;

    static LRESULT CALLBACK ListViewSubclassProc(
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
    static LRESULT CALLBACK HeaderEraseSubclassProc(
        HWND hwnd,
        UINT message,
        WPARAM w_param,
        LPARAM l_param,
        UINT_PTR subclass_id,
        DWORD_PTR ref_data);

    bool CreateListViewControl();
    void CreateColumns();
    void ApplyScaledColumns();
    void ApplyColumnMode();
    void ApplyHeaderVisualStyle();
    void EnsureSystemImageList();
    void EnsureUiFonts();
    void EnsureDividerFont();

    bool HandleGetDispInfo(NMLVDISPINFOW* info) const;
    bool HandleColumnClick(int column_index);
    bool HandleCustomDraw(NMLVCUSTOMDRAW* custom_draw, LRESULT* result);
    bool HandleHeaderCustomDraw(NMCUSTOMDRAW* custom_draw, LRESULT* result) const;
    bool HandleDoubleClick(const NMITEMACTIVATE* item_activate);
    bool HandleItemChanging(const NMLISTVIEW* list_view, LRESULT* result) const;
    bool HandleItemChanged(const NMLISTVIEW* list_view);
    bool HandleClick(const NMITEMACTIVATE* item_activate);

    LRESULT HandleSubclassMessage(UINT message, WPARAM w_param, LPARAM l_param);

    void SortBaseEntries();
    void BuildDisplayEntriesWithGroups();
    void UpdateSortIndicators();
    void UpdateItemCountAndRefresh(bool set_default_focus);
    void PostStatusUpdate() const;
    std::wstring BuildStatusText() const;
    void SetLoadingState(bool loading);
    void DrawLoadingOverlay(HDC hdc) const;
    void CaptureViewState(std::vector<std::wstring>* selected_names, std::wstring* focused_name, int* top_index) const;
    void RestoreViewState(const std::vector<std::wstring>& selected_names, const std::wstring& focused_name, int top_index);
    void ApplyRefreshDelta(std::vector<FileEntry> incoming_entries);

    void ClearSelection();
    void SanitizeSelectionAndFocus();
    bool NavigateToParentFolder();
    bool HandleTypeAheadChar(wchar_t character);
    bool SelectByPrefix(const std::wstring& prefix);
    int FindNextSelectableIndex(int start_index, int step) const;
    bool IsSelectableIndex(int index) const;
    bool IsDividerIndex(int index) const;
    int FindSelectableIndexByName(const std::wstring& name) const;
    bool SelectSingleEntryByName(const std::wstring& name);
    int SelectedCountExcludingDividers() const;
    int FocusedIndex() const;
    bool OpenEntryAtIndex(int index, bool open_files_too);
    void ShowContextMenu(POINT screen_point, int hit_index);
    void DrawEmptyState(HDC hdc) const;
    int ResolveContextTargetIndex(int hit_index) const;
    std::vector<int> CollectSelectedSelectableIndices() const;
    std::vector<std::wstring> CollectPathsForIndices(const std::vector<int>& indices) const;
    void EnsureSingleSelectionAtIndex(int index);
    bool CopySelectionToClipboard(bool cut, int fallback_index);
    bool PasteFromClipboard();
    bool DeleteSelection(bool permanent, int fallback_index);
    bool BeginInlineRename(int index);
    bool CommitInlineRename();
    void CancelInlineRename();
    void EndInlineRenameControl();
    bool ApplyInlineRenameChange(const std::wstring& old_full_path, const std::wstring& new_name);
    bool CreateNewFolderAndBeginRename();
    std::wstring BuildUniqueNewFolderPath() const;
    void CaptureCurrentColumnWidths();
    void RequestRefresh();
    bool AppendShellContextMenu(HMENU menu, const std::wstring& path, UINT first_id, UINT last_id);
    bool InvokeShellContextMenu(UINT command_id);
    void ClearShellContextMenu();
    void ClearCutPendingState();
    void SetCutPendingPaths(const std::vector<std::wstring>& paths);
    void RemoveCutPendingPaths(const std::vector<std::wstring>& paths);
    bool IsPathCutPending(const std::wstring& full_path) const;
    std::wstring BuildDisplayName(const FileEntry& entry) const;

    std::wstring BuildFullPath(const FileEntry& entry) const;
    static std::wstring NormalizePath(std::wstring path);
    static std::wstring ParentPath(const std::wstring& path);
    static std::wstring FormatWin32ErrorMessage(DWORD error_code);
    static std::wstring EntryKey(const FileEntry& entry);
    static bool EntriesEquivalent(const FileEntry& left, const FileEntry& right);
    static std::wstring GetExtensionFromName(const std::wstring& name);
    static int CompareTextInsensitive(const std::wstring& left, const std::wstring& right);
    static std::wstring FormatDateTime(const FILETIME& file_time);
    static std::wstring FormatFileSize(ULONGLONG size_bytes);
    static int ResolveIconIndex(const std::wstring& full_path, bool is_folder);
    static COLORREF ColorForExtension(const std::wstring& extension, bool is_folder);
    static int ClampColumnWidthLogical(int column_index, int value) noexcept;
    static DateBucket BucketForFileTime(const FILETIME& file_time);
    static const wchar_t* LabelForBucket(DateBucket bucket);

    HWND parent_hwnd_{nullptr};
    HWND hwnd_{nullptr};
    HWND header_hwnd_{nullptr};
    HINSTANCE instance_{nullptr};
    int control_id_{0};
    UINT dpi_{96U};
    HIMAGELIST system_image_list_{nullptr};

    std::wstring current_path_{};
    std::vector<FileEntry> base_entries_{};
    std::vector<FileEntry> display_entries_{};
    ColumnWidthsLogical column_widths_logical_{DefaultColumnWidthsLogical()};

    SortColumn sort_column_{SortColumn::Name};
    SortDirection sort_direction_{SortDirection::Ascending};

    UniqueFont divider_font_{nullptr};
    int divider_font_height_{0};
    UniqueFont list_font_{nullptr};
    int list_font_height_{0};
    UniqueFont header_font_{nullptr};
    int header_font_height_{0};

    std::wstring type_ahead_buffer_{};
    DWORD type_ahead_last_tick_{0};
    bool sanitizing_selection_{false};
    bool loading_{false};
    bool pending_incremental_refresh_{false};
    int loading_frame_{0};
    std::wstring post_load_selection_name_{};
    bool search_mode_{false};
    std::wstring search_root_path_{};
    std::wstring search_pattern_{};
    ULONGLONG search_elapsed_ms_{0};
    bool has_pending_view_state_snapshot_{false};
    ViewStateSnapshot pending_view_state_snapshot_{};
    bool show_hidden_files_{false};
    bool show_extensions_{true};

    HWND rename_edit_hwnd_{nullptr};
    int rename_item_index_{-1};
    std::wstring rename_original_name_{};
    std::wstring rename_original_full_path_{};
    IContextMenu* shell_context_menu_{nullptr};
    UINT shell_context_menu_first_id_{0};
    UINT shell_context_menu_last_id_{0};
    std::unordered_set<std::wstring> cut_pending_paths_{};
};

}  // namespace fileexplorer
