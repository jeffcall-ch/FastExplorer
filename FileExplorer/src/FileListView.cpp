#include "FileListView.h"

#include <Shellapi.h>
#include <Shlwapi.h>
#include <Windowsx.h>
#include <strsafe.h>
#include <uxtheme.h>

#include <algorithm>
#include <cwchar>
#include <cwctype>
#include <memory>
#include <unordered_map>
#include <unordered_set>
#include <utility>

namespace {

constexpr UINT_PTR kListViewSubclassId = 1;
constexpr UINT_PTR kLoadingTimerId = 9001;

constexpr COLORREF kListBackgroundColor = RGB(0x18, 0x1C, 0x22);
constexpr COLORREF kListTextColor = RGB(0xE2, 0xE8, 0xEF);
constexpr COLORREF kHeaderBackgroundColor = RGB(0x1F, 0x25, 0x2D);
constexpr COLORREF kHeaderHoverBackgroundColor = RGB(0x24, 0x2B, 0x35);
constexpr COLORREF kHeaderTextColor = RGB(0xD2, 0xDA, 0xE4);
constexpr COLORREF kHeaderDividerColor = RGB(0x33, 0x3B, 0x47);
constexpr COLORREF kHeaderSortGlyphColor = RGB(0xA9, 0xB4, 0xC2);
constexpr COLORREF kDividerBackgroundColor = RGB(0x1E, 0x24, 0x2B);
constexpr COLORREF kDividerTextColor = RGB(0xAC, 0xB6, 0xC2);
constexpr COLORREF kEmptyStateTextColor = RGB(0x9F, 0xA8, 0xB5);
constexpr COLORREF kLoadingOverlayBackgroundColor = RGB(0x24, 0x2A, 0x33);
constexpr COLORREF kLoadingOverlayBorderColor = RGB(0x3A, 0x43, 0x4F);
constexpr COLORREF kLoadingOverlayTextColor = RGB(0xDF, 0xE5, 0xEE);

constexpr UINT kMenuCopy = 7001;
constexpr UINT kMenuCut = 7002;
constexpr UINT kMenuPaste = 7003;
constexpr UINT kMenuRename = 7004;
constexpr UINT kMenuDelete = 7005;
constexpr UINT kMenuProperties = 7006;
constexpr UINT kMenuCopyPath = 7007;
constexpr UINT kMenuMakeFavourite = 7008;
constexpr UINT kMenuMakeFlyoutFavourite = 7009;
constexpr UINT kMenuOpenFileLocation = 7010;
constexpr DWORD kTypeAheadResetMs = 1200;
constexpr UINT kLoadingTimerMs = 120;

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

std::wstring ToLowerCopy(std::wstring value) {
    std::transform(
        value.begin(),
        value.end(),
        value.begin(),
        [](wchar_t ch) { return static_cast<wchar_t>(towlower(ch)); });
    return value;
}

ULONGLONG FileTimeToUint64(const FILETIME& value) {
    ULARGE_INTEGER ui = {};
    ui.LowPart = value.dwLowDateTime;
    ui.HighPart = value.dwHighDateTime;
    return ui.QuadPart;
}

FILETIME Uint64ToFileTime(ULONGLONG value) {
    ULARGE_INTEGER ui = {};
    ui.QuadPart = value;
    FILETIME file_time = {};
    file_time.dwLowDateTime = ui.LowPart;
    file_time.dwHighDateTime = ui.HighPart;
    return file_time;
}

bool SameDate(const SYSTEMTIME& left, const SYSTEMTIME& right) {
    return left.wYear == right.wYear && left.wMonth == right.wMonth && left.wDay == right.wDay;
}

bool StartsWithInsensitive(const std::wstring& text, const std::wstring& prefix) {
    if (prefix.empty()) {
        return true;
    }
    if (prefix.size() > text.size()) {
        return false;
    }

    return CompareStringOrdinal(
               text.c_str(),
               static_cast<int>(prefix.size()),
               prefix.c_str(),
               static_cast<int>(prefix.size()),
               TRUE) == CSTR_EQUAL;
}

bool CopyTextToClipboard(HWND owner, const std::wstring& text) {
    if (!OpenClipboard(owner)) {
        LogLastError(L"OpenClipboard");
        return false;
    }

    if (!EmptyClipboard()) {
        LogLastError(L"EmptyClipboard");
        CloseClipboard();
        return false;
    }

    const size_t bytes = (text.size() + 1) * sizeof(wchar_t);
    HGLOBAL global_handle = GlobalAlloc(GMEM_MOVEABLE, bytes);
    if (global_handle == nullptr) {
        LogLastError(L"GlobalAlloc(Clipboard)");
        CloseClipboard();
        return false;
    }

    void* destination = GlobalLock(global_handle);
    if (destination == nullptr) {
        LogLastError(L"GlobalLock(Clipboard)");
        GlobalFree(global_handle);
        CloseClipboard();
        return false;
    }

    memcpy(destination, text.c_str(), bytes);
    GlobalUnlock(global_handle);

    if (SetClipboardData(CF_UNICODETEXT, global_handle) == nullptr) {
        LogLastError(L"SetClipboardData");
        GlobalFree(global_handle);
        CloseClipboard();
        return false;
    }

    CloseClipboard();
    return true;
}

}  // namespace

namespace fileexplorer {

void FileListView::FontDeleter::operator()(HFONT font) const noexcept {
    if (font != nullptr) {
        DeleteObject(font);
    }
}

FileListView::FileListView() = default;

FileListView::~FileListView() {
    if (hwnd_ != nullptr) {
        RemoveWindowSubclass(hwnd_, &FileListView::ListViewSubclassProc, kListViewSubclassId);
    }
}

bool FileListView::Create(HWND parent, HINSTANCE instance, int control_id) {
    parent_hwnd_ = parent;
    instance_ = instance;
    control_id_ = control_id;

    if (!CreateListViewControl()) {
        return false;
    }

    CreateColumns();
    ApplyColumnMode();
    ApplyHeaderVisualStyle();
    EnsureSystemImageList();
    EnsureDividerFont();
    UpdateSortIndicators();
    PostStatusUpdate();
    return true;
}

HWND FileListView::hwnd() const noexcept {
    return hwnd_;
}

void FileListView::SetDpi(UINT dpi) {
    dpi_ = (dpi == 0U) ? 96U : dpi;
    divider_font_.reset();
    divider_font_height_ = 0;
    EnsureDividerFont();
    ApplyColumnMode();
    ApplyHeaderVisualStyle();
    InvalidateRect(hwnd_, nullptr, TRUE);
}

void FileListView::SetPostLoadSelectionName(std::wstring name) {
    post_load_selection_name_ = std::move(name);
}

void FileListView::BeginFolderLoad(const std::wstring& path, bool incremental_refresh) {
    if (search_mode_) {
        search_mode_ = false;
        search_root_path_.clear();
        search_pattern_.clear();
        ApplyColumnMode();
    }

    current_path_ = NormalizePath(path);
    pending_incremental_refresh_ = incremental_refresh;
    type_ahead_buffer_.clear();
    type_ahead_last_tick_ = 0;

    SetLoadingState(true);
    PostStatusUpdate();
}

void FileListView::ApplyLoadedFolderEntries(std::vector<FileEntry> entries, bool incremental_refresh) {
    std::vector<std::wstring> selected_names;
    std::wstring focused_name;
    int top_index = (hwnd_ != nullptr) ? ListView_GetTopIndex(hwnd_) : 0;
    if (incremental_refresh) {
        CaptureViewState(&selected_names, &focused_name, &top_index);
        ApplyRefreshDelta(std::move(entries));
    } else {
        base_entries_ = std::move(entries);
    }

    SortBaseEntries();
    BuildDisplayEntriesWithGroups();
    UpdateSortIndicators();
    UpdateItemCountAndRefresh(!incremental_refresh);
    if (incremental_refresh) {
        RestoreViewState(selected_names, focused_name, top_index);
    } else if (!post_load_selection_name_.empty()) {
        SelectSingleEntryByName(post_load_selection_name_);
    }
    post_load_selection_name_.clear();

    SetLoadingState(false);
    pending_incremental_refresh_ = false;
    PostStatusUpdate();
}

void FileListView::BeginSearch(const std::wstring& root_path, const std::wstring& pattern) {
    search_mode_ = true;
    search_root_path_ = NormalizePath(root_path);
    search_pattern_ = pattern;
    search_elapsed_ms_ = 0;
    current_path_ = search_root_path_;
    pending_incremental_refresh_ = false;
    post_load_selection_name_.clear();
    type_ahead_buffer_.clear();
    type_ahead_last_tick_ = 0;

    base_entries_.clear();
    display_entries_.clear();
    sort_column_ = SortColumn::Name;
    sort_direction_ = SortDirection::Ascending;

    ApplyColumnMode();
    UpdateSortIndicators();
    UpdateItemCountAndRefresh(true);
    SetLoadingState(true);
    PostStatusUpdate();
}

void FileListView::AppendSearchResults(std::vector<FileEntry> entries) {
    if (!search_mode_ || entries.empty()) {
        return;
    }

    for (FileEntry& entry : entries) {
        if (entry.full_path.empty()) {
            entry.full_path = BuildFullPath(entry);
        }
        if (!entry.is_folder && entry.extension.empty()) {
            entry.extension = GetExtensionFromName(entry.name);
        }
        base_entries_.push_back(std::move(entry));
    }

    SortBaseEntries();
    BuildDisplayEntriesWithGroups();

    const bool set_default_focus = FocusedIndex() < 0;
    UpdateItemCountAndRefresh(set_default_focus);
    PostStatusUpdate();
}

void FileListView::SetSearchElapsedMs(ULONGLONG elapsed_ms) {
    search_elapsed_ms_ = elapsed_ms;
    if (search_mode_) {
        PostStatusUpdate();
    }
}

void FileListView::CompleteSearch() {
    if (!search_mode_) {
        return;
    }

    SetLoadingState(false);
    pending_incremental_refresh_ = false;
    PostStatusUpdate();
}

void FileListView::LeaveSearchMode() {
    if (!search_mode_) {
        return;
    }

    search_mode_ = false;
    search_root_path_.clear();
    search_pattern_.clear();
    search_elapsed_ms_ = 0;
    if (sort_column_ == SortColumn::Path) {
        sort_column_ = SortColumn::Name;
        sort_direction_ = SortDirection::Ascending;
    }
    SetLoadingState(false);
    ApplyColumnMode();
    UpdateSortIndicators();
}

bool FileListView::IsSearchMode() const noexcept {
    return search_mode_;
}

bool FileListView::CaptureSearchSnapshot(SearchSnapshot* snapshot) const {
    if (snapshot == nullptr || !search_mode_) {
        return false;
    }

    snapshot->root_path = search_root_path_;
    snapshot->pattern = search_pattern_;
    snapshot->elapsed_ms = search_elapsed_ms_;
    snapshot->entries = base_entries_;
    snapshot->sort_column = sort_column_;
    snapshot->sort_direction = sort_direction_;
    return true;
}

bool FileListView::RestoreSearchSnapshot(const SearchSnapshot& snapshot) {
    if (snapshot.root_path.empty()) {
        return false;
    }

    search_mode_ = true;
    search_root_path_ = NormalizePath(snapshot.root_path);
    search_pattern_ = snapshot.pattern;
    search_elapsed_ms_ = snapshot.elapsed_ms;
    current_path_ = search_root_path_;
    pending_incremental_refresh_ = false;
    post_load_selection_name_.clear();
    type_ahead_buffer_.clear();
    type_ahead_last_tick_ = 0;

    base_entries_ = snapshot.entries;
    sort_column_ = snapshot.sort_column;
    sort_direction_ = snapshot.sort_direction;

    ApplyColumnMode();
    SortBaseEntries();
    BuildDisplayEntriesWithGroups();
    UpdateSortIndicators();
    UpdateItemCountAndRefresh(true);
    SetLoadingState(false);
    PostStatusUpdate();
    return true;
}

bool FileListView::LoadFolder(const std::wstring& path) {
    BeginFolderLoad(path, false);
    base_entries_.clear();
    display_entries_.clear();

    if (current_path_.empty()) {
        UpdateItemCountAndRefresh(true);
        SetLoadingState(false);
        PostStatusUpdate();
        return false;
    }

    std::wstring wildcard = current_path_;
    if (wildcard.back() != L'\\') {
        wildcard.push_back(L'\\');
    }
    wildcard.push_back(L'*');

    bool loaded = true;
    WIN32_FIND_DATAW find_data = {};
    HANDLE find_handle = FindFirstFileW(wildcard.c_str(), &find_data);
    if (find_handle != INVALID_HANDLE_VALUE) {
        do {
            const wchar_t* name = find_data.cFileName;
            if (name[0] == L'.' && (name[1] == L'\0' || (name[1] == L'.' && name[2] == L'\0'))) {
                continue;
            }

            if ((find_data.dwFileAttributes & (FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM)) != 0) {
                continue;
            }

            FileEntry entry = {};
            entry.name = name;
            entry.attributes = find_data.dwFileAttributes;
            entry.is_folder = (find_data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
            entry.modified_time = find_data.ftLastWriteTime;
            entry.size_bytes =
                (static_cast<ULONGLONG>(find_data.nFileSizeHigh) << 32ULL) |
                static_cast<ULONGLONG>(find_data.nFileSizeLow);
            if (entry.is_folder) {
                entry.size_bytes = 0;
            } else {
                entry.extension = GetExtensionFromName(entry.name);
            }

            const std::wstring full_path = BuildFullPath(entry);
            entry.icon_index = ResolveIconIndex(full_path, entry.is_folder);
            base_entries_.push_back(std::move(entry));
        } while (FindNextFileW(find_handle, &find_data));

        FindClose(find_handle);
    } else {
        loaded = false;
    }

    SortBaseEntries();
    BuildDisplayEntriesWithGroups();
    UpdateItemCountAndRefresh(true);
    if (!post_load_selection_name_.empty()) {
        SelectSingleEntryByName(post_load_selection_name_);
        post_load_selection_name_.clear();
    }
    SetLoadingState(false);
    PostStatusUpdate();
    return loaded;
}

const std::wstring& FileListView::current_path() const noexcept {
    return current_path_;
}

bool FileListView::HandleNotify(LPARAM l_param, LRESULT* result) {
    auto* header = reinterpret_cast<NMHDR*>(l_param);
    if (header == nullptr) {
        return false;
    }

    HWND list_header = (hwnd_ != nullptr) ? ListView_GetHeader(hwnd_) : nullptr;
    if (list_header != nullptr && header->hwndFrom == list_header && header->code == NM_CUSTOMDRAW) {
        return HandleHeaderCustomDraw(reinterpret_cast<NMCUSTOMDRAW*>(l_param), result);
    }

    if (header->hwndFrom != hwnd_) {
        return false;
    }

    switch (header->code) {
    case LVN_GETDISPINFOW:
        HandleGetDispInfo(reinterpret_cast<NMLVDISPINFOW*>(l_param));
        if (result != nullptr) {
            *result = 0;
        }
        return true;

    case LVN_COLUMNCLICK:
        HandleColumnClick(reinterpret_cast<NMLISTVIEW*>(l_param)->iSubItem);
        if (result != nullptr) {
            *result = 0;
        }
        return true;

    case NM_DBLCLK:
        HandleDoubleClick(reinterpret_cast<NMITEMACTIVATE*>(l_param));
        if (result != nullptr) {
            *result = 0;
        }
        return true;

    case NM_CUSTOMDRAW:
        return HandleCustomDraw(reinterpret_cast<NMLVCUSTOMDRAW*>(l_param), result);

    case LVN_ITEMCHANGING:
        return HandleItemChanging(reinterpret_cast<NMLISTVIEW*>(l_param), result);

    case LVN_ITEMCHANGED:
        if (HandleItemChanged(reinterpret_cast<NMLISTVIEW*>(l_param))) {
            if (result != nullptr) {
                *result = 0;
            }
            return true;
        }
        break;

    case NM_CLICK:
        if (HandleClick(reinterpret_cast<NMITEMACTIVATE*>(l_param))) {
            if (result != nullptr) {
                *result = 0;
            }
            return true;
        }
        break;

    default:
        break;
    }

    return false;
}

LRESULT CALLBACK FileListView::ListViewSubclassProc(
    HWND hwnd,
    UINT message,
    WPARAM w_param,
    LPARAM l_param,
    UINT_PTR subclass_id,
    DWORD_PTR ref_data) {
    (void)subclass_id;

    auto* self = reinterpret_cast<FileListView*>(ref_data);
    if (self == nullptr) {
        return DefSubclassProc(hwnd, message, w_param, l_param);
    }

    return self->HandleSubclassMessage(message, w_param, l_param);
}

bool FileListView::CreateListViewControl() {
    constexpr DWORD style =
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_OWNERDATA | LVS_SHOWSELALWAYS;

    hwnd_ = CreateWindowExW(
        0,
        WC_LISTVIEWW,
        L"",
        style,
        0,
        0,
        0,
        0,
        parent_hwnd_,
        reinterpret_cast<HMENU>(static_cast<INT_PTR>(control_id_)),
        instance_,
        nullptr);
    if (hwnd_ == nullptr) {
        LogLastError(L"CreateWindowExW(FileListView)");
        return false;
    }

    if (!SetWindowSubclass(hwnd_, &FileListView::ListViewSubclassProc, kListViewSubclassId, reinterpret_cast<DWORD_PTR>(this))) {
        LogLastError(L"SetWindowSubclass(FileListView)");
        return false;
    }

    const DWORD ex_style = LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_HEADERDRAGDROP;
    ListView_SetExtendedListViewStyle(hwnd_, ex_style);
    ListView_SetBkColor(hwnd_, kListBackgroundColor);
    ListView_SetTextBkColor(hwnd_, kListBackgroundColor);
    ListView_SetTextColor(hwnd_, kListTextColor);
    return true;
}

void FileListView::CreateColumns() {
    const struct ColumnSpec {
        const wchar_t* title;
        int base_width;
        int format;
    } columns[] = {
        {L"Name", 260, LVCFMT_LEFT},
        {L"Extension", 60, LVCFMT_LEFT},
        {L"Date Modified", 140, LVCFMT_LEFT},
        {L"Size", 90, LVCFMT_RIGHT},
        {L"Path", 0, LVCFMT_LEFT},
    };

    for (int i = 0; i < static_cast<int>(std::size(columns)); ++i) {
        LVCOLUMNW column = {};
        column.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
        column.fmt = columns[i].format;
        column.cx = MulDiv(columns[i].base_width, static_cast<int>(dpi_), 96);
        column.pszText = const_cast<LPWSTR>(columns[i].title);
        column.iSubItem = i;
        if (ListView_InsertColumn(hwnd_, i, &column) == -1) {
            LogLastError(L"ListView_InsertColumn");
        }
    }
}

void FileListView::ApplyScaledColumns() {
    if (hwnd_ == nullptr) {
        return;
    }

    const int widths[] = {
        MulDiv(260, static_cast<int>(dpi_), 96),
        MulDiv(60, static_cast<int>(dpi_), 96),
        MulDiv(140, static_cast<int>(dpi_), 96),
        MulDiv(90, static_cast<int>(dpi_), 96),
        search_mode_ ? MulDiv(320, static_cast<int>(dpi_), 96) : 0,
    };
    for (int i = 0; i < static_cast<int>(std::size(widths)); ++i) {
        ListView_SetColumnWidth(hwnd_, i, widths[i]);
    }
}

void FileListView::ApplyColumnMode() {
    if (hwnd_ == nullptr) {
        return;
    }

    ApplyScaledColumns();
    InvalidateRect(hwnd_, nullptr, TRUE);
}

void FileListView::ApplyHeaderVisualStyle() {
    if (hwnd_ == nullptr) {
        return;
    }

    HWND header = ListView_GetHeader(hwnd_);
    if (header == nullptr) {
        return;
    }

    const HRESULT theme_result = SetWindowTheme(header, L"", L"");
    if (FAILED(theme_result)) {
        LogHResult(L"SetWindowTheme(ListHeader)", theme_result);
    }

    InvalidateRect(header, nullptr, TRUE);
}

void FileListView::EnsureSystemImageList() {
    if (system_image_list_ != nullptr || hwnd_ == nullptr) {
        return;
    }

    SHFILEINFOW shell_info = {};
    system_image_list_ = reinterpret_cast<HIMAGELIST>(SHGetFileInfoW(
        L"C:\\",
        FILE_ATTRIBUTE_DIRECTORY,
        &shell_info,
        sizeof(shell_info),
        SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES));
    if (system_image_list_ == nullptr) {
        LogLastError(L"SHGetFileInfoW(SystemImageList)");
        return;
    }

    ListView_SetImageList(hwnd_, system_image_list_, LVSIL_SMALL);
}

void FileListView::EnsureDividerFont() {
    const int requested_height = -MulDiv(11, static_cast<int>(dpi_), 72);
    if (divider_font_ != nullptr && divider_font_height_ == requested_height) {
        return;
    }

    divider_font_.reset(CreateFontW(
        requested_height,
        0,
        0,
        0,
        FW_SEMIBOLD,
        FALSE,
        FALSE,
        FALSE,
        DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS,
        CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY,
        DEFAULT_PITCH | FF_DONTCARE,
        L"Segoe UI"));
    if (divider_font_ == nullptr) {
        LogLastError(L"CreateFontW(Divider)");
    }
    divider_font_height_ = requested_height;
}

bool FileListView::HandleGetDispInfo(NMLVDISPINFOW* info) const {
    if (info == nullptr || info->item.iItem < 0 || info->item.iItem >= static_cast<int>(display_entries_.size())) {
        return false;
    }

    const FileEntry& entry = display_entries_[info->item.iItem];
    if ((info->item.mask & LVIF_TEXT) != 0 && info->item.pszText != nullptr && info->item.cchTextMax > 0) {
        std::wstring text;
        if (entry.is_group_divider) {
            text = (info->item.iSubItem == 0) ? entry.group_label : L"";
        } else {
            switch (info->item.iSubItem) {
            case 0:
                text = entry.name;
                break;
            case 1:
                text = entry.is_folder ? L"" : entry.extension;
                break;
            case 2:
                text = FormatDateTime(entry.modified_time);
                break;
            case 3:
                text = entry.is_folder ? L"" : FormatFileSize(entry.size_bytes);
                break;
            case 4:
                text = entry.full_path;
                break;
            default:
                text.clear();
                break;
            }
        }
        StringCchCopyW(info->item.pszText, static_cast<size_t>(info->item.cchTextMax), text.c_str());
    }

    if ((info->item.mask & LVIF_IMAGE) != 0) {
        info->item.iImage = entry.is_group_divider ? -1 : entry.icon_index;
    }

    return true;
}

bool FileListView::HandleColumnClick(int column_index) {
    const int max_column_index = search_mode_ ? 4 : 3;
    if (column_index < 0 || column_index > max_column_index) {
        return false;
    }

    const SortColumn clicked_column = static_cast<SortColumn>(column_index);
    if (sort_column_ == clicked_column) {
        sort_direction_ = (sort_direction_ == SortDirection::Ascending)
            ? SortDirection::Descending
            : SortDirection::Ascending;
    } else {
        sort_column_ = clicked_column;
        sort_direction_ = SortDirection::Ascending;
    }

    SortBaseEntries();
    BuildDisplayEntriesWithGroups();
    UpdateSortIndicators();
    UpdateItemCountAndRefresh(true);
    PostStatusUpdate();
    return true;
}

bool FileListView::HandleCustomDraw(NMLVCUSTOMDRAW* custom_draw, LRESULT* result) {
    if (custom_draw == nullptr || result == nullptr) {
        return false;
    }

    switch (custom_draw->nmcd.dwDrawStage) {
    case CDDS_PREPAINT:
        *result = CDRF_NOTIFYITEMDRAW;
        return true;

    case CDDS_ITEMPREPAINT: {
        const size_t index = static_cast<size_t>(custom_draw->nmcd.dwItemSpec);
        if (index >= display_entries_.size()) {
            *result = CDRF_DODEFAULT;
            return true;
        }

        const FileEntry& entry = display_entries_[index];
        custom_draw->clrTextBk = kListBackgroundColor;

        if (entry.is_group_divider) {
            EnsureDividerFont();
            custom_draw->clrText = kDividerTextColor;
            custom_draw->clrTextBk = kDividerBackgroundColor;
            if (divider_font_ != nullptr) {
                SelectObject(custom_draw->nmcd.hdc, divider_font_.get());
                *result = CDRF_NEWFONT;
            } else {
                *result = CDRF_DODEFAULT;
            }
            return true;
        }

        const COLORREF extension_color = ColorForExtension(entry.extension, entry.is_folder);
        custom_draw->clrText = (extension_color == CLR_INVALID) ? kListTextColor : extension_color;
        *result = CDRF_DODEFAULT;
        return true;
    }

    default:
        break;
    }

    return false;
}

bool FileListView::HandleHeaderCustomDraw(NMCUSTOMDRAW* custom_draw, LRESULT* result) const {
    if (custom_draw == nullptr || result == nullptr) {
        return false;
    }

    HWND header = (hwnd_ != nullptr) ? ListView_GetHeader(hwnd_) : nullptr;
    if (header == nullptr) {
        return false;
    }

    switch (custom_draw->dwDrawStage) {
    case CDDS_PREPAINT:
        *result = CDRF_NOTIFYITEMDRAW;
        return true;

    case CDDS_ITEMPREPAINT: {
        const int column_index = static_cast<int>(custom_draw->dwItemSpec);
        RECT rect = custom_draw->rc;
        const bool hot = (custom_draw->uItemState & CDIS_HOT) != 0;

        HBRUSH header_brush = CreateSolidBrush(hot ? kHeaderHoverBackgroundColor : kHeaderBackgroundColor);
        if (header_brush != nullptr) {
            FillRect(custom_draw->hdc, &rect, header_brush);
            DeleteObject(header_brush);
        }

        wchar_t text_buffer[128] = {};
        HDITEMW item = {};
        item.mask = HDI_TEXT | HDI_FORMAT;
        item.pszText = text_buffer;
        item.cchTextMax = static_cast<int>(sizeof(text_buffer) / sizeof(text_buffer[0]));
        if (!Header_GetItem(header, column_index, &item)) {
            LogLastError(L"Header_GetItem(CustomDraw)");
            item.pszText = const_cast<LPWSTR>(L"");
            item.fmt = HDF_LEFT;
        }

        const int text_padding = MulDiv(8, static_cast<int>(dpi_), 96);
        const bool sorted_column = column_index == static_cast<int>(sort_column_);
        const int sort_reserved_width = sorted_column ? MulDiv(12, static_cast<int>(dpi_), 96) : 0;

        RECT text_rect = rect;
        text_rect.left += text_padding;
        text_rect.right -= text_padding + sort_reserved_width;

        SetBkMode(custom_draw->hdc, TRANSPARENT);
        SetTextColor(custom_draw->hdc, kHeaderTextColor);

        UINT text_flags = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS | DT_NOPREFIX;
        if ((item.fmt & HDF_RIGHT) != 0) {
            text_flags |= DT_RIGHT;
        } else {
            text_flags |= DT_LEFT;
        }
        DrawTextW(custom_draw->hdc, text_buffer, -1, &text_rect, text_flags);

        if (sorted_column) {
            const int arrow_width = MulDiv(8, static_cast<int>(dpi_), 96);
            const int arrow_height = MulDiv(5, static_cast<int>(dpi_), 96);
            const int center_x = rect.right - text_padding - (arrow_width / 2);
            const int center_y = rect.top + ((rect.bottom - rect.top) / 2);

            POINT points[3] = {};
            if (sort_direction_ == SortDirection::Ascending) {
                points[0] = {center_x - (arrow_width / 2), center_y + (arrow_height / 2)};
                points[1] = {center_x + (arrow_width / 2), center_y + (arrow_height / 2)};
                points[2] = {center_x, center_y - (arrow_height / 2)};
            } else {
                points[0] = {center_x - (arrow_width / 2), center_y - (arrow_height / 2)};
                points[1] = {center_x + (arrow_width / 2), center_y - (arrow_height / 2)};
                points[2] = {center_x, center_y + (arrow_height / 2)};
            }

            HBRUSH arrow_brush = CreateSolidBrush(kHeaderSortGlyphColor);
            if (arrow_brush != nullptr) {
                HGDIOBJ old_pen = SelectObject(custom_draw->hdc, GetStockObject(NULL_PEN));
                HGDIOBJ old_brush = SelectObject(custom_draw->hdc, arrow_brush);
                Polygon(custom_draw->hdc, points, 3);
                SelectObject(custom_draw->hdc, old_brush);
                SelectObject(custom_draw->hdc, old_pen);
                DeleteObject(arrow_brush);
            }
        }

        const int column_count = Header_GetItemCount(header);
        HPEN divider_pen = CreatePen(PS_SOLID, 1, kHeaderDividerColor);
        if (divider_pen != nullptr) {
            HGDIOBJ old_pen = SelectObject(custom_draw->hdc, divider_pen);
            if (column_index + 1 < column_count) {
                MoveToEx(custom_draw->hdc, rect.right - 1, rect.top + MulDiv(3, static_cast<int>(dpi_), 96), nullptr);
                LineTo(custom_draw->hdc, rect.right - 1, rect.bottom - 1);
            }
            MoveToEx(custom_draw->hdc, rect.left, rect.bottom - 1, nullptr);
            LineTo(custom_draw->hdc, rect.right, rect.bottom - 1);
            SelectObject(custom_draw->hdc, old_pen);
            DeleteObject(divider_pen);
        }

        *result = CDRF_SKIPDEFAULT;
        return true;
    }

    default:
        break;
    }

    return false;
}

bool FileListView::HandleDoubleClick(const NMITEMACTIVATE* item_activate) {
    if (item_activate == nullptr || item_activate->iItem < 0) {
        return false;
    }
    return OpenEntryAtIndex(item_activate->iItem, true);
}

bool FileListView::HandleItemChanging(const NMLISTVIEW* list_view, LRESULT* result) const {
    if (list_view == nullptr || result == nullptr) {
        return false;
    }

    if (list_view->iItem < 0 || list_view->iItem >= static_cast<int>(display_entries_.size())) {
        return false;
    }

    if (!IsDividerIndex(list_view->iItem)) {
        return false;
    }

    if ((list_view->uChanged & LVIF_STATE) == 0) {
        return false;
    }

    if ((list_view->uNewState & (LVIS_SELECTED | LVIS_FOCUSED)) != 0) {
        *result = TRUE;
        return true;
    }

    return false;
}

bool FileListView::HandleItemChanged(const NMLISTVIEW* list_view) {
    if (list_view == nullptr) {
        return false;
    }

    if ((list_view->uChanged & LVIF_STATE) == 0) {
        return false;
    }

    SanitizeSelectionAndFocus();
    PostStatusUpdate();
    return true;
}

bool FileListView::HandleClick(const NMITEMACTIVATE* item_activate) {
    if (item_activate == nullptr) {
        return false;
    }

    if (item_activate->iItem < 0) {
        ClearSelection();
        PostStatusUpdate();
        return true;
    }

    if (IsDividerIndex(item_activate->iItem)) {
        ClearSelection();
        PostStatusUpdate();
        return true;
    }

    return false;
}

LRESULT FileListView::HandleSubclassMessage(UINT message, WPARAM w_param, LPARAM l_param) {
    switch (message) {
    case WM_KEYDOWN: {
        const bool ctrl_down = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        const bool shift_down = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        const bool alt_down = (GetKeyState(VK_MENU) & 0x8000) != 0;
        if (w_param == VK_F5) {
            if (parent_hwnd_ != nullptr) {
                PostMessageW(parent_hwnd_, WM_FE_FILELIST_REFRESH, 0, 0);
            }
            return 0;
        }

        if (ctrl_down && (w_param == 'A' || w_param == 'a')) {
            ClearSelection();
            for (int i = 0; i < static_cast<int>(display_entries_.size()); ++i) {
                if (IsSelectableIndex(i)) {
                    ListView_SetItemState(hwnd_, i, LVIS_SELECTED, LVIS_SELECTED);
                }
            }
            const int first = FindNextSelectableIndex(-1, 1);
            if (first >= 0) {
                ListView_SetItemState(hwnd_, first, LVIS_FOCUSED, LVIS_FOCUSED);
                ListView_EnsureVisible(hwnd_, first, FALSE);
            }
            PostStatusUpdate();
            return 0;
        }

        if (w_param == VK_UP || w_param == VK_DOWN) {
            const int step = (w_param == VK_DOWN) ? 1 : -1;
            const int fallback_start =
                (step > 0) ? -1 : static_cast<int>(display_entries_.size());

            int focused = FocusedIndex();
            if (!IsSelectableIndex(focused)) {
                focused = FindNextSelectableIndex(fallback_start, step);
            }

            if (focused < 0) {
                return 0;
            }

            int target = FindNextSelectableIndex(focused, step);
            if (target < 0) {
                target = focused;
            }

            if (shift_down) {
                int anchor = ListView_GetSelectionMark(hwnd_);
                if (!IsSelectableIndex(anchor)) {
                    anchor = focused;
                }
                if (!IsSelectableIndex(anchor)) {
                    anchor = target;
                }

                ClearSelection();
                const int range_start = (std::min)(anchor, target);
                const int range_end = (std::max)(anchor, target);
                for (int i = range_start; i <= range_end; ++i) {
                    if (IsSelectableIndex(i)) {
                        ListView_SetItemState(hwnd_, i, LVIS_SELECTED, LVIS_SELECTED);
                    }
                }
                ListView_SetItemState(hwnd_, target, LVIS_FOCUSED, LVIS_FOCUSED);
                ListView_SetSelectionMark(hwnd_, anchor);
            } else if (ctrl_down) {
                ListView_SetItemState(hwnd_, -1, 0, LVIS_FOCUSED);
                ListView_SetItemState(hwnd_, target, LVIS_FOCUSED, LVIS_FOCUSED);
            } else {
                ClearSelection();
                ListView_SetItemState(
                    hwnd_,
                    target,
                    LVIS_SELECTED | LVIS_FOCUSED,
                    LVIS_SELECTED | LVIS_FOCUSED);
                ListView_SetSelectionMark(hwnd_, target);
            }

            ListView_EnsureVisible(hwnd_, target, FALSE);
            PostStatusUpdate();
            return 0;
        }

        if ((w_param == VK_LEFT || w_param == VK_BACK) && !alt_down) {
            NavigateToParentFolder();
            return 0;
        }

        if (w_param == VK_RIGHT) {
            int index = FocusedIndex();
            if (index < 0 || IsDividerIndex(index)) {
                index = FindNextSelectableIndex(index, 1);
            }
            OpenEntryAtIndex(index, false);
            return 0;
        }

        if (w_param == VK_RETURN) {
            int index = FocusedIndex();
            if (index < 0 || IsDividerIndex(index)) {
                index = FindNextSelectableIndex(index, 1);
            }
            OpenEntryAtIndex(index, true);
            return 0;
        }
        break;
    }

    case WM_CHAR: {
        const bool ctrl_down = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        const bool alt_down = (GetKeyState(VK_MENU) & 0x8000) != 0;
        const wchar_t ch = static_cast<wchar_t>(w_param);
        if (!ctrl_down && !alt_down && ch >= L' ' && HandleTypeAheadChar(ch)) {
            return 0;
        }
        break;
    }

    case WM_LBUTTONDOWN: {
        POINT point = {GET_X_LPARAM(l_param), GET_Y_LPARAM(l_param)};
        LVHITTESTINFO hit = {};
        hit.pt = point;
        const int item = ListView_SubItemHitTest(hwnd_, &hit);
        if (item < 0) {
            ClearSelection();
            PostStatusUpdate();
        }
        break;
    }

    case WM_CONTEXTMENU: {
        POINT screen_point = {GET_X_LPARAM(l_param), GET_Y_LPARAM(l_param)};
        int hit_index = -1;
        if (screen_point.x == -1 && screen_point.y == -1) {
            hit_index = FocusedIndex();
            if (hit_index >= 0) {
                RECT item_rect = {};
                item_rect.left = LVIR_BOUNDS;
                if (ListView_GetItemRect(hwnd_, hit_index, &item_rect, LVIR_BOUNDS)) {
                    screen_point.x = (item_rect.left + item_rect.right) / 2;
                    screen_point.y = (item_rect.top + item_rect.bottom) / 2;
                    ClientToScreen(hwnd_, &screen_point);
                }
            }
        } else {
            POINT client_point = screen_point;
            ScreenToClient(hwnd_, &client_point);
            LVHITTESTINFO hit = {};
            hit.pt = client_point;
            hit_index = ListView_SubItemHitTest(hwnd_, &hit);
        }
        ShowContextMenu(screen_point, hit_index);
        return 0;
    }

    case WM_PAINT: {
        const LRESULT paint_result = DefSubclassProc(hwnd_, message, w_param, l_param);
        if (!loading_ && base_entries_.empty()) {
            HDC hdc = GetDC(hwnd_);
            if (hdc != nullptr) {
                DrawEmptyState(hdc);
                ReleaseDC(hwnd_, hdc);
            }
        }
        if (loading_) {
            HDC hdc = GetDC(hwnd_);
            if (hdc != nullptr) {
                DrawLoadingOverlay(hdc);
                ReleaseDC(hwnd_, hdc);
            }
        }
        return paint_result;
    }

    case WM_TIMER:
        if (w_param == kLoadingTimerId && loading_) {
            loading_frame_ = (loading_frame_ + 1) % 12;
            InvalidateRect(hwnd_, nullptr, FALSE);
            return 0;
        }
        break;

    case WM_NCDESTROY:
        SetLoadingState(false);
        RemoveWindowSubclass(hwnd_, &FileListView::ListViewSubclassProc, kListViewSubclassId);
        break;

    default:
        break;
    }

    return DefSubclassProc(hwnd_, message, w_param, l_param);
}

void FileListView::SortBaseEntries() {
    auto compare = [this](const FileEntry& left, const FileEntry& right) {
        if (left.is_folder != right.is_folder) {
            return left.is_folder;
        }

        int cmp = 0;
        switch (sort_column_) {
        case SortColumn::Name:
            cmp = CompareTextInsensitive(left.name, right.name);
            break;

        case SortColumn::Extension:
            cmp = CompareTextInsensitive(left.extension, right.extension);
            if (cmp == 0) {
                cmp = CompareTextInsensitive(left.name, right.name);
            }
            break;

        case SortColumn::DateModified:
            cmp = static_cast<int>(CompareFileTime(&left.modified_time, &right.modified_time));
            if (cmp == 0) {
                cmp = CompareTextInsensitive(left.name, right.name);
            }
            break;

        case SortColumn::Size:
            if (left.size_bytes < right.size_bytes) {
                cmp = -1;
            } else if (left.size_bytes > right.size_bytes) {
                cmp = 1;
            } else {
                cmp = CompareTextInsensitive(left.name, right.name);
            }
            break;

        case SortColumn::Path:
            cmp = CompareTextInsensitive(left.full_path, right.full_path);
            if (cmp == 0) {
                cmp = CompareTextInsensitive(left.name, right.name);
            }
            break;
        }

        if (sort_direction_ == SortDirection::Ascending) {
            return cmp < 0;
        }
        return cmp > 0;
    };

    std::sort(base_entries_.begin(), base_entries_.end(), compare);
}

void FileListView::BuildDisplayEntriesWithGroups() {
    display_entries_.clear();
    if (base_entries_.empty()) {
        return;
    }

    if (search_mode_ || sort_column_ != SortColumn::DateModified) {
        display_entries_ = base_entries_;
        return;
    }

    bool have_bucket = false;
    DateBucket current_bucket = DateBucket::Older;
    for (const FileEntry& entry : base_entries_) {
        const DateBucket bucket = BucketForFileTime(entry.modified_time);
        if (!have_bucket || bucket != current_bucket) {
            FileEntry divider = {};
            divider.is_group_divider = true;
            divider.group_label = LabelForBucket(bucket);
            divider.icon_index = -1;
            display_entries_.push_back(std::move(divider));
            current_bucket = bucket;
            have_bucket = true;
        }
        display_entries_.push_back(entry);
    }
}

void FileListView::UpdateSortIndicators() {
    if (hwnd_ == nullptr) {
        return;
    }

    HWND header = ListView_GetHeader(hwnd_);
    if (header == nullptr) {
        return;
    }

    const int column_count = Header_GetItemCount(header);
    const int max_sort_index = search_mode_ ? 4 : 3;
    const int raw_active_index = static_cast<int>(sort_column_);
    const int active_index = (raw_active_index >= 0 && raw_active_index <= max_sort_index) ? raw_active_index : -1;

    for (int i = 0; i < column_count; ++i) {
        HDITEMW item = {};
        item.mask = HDI_FORMAT;
        if (!Header_GetItem(header, i, &item)) {
            continue;
        }

        item.fmt &= ~(HDF_SORTUP | HDF_SORTDOWN);
        if (i == active_index) {
            item.fmt |= (sort_direction_ == SortDirection::Ascending) ? HDF_SORTUP : HDF_SORTDOWN;
        }
        Header_SetItem(header, i, &item);
    }
}

void FileListView::UpdateItemCountAndRefresh(bool set_default_focus) {
    if (hwnd_ == nullptr) {
        return;
    }

    ListView_SetItemCountEx(
        hwnd_,
        static_cast<int>(display_entries_.size()),
        LVSICF_NOINVALIDATEALL | LVSICF_NOSCROLL);

    if (set_default_focus) {
        ListView_SetItemState(hwnd_, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);

        const int first = FindNextSelectableIndex(-1, 1);
        if (first >= 0) {
            ListView_SetItemState(
                hwnd_,
                first,
                LVIS_SELECTED | LVIS_FOCUSED,
                LVIS_SELECTED | LVIS_FOCUSED);
            ListView_SetSelectionMark(hwnd_, first);
            ListView_EnsureVisible(hwnd_, first, FALSE);
        }
    }

    InvalidateRect(hwnd_, nullptr, TRUE);
}

void FileListView::PostStatusUpdate() const {
    if (parent_hwnd_ == nullptr) {
        return;
    }

    auto text_payload = std::make_unique<std::wstring>(BuildStatusText());
    if (!PostMessageW(parent_hwnd_, WM_FE_FILELIST_STATUS_UPDATE, 0, reinterpret_cast<LPARAM>(text_payload.get()))) {
        LogLastError(L"PostMessageW(WM_FE_FILELIST_STATUS_UPDATE)");
        return;
    }
    text_payload.release();
}

std::wstring FileListView::BuildStatusText() const {
    const int total_items = static_cast<int>(base_entries_.size());
    const int selected_items = SelectedCountExcludingDividers();

    wchar_t buffer[384] = {};
    if (search_mode_) {
        if (loading_) {
            swprintf_s(
                buffer,
                L"Searching \"%s\" in %s...    Results: %d    Selected: %d",
                search_pattern_.c_str(),
                search_root_path_.c_str(),
                total_items,
                selected_items);
        } else {
            if (search_elapsed_ms_ > 0) {
                swprintf_s(
                    buffer,
                    L"Search results: %d    Selected: %d    Time: %llu ms",
                    total_items,
                    selected_items,
                    search_elapsed_ms_);
            } else {
                swprintf_s(
                    buffer,
                    L"Search results: %d    Selected: %d",
                    total_items,
                    selected_items);
            }
        }
    } else if (loading_) {
        swprintf_s(buffer, L"Loading...    Items: %d    Selected: %d", total_items, selected_items);
    } else {
        swprintf_s(buffer, L"Items: %d    Selected: %d", total_items, selected_items);
    }
    return buffer;
}

void FileListView::SetLoadingState(bool loading) {
    if (hwnd_ == nullptr) {
        loading_ = loading;
        if (!loading_) {
            loading_frame_ = 0;
        }
        return;
    }

    if (loading_ == loading) {
        return;
    }

    loading_ = loading;
    if (loading_) {
        loading_frame_ = 0;
        if (SetTimer(hwnd_, kLoadingTimerId, kLoadingTimerMs, nullptr) == 0) {
            LogLastError(L"SetTimer(FileListLoading)");
        }
    } else {
        if (!KillTimer(hwnd_, kLoadingTimerId)) {
            const DWORD last_error = GetLastError();
            if (last_error != ERROR_SUCCESS) {
                LogLastError(L"KillTimer(FileListLoading)");
            }
        }
        loading_frame_ = 0;
    }

    InvalidateRect(hwnd_, nullptr, FALSE);
}

void FileListView::DrawLoadingOverlay(HDC hdc) const {
    if (hdc == nullptr || hwnd_ == nullptr) {
        return;
    }

    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        return;
    }

    const int overlay_width = MulDiv(190, static_cast<int>(dpi_), 96);
    const int overlay_height = MulDiv(42, static_cast<int>(dpi_), 96);
    RECT overlay_rect = {};
    overlay_rect.left = ((client_rect.right - client_rect.left) - overlay_width) / 2;
    overlay_rect.top = ((client_rect.bottom - client_rect.top) - overlay_height) / 2;
    overlay_rect.right = overlay_rect.left + overlay_width;
    overlay_rect.bottom = overlay_rect.top + overlay_height;

    HBRUSH overlay_brush = CreateSolidBrush(kLoadingOverlayBackgroundColor);
    if (overlay_brush != nullptr) {
        FillRect(hdc, &overlay_rect, overlay_brush);
        DeleteObject(overlay_brush);
    }

    HPEN border_pen = CreatePen(PS_SOLID, 1, kLoadingOverlayBorderColor);
    if (border_pen != nullptr) {
        HGDIOBJ old_pen = SelectObject(hdc, border_pen);
        HGDIOBJ old_brush = SelectObject(hdc, GetStockObject(NULL_BRUSH));
        Rectangle(hdc, overlay_rect.left, overlay_rect.top, overlay_rect.right, overlay_rect.bottom);
        SelectObject(hdc, old_brush);
        SelectObject(hdc, old_pen);
        DeleteObject(border_pen);
    }

    static constexpr wchar_t kSpinnerFrames[] = L"|/-\\";
    const wchar_t spinner = kSpinnerFrames[loading_frame_ % 4];

    wchar_t text_buffer[64] = {};
    swprintf_s(text_buffer, L"Loading...  %c", spinner);

    RECT text_rect = overlay_rect;
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, kLoadingOverlayTextColor);
    DrawTextW(hdc, text_buffer, -1, &text_rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
}

void FileListView::CaptureViewState(
    std::vector<std::wstring>* selected_names,
    std::wstring* focused_name,
    int* top_index) const {
    if (selected_names == nullptr || focused_name == nullptr || top_index == nullptr || hwnd_ == nullptr) {
        return;
    }

    selected_names->clear();
    focused_name->clear();
    *top_index = ListView_GetTopIndex(hwnd_);

    int selected_index = -1;
    while (true) {
        selected_index = ListView_GetNextItem(hwnd_, selected_index, LVNI_SELECTED);
        if (selected_index < 0) {
            break;
        }
        if (!IsSelectableIndex(selected_index)) {
            continue;
        }
        selected_names->push_back(display_entries_[selected_index].name);
    }

    const int focused_index = FocusedIndex();
    if (IsSelectableIndex(focused_index)) {
        *focused_name = display_entries_[focused_index].name;
    }
}

void FileListView::RestoreViewState(
    const std::vector<std::wstring>& selected_names,
    const std::wstring& focused_name,
    int top_index) {
    if (hwnd_ == nullptr) {
        return;
    }

    ClearSelection();

    for (const std::wstring& name : selected_names) {
        const int index = FindSelectableIndexByName(name);
        if (index >= 0) {
            ListView_SetItemState(hwnd_, index, LVIS_SELECTED, LVIS_SELECTED);
        }
    }

    int focus_index = FindSelectableIndexByName(focused_name);
    if (focus_index < 0) {
        focus_index = ListView_GetNextItem(hwnd_, -1, LVNI_SELECTED);
    }
    if (focus_index < 0) {
        focus_index = FindNextSelectableIndex(-1, 1);
    }

    if (focus_index >= 0) {
        ListView_SetItemState(hwnd_, focus_index, LVIS_FOCUSED, LVIS_FOCUSED);
        ListView_SetSelectionMark(hwnd_, focus_index);
    }

    const int item_count = static_cast<int>(display_entries_.size());
    if (item_count > 0) {
        const int clamped_top = (std::max)(0, (std::min)(top_index, item_count - 1));
        ListView_EnsureVisible(hwnd_, clamped_top, FALSE);
    }

    if (focus_index >= 0) {
        ListView_EnsureVisible(hwnd_, focus_index, FALSE);
    }
}

void FileListView::ApplyRefreshDelta(std::vector<FileEntry> incoming_entries) {
    std::unordered_map<std::wstring, FileEntry> incoming_by_key;
    incoming_by_key.reserve(incoming_entries.size());
    for (FileEntry& entry : incoming_entries) {
        incoming_by_key.insert_or_assign(EntryKey(entry), std::move(entry));
    }

    std::vector<FileEntry> merged_entries;
    merged_entries.reserve(incoming_by_key.size());

    for (const FileEntry& current : base_entries_) {
        auto it = incoming_by_key.find(EntryKey(current));
        if (it == incoming_by_key.end()) {
            continue;
        }

        if (EntriesEquivalent(current, it->second)) {
            merged_entries.push_back(current);
        } else {
            merged_entries.push_back(std::move(it->second));
        }
        incoming_by_key.erase(it);
    }

    for (auto& [key, entry] : incoming_by_key) {
        (void)key;
        merged_entries.push_back(std::move(entry));
    }

    base_entries_ = std::move(merged_entries);
}

void FileListView::ClearSelection() {
    if (hwnd_ == nullptr) {
        return;
    }
    ListView_SetItemState(hwnd_, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
}

void FileListView::SanitizeSelectionAndFocus() {
    if (hwnd_ == nullptr || sanitizing_selection_) {
        return;
    }

    sanitizing_selection_ = true;

    int selected_index = -1;
    while (true) {
        selected_index = ListView_GetNextItem(hwnd_, selected_index, LVNI_SELECTED);
        if (selected_index < 0) {
            break;
        }
        if (IsDividerIndex(selected_index)) {
            ListView_SetItemState(hwnd_, selected_index, 0, LVIS_SELECTED);
        }
    }

    const int focused_index = FocusedIndex();
    if (IsDividerIndex(focused_index)) {
        const int candidate_down = FindNextSelectableIndex(focused_index, 1);
        const int candidate_up = FindNextSelectableIndex(focused_index, -1);
        const int replacement = (candidate_down >= 0) ? candidate_down : candidate_up;

        ListView_SetItemState(hwnd_, focused_index, 0, LVIS_FOCUSED);
        if (replacement >= 0) {
            ListView_SetItemState(hwnd_, replacement, LVIS_FOCUSED, LVIS_FOCUSED);
            ListView_EnsureVisible(hwnd_, replacement, FALSE);
        }
    }

    sanitizing_selection_ = false;
}

bool FileListView::NavigateToParentFolder() {
    const std::wstring parent = ParentPath(current_path_);
    if (parent.empty()) {
        return false;
    }

    auto payload = std::make_unique<std::wstring>(parent);
    if (!PostMessageW(parent_hwnd_, WM_FE_FILELIST_NAVIGATE, 0, reinterpret_cast<LPARAM>(payload.get()))) {
        LogLastError(L"PostMessageW(WM_FE_FILELIST_NAVIGATE Parent)");
        return false;
    }
    payload.release();
    return true;
}

bool FileListView::HandleTypeAheadChar(wchar_t character) {
    if (display_entries_.empty()) {
        return false;
    }

    const wchar_t normalized = static_cast<wchar_t>(towlower(character));
    const DWORD now = GetTickCount();
    if (type_ahead_buffer_.empty() || (now - type_ahead_last_tick_) > kTypeAheadResetMs) {
        type_ahead_buffer_.clear();
    }
    type_ahead_last_tick_ = now;

    type_ahead_buffer_.push_back(normalized);
    if (SelectByPrefix(type_ahead_buffer_)) {
        return true;
    }

    if (type_ahead_buffer_.size() > 1) {
        type_ahead_buffer_.assign(1, normalized);
        return SelectByPrefix(type_ahead_buffer_);
    }

    return false;
}

bool FileListView::SelectByPrefix(const std::wstring& prefix) {
    if (prefix.empty()) {
        return false;
    }

    const int total = static_cast<int>(display_entries_.size());
    if (total <= 0) {
        return false;
    }

    int start = FocusedIndex();
    if (start < -1 || start >= total) {
        start = -1;
    }

    for (int offset = 1; offset <= total; ++offset) {
        const int index = (start + offset) % total;
        if (!IsSelectableIndex(index)) {
            continue;
        }
        if (!StartsWithInsensitive(display_entries_[index].name, prefix)) {
            continue;
        }

        ClearSelection();
        ListView_SetItemState(hwnd_, index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
        ListView_EnsureVisible(hwnd_, index, FALSE);
        PostStatusUpdate();
        return true;
    }

    return false;
}

int FileListView::FindNextSelectableIndex(int start_index, int step) const {
    if (display_entries_.empty() || step == 0) {
        return -1;
    }

    int index = start_index;
    while (true) {
        index += step;
        if (index < 0 || index >= static_cast<int>(display_entries_.size())) {
            return -1;
        }
        if (IsSelectableIndex(index)) {
            return index;
        }
    }
}

bool FileListView::IsSelectableIndex(int index) const {
    return index >= 0 &&
        index < static_cast<int>(display_entries_.size()) &&
        !display_entries_[index].is_group_divider;
}

bool FileListView::IsDividerIndex(int index) const {
    return index >= 0 &&
        index < static_cast<int>(display_entries_.size()) &&
        display_entries_[index].is_group_divider;
}

int FileListView::FindSelectableIndexByName(const std::wstring& name) const {
    if (name.empty()) {
        return -1;
    }

    for (int i = 0; i < static_cast<int>(display_entries_.size()); ++i) {
        if (!IsSelectableIndex(i)) {
            continue;
        }
        if (CompareTextInsensitive(display_entries_[i].name, name) == 0) {
            return i;
        }
    }

    return -1;
}

bool FileListView::SelectSingleEntryByName(const std::wstring& name) {
    if (name.empty() || hwnd_ == nullptr) {
        return false;
    }

    const int index = FindSelectableIndexByName(name);
    if (!IsSelectableIndex(index)) {
        return false;
    }

    ClearSelection();
    ListView_SetItemState(
        hwnd_,
        index,
        LVIS_SELECTED | LVIS_FOCUSED,
        LVIS_SELECTED | LVIS_FOCUSED);
    ListView_SetSelectionMark(hwnd_, index);
    ListView_EnsureVisible(hwnd_, index, FALSE);
    return true;
}

int FileListView::SelectedCountExcludingDividers() const {
    if (hwnd_ == nullptr) {
        return 0;
    }

    int count = 0;
    int index = -1;
    while (true) {
        index = ListView_GetNextItem(hwnd_, index, LVNI_SELECTED);
        if (index < 0) {
            break;
        }
        if (!IsDividerIndex(index)) {
            ++count;
        }
    }
    return count;
}

int FileListView::FocusedIndex() const {
    if (hwnd_ == nullptr) {
        return -1;
    }
    return ListView_GetNextItem(hwnd_, -1, LVNI_FOCUSED);
}

bool FileListView::OpenEntryAtIndex(int index, bool open_files_too) {
    if (!IsSelectableIndex(index)) {
        return false;
    }

    const FileEntry& entry = display_entries_[index];
    const std::wstring full_path = BuildFullPath(entry);

    if (entry.is_folder) {
        auto payload = std::make_unique<std::wstring>(full_path);
        if (!PostMessageW(parent_hwnd_, WM_FE_FILELIST_NAVIGATE, 0, reinterpret_cast<LPARAM>(payload.get()))) {
            LogLastError(L"PostMessageW(WM_FE_FILELIST_NAVIGATE)");
            return false;
        }
        payload.release();
        return true;
    }

    if (!open_files_too) {
        return false;
    }

    const HINSTANCE open_result = ShellExecuteW(
        nullptr,
        L"open",
        full_path.c_str(),
        nullptr,
        nullptr,
        SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(open_result) <= 32) {
        LogLastError(L"ShellExecuteW(FileOpen)");
        return false;
    }
    return true;
}

void FileListView::ShowContextMenu(POINT screen_point, int hit_index) {
    HMENU menu = CreatePopupMenu();
    if (menu == nullptr) {
        LogLastError(L"CreatePopupMenu(FileList)");
        return;
    }

    AppendMenuW(menu, MF_STRING, kMenuCopy, L"Copy");
    AppendMenuW(menu, MF_STRING, kMenuCut, L"Cut");
    AppendMenuW(menu, MF_STRING, kMenuPaste, L"Paste");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING, kMenuRename, L"Rename");
    AppendMenuW(menu, MF_STRING, kMenuDelete, L"Delete");
    AppendMenuW(menu, MF_STRING, kMenuProperties, L"Properties");
    AppendMenuW(menu, MF_STRING, kMenuCopyPath, L"Copy Path");
    if (search_mode_) {
        AppendMenuW(menu, MF_STRING, kMenuOpenFileLocation, L"Open file location");
    }
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING, kMenuMakeFavourite, L"Make Favourite");
    AppendMenuW(menu, MF_STRING, kMenuMakeFlyoutFavourite, L"Make Flyout Favourite");

    int target_index = hit_index;
    if (!IsSelectableIndex(target_index)) {
        target_index = FocusedIndex();
    }
    if (!IsSelectableIndex(target_index)) {
        target_index = FindNextSelectableIndex(-1, 1);
    }

    const bool has_target = IsSelectableIndex(target_index);
    if (!has_target) {
        const UINT disable_flags = MF_BYCOMMAND | MF_GRAYED;
        EnableMenuItem(menu, kMenuCopy, disable_flags);
        EnableMenuItem(menu, kMenuCut, disable_flags);
        EnableMenuItem(menu, kMenuRename, disable_flags);
        EnableMenuItem(menu, kMenuDelete, disable_flags);
        EnableMenuItem(menu, kMenuProperties, disable_flags);
        EnableMenuItem(menu, kMenuCopyPath, disable_flags);
        if (search_mode_) {
            EnableMenuItem(menu, kMenuOpenFileLocation, disable_flags);
        }
        EnableMenuItem(menu, kMenuMakeFavourite, disable_flags);
        EnableMenuItem(menu, kMenuMakeFlyoutFavourite, disable_flags);
    } else if (!search_mode_ || display_entries_[target_index].is_folder) {
        if (search_mode_) {
            EnableMenuItem(menu, kMenuOpenFileLocation, MF_BYCOMMAND | MF_GRAYED);
        }
    }

    if (has_target && !display_entries_[target_index].is_folder) {
        const UINT disable_flags = MF_BYCOMMAND | MF_GRAYED;
        EnableMenuItem(menu, kMenuMakeFavourite, disable_flags);
        EnableMenuItem(menu, kMenuMakeFlyoutFavourite, disable_flags);
    }

    const UINT command = static_cast<UINT>(TrackPopupMenuEx(
        menu,
        TPM_RIGHTBUTTON | TPM_RETURNCMD,
        screen_point.x,
        screen_point.y,
        hwnd_,
        nullptr));

    if (command != 0 && has_target) {
        const FileEntry& entry = display_entries_[target_index];
        const std::wstring full_path = BuildFullPath(entry);

        switch (command) {
        case kMenuCopyPath:
            CopyTextToClipboard(parent_hwnd_, full_path);
            break;

        case kMenuOpenFileLocation: {
            const std::wstring containing_folder = ParentPath(full_path);
            if (!containing_folder.empty()) {
                auto payload = std::make_unique<std::wstring>(containing_folder);
                if (!PostMessageW(
                        parent_hwnd_,
                        WM_FE_FILELIST_OPEN_LOCATION_NEW_TAB,
                        0,
                        reinterpret_cast<LPARAM>(payload.get()))) {
                    LogLastError(L"PostMessageW(WM_FE_FILELIST_OPEN_LOCATION_NEW_TAB)");
                } else {
                    payload.release();
                }
            }
            break;
        }

        case kMenuMakeFavourite:
            if (entry.is_folder) {
                auto payload = std::make_unique<std::wstring>(full_path);
                if (!PostMessageW(
                        parent_hwnd_,
                        WM_FE_FILELIST_ADD_REGULAR_FAVOURITE,
                        0,
                        reinterpret_cast<LPARAM>(payload.get()))) {
                    LogLastError(L"PostMessageW(WM_FE_FILELIST_ADD_REGULAR_FAVOURITE)");
                } else {
                    payload.release();
                }
            }
            break;

        case kMenuMakeFlyoutFavourite:
            if (entry.is_folder) {
                auto payload = std::make_unique<std::wstring>(full_path);
                if (!PostMessageW(
                        parent_hwnd_,
                        WM_FE_FILELIST_ADD_FLYOUT_FAVOURITE,
                        0,
                        reinterpret_cast<LPARAM>(payload.get()))) {
                    LogLastError(L"PostMessageW(WM_FE_FILELIST_ADD_FLYOUT_FAVOURITE)");
                } else {
                    payload.release();
                }
            }
            break;

        case kMenuProperties:
            ShellExecuteW(nullptr, L"properties", full_path.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
            break;

        case kMenuRename:
            PostMessageW(hwnd_, WM_KEYDOWN, VK_F2, 0);
            break;

        default:
            break;
        }
    }

    DestroyMenu(menu);
}

void FileListView::DrawEmptyState(HDC hdc) const {
    if (hdc == nullptr || hwnd_ == nullptr || !base_entries_.empty()) {
        return;
    }

    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        return;
    }

    std::wstring text = L"This folder is empty";
    if (search_mode_) {
        text = L"No files found matching \"";
        text.append(search_pattern_);
        text.push_back(L'"');
    }
    RECT text_rect = client_rect;
    text_rect.top += MulDiv(36, static_cast<int>(dpi_), 96);

    SHFILEINFOW shell_info = {};
    if (SHGetFileInfoW(
            L"C:\\",
            FILE_ATTRIBUTE_DIRECTORY,
            &shell_info,
            sizeof(shell_info),
            SHGFI_ICON | SHGFI_LARGEICON | SHGFI_USEFILEATTRIBUTES) != 0 &&
        shell_info.hIcon != nullptr) {
        const int icon_x = ((client_rect.right - client_rect.left) - 32) / 2;
        const int icon_y = ((client_rect.bottom - client_rect.top) / 2) - 28;
        DrawIconEx(hdc, icon_x, icon_y, shell_info.hIcon, 32, 32, 0, nullptr, DI_NORMAL);
        DestroyIcon(shell_info.hIcon);
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, kEmptyStateTextColor);
    DrawTextW(hdc, text.c_str(), -1, &text_rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
}

std::wstring FileListView::BuildFullPath(const FileEntry& entry) const {
    if (!entry.full_path.empty()) {
        return entry.full_path;
    }

    std::wstring result = current_path_;
    if (result.empty()) {
        return entry.name;
    }
    if (result.back() != L'\\') {
        result.push_back(L'\\');
    }
    result.append(entry.name);
    return result;
}

std::wstring FileListView::NormalizePath(std::wstring path) {
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

std::wstring FileListView::ParentPath(const std::wstring& path) {
    std::wstring normalized = NormalizePath(path);
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

std::wstring FileListView::EntryKey(const FileEntry& entry) {
    std::wstring key = ToLowerCopy(entry.name);
    key.push_back(L'|');
    key.push_back(entry.is_folder ? L'D' : L'F');
    return key;
}

bool FileListView::EntriesEquivalent(const FileEntry& left, const FileEntry& right) {
    return CompareTextInsensitive(left.name, right.name) == 0 &&
        left.is_folder == right.is_folder &&
        left.attributes == right.attributes &&
        left.size_bytes == right.size_bytes &&
        CompareFileTime(&left.modified_time, &right.modified_time) == 0 &&
        left.icon_index == right.icon_index &&
        CompareTextInsensitive(left.extension, right.extension) == 0;
}

std::wstring FileListView::GetExtensionFromName(const std::wstring& name) {
    const std::wstring::size_type dot = name.find_last_of(L'.');
    if (dot == std::wstring::npos || dot == 0 || dot + 1 >= name.size()) {
        return L"";
    }
    return ToLowerCopy(name.substr(dot));
}

int FileListView::CompareTextInsensitive(const std::wstring& left, const std::wstring& right) {
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

std::wstring FileListView::FormatDateTime(const FILETIME& file_time) {
    FILETIME local_time = {};
    SYSTEMTIME system_time = {};
    if (!FileTimeToLocalFileTime(&file_time, &local_time) || !FileTimeToSystemTime(&local_time, &system_time)) {
        return L"";
    }

    wchar_t buffer[64] = {};
    swprintf_s(
        buffer,
        L"%04u-%02u-%02u %02u:%02u",
        static_cast<unsigned>(system_time.wYear),
        static_cast<unsigned>(system_time.wMonth),
        static_cast<unsigned>(system_time.wDay),
        static_cast<unsigned>(system_time.wHour),
        static_cast<unsigned>(system_time.wMinute));
    return buffer;
}

std::wstring FileListView::FormatFileSize(ULONGLONG size_bytes) {
    wchar_t buffer[64] = {};
    if (size_bytes < 1024ULL) {
        swprintf_s(buffer, L"%llu B", size_bytes);
        return buffer;
    }

    const double kb = static_cast<double>(size_bytes) / 1024.0;
    if (kb < 1024.0) {
        swprintf_s(buffer, L"%.1f KB", kb);
        return buffer;
    }

    const double mb = kb / 1024.0;
    if (mb < 1024.0) {
        swprintf_s(buffer, L"%.1f MB", mb);
        return buffer;
    }

    const double gb = mb / 1024.0;
    swprintf_s(buffer, L"%.1f GB", gb);
    return buffer;
}

int FileListView::ResolveIconIndex(const std::wstring& full_path, bool is_folder) {
    SHFILEINFOW shell_info = {};
    const DWORD attributes = is_folder ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL;
    if (SHGetFileInfoW(
            full_path.c_str(),
            attributes,
            &shell_info,
            sizeof(shell_info),
            SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES) == 0) {
        return 0;
    }
    return shell_info.iIcon;
}

COLORREF FileListView::ColorForExtension(const std::wstring& extension, bool is_folder) {
    if (is_folder || extension.empty()) {
        return CLR_INVALID;
    }

    const std::wstring lower = ToLowerCopy(extension);
    if (lower == L".pdf") {
        return RGB(0xE0, 0x6C, 0x75);
    }
    if (lower == L".xls" || lower == L".xlsx" || lower == L".xlsm") {
        return RGB(0x98, 0xC3, 0x79);
    }
    if (lower == L".doc" || lower == L".docx") {
        return RGB(0x61, 0xAF, 0xEF);
    }
    if (lower == L".zip" || lower == L".7z" || lower == L".rar" || lower == L".tar" || lower == L".gz") {
        return RGB(0xE5, 0xC0, 0x7B);
    }
    if (lower == L".txt" || lower == L".md" || lower == L".log") {
        return RGB(0x56, 0xB6, 0xC2);
    }
    return CLR_INVALID;
}

FileListView::DateBucket FileListView::BucketForFileTime(const FILETIME& file_time) {
    FILETIME file_local = {};
    if (!FileTimeToLocalFileTime(&file_time, &file_local)) {
        return DateBucket::Older;
    }

    FILETIME now_utc = {};
    GetSystemTimeAsFileTime(&now_utc);

    FILETIME now_local = {};
    if (!FileTimeToLocalFileTime(&now_utc, &now_local)) {
        return DateBucket::Older;
    }

    SYSTEMTIME file_st = {};
    SYSTEMTIME now_st = {};
    if (!FileTimeToSystemTime(&file_local, &file_st) || !FileTimeToSystemTime(&now_local, &now_st)) {
        return DateBucket::Older;
    }

    if (SameDate(file_st, now_st)) {
        return DateBucket::Today;
    }

    constexpr ULONGLONG kTicksPerDay = 24ULL * 60ULL * 60ULL * 10000000ULL;
    const ULONGLONG now_ticks = FileTimeToUint64(now_local);
    const ULONGLONG file_ticks = FileTimeToUint64(file_local);

    if (now_ticks >= kTicksPerDay) {
        const FILETIME yesterday_ft = Uint64ToFileTime(now_ticks - kTicksPerDay);
        SYSTEMTIME yesterday_st = {};
        if (FileTimeToSystemTime(&yesterday_ft, &yesterday_st) && SameDate(file_st, yesterday_st)) {
            return DateBucket::Yesterday;
        }
    }

    if (now_ticks >= file_ticks && (now_ticks - file_ticks) <= (7ULL * kTicksPerDay)) {
        return DateBucket::ThisWeek;
    }

    if (file_st.wYear == now_st.wYear && file_st.wMonth == now_st.wMonth) {
        return DateBucket::ThisMonth;
    }

    return DateBucket::Older;
}

const wchar_t* FileListView::LabelForBucket(DateBucket bucket) {
    switch (bucket) {
    case DateBucket::Today:
        return L"Today";
    case DateBucket::Yesterday:
        return L"Yesterday";
    case DateBucket::ThisWeek:
        return L"This week";
    case DateBucket::ThisMonth:
        return L"This month";
    case DateBucket::Older:
    default:
        return L"Older";
    }
}

}  // namespace fileexplorer
