#include "FileListView.h"
#include "FileOps.h"
#include "NavBar.h"
#include "Colors.h"

#include <Objbase.h>
#include <Objidl.h>
#include <Ole2.h>
#include <OleIdl.h>
#include <Shellapi.h>
#include <ShlObj.h>
#include <Shobjidl.h>
#include <Shlwapi.h>
#include <Windowsx.h>
#include <strsafe.h>
#include <uxtheme.h>

#include <algorithm>
#include <atomic>
#include <cstdlib>
#include <cwchar>
#include <cwctype>
#include <memory>
#include <new>
#include <unordered_map>
#include <unordered_set>
#include <utility>

namespace {

constexpr UINT_PTR kListViewSubclassId = 1;
constexpr UINT_PTR kLoadingTimerId = 9001;

constexpr COLORREF kListBackgroundColor = fileexplorer::colors::kBgFileList;
constexpr COLORREF kListTextColor = RGB(0xE2, 0xE8, 0xEF);
constexpr COLORREF kListCutPendingTextColor = RGB(0x93, 0x9C, 0xAA);
constexpr COLORREF kHeaderBackgroundColor = kListBackgroundColor;
constexpr COLORREF kHeaderHoverBackgroundColor = RGB(0x1C, 0x22, 0x2A);
constexpr COLORREF kHeaderTextColor = RGB(0xD2, 0xDA, 0xE4);
constexpr COLORREF kHeaderDividerColor = fileexplorer::colors::kSeparatorSubtle;
constexpr COLORREF kHeaderSortGlyphColor = RGB(0xA9, 0xB4, 0xC2);
constexpr COLORREF kDividerBackgroundColor = RGB(0x1E, 0x24, 0x2B);
constexpr COLORREF kDividerTextColor = RGB(0xAC, 0xB6, 0xC2);
constexpr COLORREF kEmptyStateTitleColor = RGB(0xD9, 0xE1, 0xEB);
constexpr COLORREF kEmptyStateTextColor = RGB(0x9F, 0xA8, 0xB5);
constexpr COLORREF kLoadingOverlayBackgroundColor = RGB(0x24, 0x2A, 0x33);
constexpr COLORREF kLoadingOverlayBorderColor = RGB(0x3A, 0x43, 0x4F);
constexpr COLORREF kLoadingOverlayTextColor = RGB(0xDF, 0xE5, 0xEE);
constexpr COLORREF kLoadingSpinnerBaseColor = RGB(0x5B, 0x67, 0x76);
constexpr COLORREF kLoadingSpinnerActiveColor = RGB(0xE2, 0xEC, 0xF8);

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
constexpr UINT kMenuNewFolder = 7011;
constexpr UINT kMenuToggleShowHidden = 7012;
constexpr UINT kMenuToggleShowExtensions = 7013;
constexpr UINT kMenuIncreaseDetailScale = 7014;
constexpr UINT kMenuResetDetailScale = 7015;
constexpr UINT_PTR kRenameEditSubclassId = 2;
constexpr UINT_PTR kHeaderEraseSubclassId = 3;
constexpr UINT kShellMenuCommandFirst = 7400;
constexpr UINT kShellMenuCommandLast = 7599;
constexpr DWORD kTypeAheadResetMs = 1200;
constexpr UINT kLoadingTimerMs = 80;
constexpr int kMinColumnWidthLogical = 40;
constexpr int kMaxColumnWidthLogical = 1800;
constexpr int kDefaultDetailScalePercent = 100;
constexpr int kDetailScaleMinPercent = 100;
constexpr int kDetailScaleMaxPercent = 300;
constexpr int kListBodyPointSize = 9;
constexpr int kListHeaderPointSize = 9;
constexpr int kListRowPaddingLogical = 4;

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

bool TryGetMenuItemTextByPosition(HMENU menu, UINT position, std::wstring* text) {
    if (menu == nullptr || text == nullptr) {
        return false;
    }

    text->clear();
    const int required_chars = GetMenuStringW(menu, position, nullptr, 0, MF_BYPOSITION);
    if (required_chars <= 0) {
        return false;
    }

    std::vector<wchar_t> buffer(static_cast<size_t>(required_chars) + 1U, L'\0');
    const int copied_chars = GetMenuStringW(
        menu,
        position,
        buffer.data(),
        static_cast<int>(buffer.size()),
        MF_BYPOSITION);
    if (copied_chars <= 0) {
        return false;
    }

    text->assign(buffer.data(), static_cast<size_t>(copied_chars));
    return true;
}

std::wstring NormalizeMenuLabel(const std::wstring& text) {
    std::wstring normalized;
    normalized.reserve(text.size());
    for (wchar_t ch : text) {
        if (ch == L'&') {
            continue;
        }
        if (ch == L'\t') {
            break;
        }
        normalized.push_back(static_cast<wchar_t>(std::towlower(ch)));
    }

    const size_t first_non_space = normalized.find_first_not_of(L" ");
    if (first_non_space == std::wstring::npos) {
        return std::wstring();
    }
    const size_t last_non_space = normalized.find_last_not_of(L" ");
    return normalized.substr(first_non_space, (last_non_space - first_non_space) + 1U);
}

bool IsSevenZipMenuLabel(const std::wstring& text) {
    const std::wstring normalized = NormalizeMenuLabel(text);
    if (normalized.empty()) {
        return false;
    }
    return normalized.find(L"7-zip") != std::wstring::npos ||
           normalized.find(L"7zip") != std::wstring::npos;
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

std::wstring ToLowerCopy(std::wstring value) {
    std::transform(
        value.begin(),
        value.end(),
        value.begin(),
        [](wchar_t ch) { return static_cast<wchar_t>(towlower(ch)); });
    return value;
}

std::wstring TrimExtensionForDisplay(const std::wstring& file_name) {
    const std::wstring::size_type dot = file_name.find_last_of(L'.');
    if (dot == std::wstring::npos || dot == 0) {
        return file_name;
    }
    return file_name.substr(0, dot);
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

BYTE LerpChannel(BYTE from, BYTE to, int numerator, int denominator) {
    if (denominator <= 0) {
        return from;
    }
    const int value = static_cast<int>(from) +
        ((static_cast<int>(to) - static_cast<int>(from)) * numerator) / denominator;
    return static_cast<BYTE>((std::max)(0, (std::min)(255, value)));
}

COLORREF LerpColor(COLORREF from, COLORREF to, int numerator, int denominator) {
    return RGB(
        LerpChannel(GetRValue(from), GetRValue(to), numerator, denominator),
        LerpChannel(GetGValue(from), GetGValue(to), numerator, denominator),
        LerpChannel(GetBValue(from), GetBValue(to), numerator, denominator));
}

HBITMAP CreateTransparentDibBitmap(int width, int height) {
    if (width <= 0 || height <= 0) {
        return nullptr;
    }

    BITMAPINFO bitmap_info = {};
    bitmap_info.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bitmap_info.bmiHeader.biWidth = width;
    bitmap_info.bmiHeader.biHeight = -height;  // top-down
    bitmap_info.bmiHeader.biPlanes = 1;
    bitmap_info.bmiHeader.biBitCount = 32;
    bitmap_info.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    HBITMAP bitmap = CreateDIBSection(nullptr, &bitmap_info, DIB_RGB_COLORS, &bits, nullptr, 0);
    if (bitmap == nullptr) {
        return nullptr;
    }

    if (bits != nullptr) {
        const size_t bytes = static_cast<size_t>(width) * static_cast<size_t>(height) * sizeof(DWORD);
        ZeroMemory(bits, bytes);
    }

    return bitmap;
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

HFONT CreateEmptyStateFont(int point_size, UINT dpi, LONG weight) {
    const int pixel_height = -MulDiv(point_size, static_cast<int>(dpi), 72);
    HFONT font = CreateFontW(
        pixel_height,
        0,
        0,
        0,
        weight,
        FALSE,
        FALSE,
        FALSE,
        DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS,
        CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY,
        DEFAULT_PITCH | FF_DONTCARE,
        L"Segoe UI Variable");
    if (font != nullptr) {
        return font;
    }

    return CreateFontW(
        pixel_height,
        0,
        0,
        0,
        weight,
        FALSE,
        FALSE,
        FALSE,
        DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS,
        CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY,
        DEFAULT_PITCH | FF_DONTCARE,
        L"Segoe UI");
}

HFONT CreateUiFontFromPixelHeight(int pixel_height, LONG weight) {
    HFONT font = CreateFontW(
        pixel_height,
        0,
        0,
        0,
        weight,
        FALSE,
        FALSE,
        FALSE,
        DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS,
        CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY,
        DEFAULT_PITCH | FF_DONTCARE,
        L"Segoe UI Variable Text");
    if (font != nullptr) {
        return font;
    }

    return CreateFontW(
        pixel_height,
        0,
        0,
        0,
        weight,
        FALSE,
        FALSE,
        FALSE,
        DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS,
        CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY,
        DEFAULT_PITCH | FF_DONTCARE,
        L"Segoe UI");
}

HFONT CreateMdl2Font(int point_size, UINT dpi, LONG weight) {
    const int pixel_height = -MulDiv(point_size, static_cast<int>(dpi), 72);
    return CreateFontW(
        pixel_height,
        0,
        0,
        0,
        weight,
        FALSE,
        FALSE,
        FALSE,
        DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS,
        CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY,
        DEFAULT_PITCH | FF_DONTCARE,
        L"Segoe MDL2 Assets");
}

HGLOBAL DuplicateHGlobal(HGLOBAL source) {
    if (source == nullptr) {
        return nullptr;
    }

    const SIZE_T bytes = GlobalSize(source);
    if (bytes == 0U) {
        return nullptr;
    }

    HGLOBAL copy = GlobalAlloc(GMEM_MOVEABLE, bytes);
    if (copy == nullptr) {
        LogLastError(L"GlobalAlloc(DuplicateHGlobal)");
        return nullptr;
    }

    void* source_block = GlobalLock(source);
    if (source_block == nullptr) {
        LogLastError(L"GlobalLock(DuplicateHGlobal Source)");
        GlobalFree(copy);
        return nullptr;
    }

    void* copy_block = GlobalLock(copy);
    if (copy_block == nullptr) {
        LogLastError(L"GlobalLock(DuplicateHGlobal Copy)");
        GlobalUnlock(source);
        GlobalFree(copy);
        return nullptr;
    }

    CopyMemory(copy_block, source_block, bytes);
    GlobalUnlock(copy);
    GlobalUnlock(source);
    return copy;
}

HGLOBAL CreateDropFilesHandle(const std::vector<std::wstring>& paths) {
    if (paths.empty()) {
        return nullptr;
    }

    size_t string_chars = 1;
    for (const std::wstring& path : paths) {
        if (!path.empty()) {
            string_chars += path.size() + 1;
        }
    }

    const SIZE_T bytes = sizeof(DROPFILES) + (string_chars * sizeof(wchar_t));
    HGLOBAL drop_handle = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, bytes);
    if (drop_handle == nullptr) {
        LogLastError(L"GlobalAlloc(DragDrop CF_HDROP)");
        return nullptr;
    }

    void* block = GlobalLock(drop_handle);
    if (block == nullptr) {
        LogLastError(L"GlobalLock(DragDrop CF_HDROP)");
        GlobalFree(drop_handle);
        return nullptr;
    }

    DROPFILES* drop_files = reinterpret_cast<DROPFILES*>(block);
    drop_files->pFiles = sizeof(DROPFILES);
    drop_files->fWide = TRUE;

    auto* cursor = reinterpret_cast<wchar_t*>(reinterpret_cast<BYTE*>(block) + sizeof(DROPFILES));
    for (const std::wstring& path : paths) {
        if (path.empty()) {
            continue;
        }
        const size_t byte_count = path.size() * sizeof(wchar_t);
        CopyMemory(cursor, path.c_str(), byte_count);
        cursor += path.size();
        *cursor++ = L'\0';
    }
    *cursor = L'\0';

    GlobalUnlock(drop_handle);
    return drop_handle;
}

HGLOBAL CreateDropEffectHandle(DWORD effect) {
    HGLOBAL handle = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, sizeof(DWORD));
    if (handle == nullptr) {
        LogLastError(L"GlobalAlloc(DragDrop PreferredDropEffect)");
        return nullptr;
    }

    void* block = GlobalLock(handle);
    if (block == nullptr) {
        LogLastError(L"GlobalLock(DragDrop PreferredDropEffect)");
        GlobalFree(handle);
        return nullptr;
    }

    *reinterpret_cast<DWORD*>(block) = effect;
    GlobalUnlock(handle);
    return handle;
}

FORMATETC CopyFormatEtc(const FORMATETC& source) {
    FORMATETC copy = source;
    if (source.ptd != nullptr) {
        copy.ptd = static_cast<DVTARGETDEVICE*>(CoTaskMemAlloc(sizeof(DVTARGETDEVICE)));
        if (copy.ptd != nullptr) {
            CopyMemory(copy.ptd, source.ptd, sizeof(DVTARGETDEVICE));
        }
    }
    return copy;
}

std::vector<std::wstring> ExtractDropPathsFromHandle(HDROP drop_handle) {
    std::vector<std::wstring> paths;
    if (drop_handle == nullptr) {
        return paths;
    }

    const UINT file_count = DragQueryFileW(drop_handle, 0xFFFFFFFF, nullptr, 0);
    paths.reserve(file_count);
    for (UINT i = 0; i < file_count; ++i) {
        const UINT length = DragQueryFileW(drop_handle, i, nullptr, 0);
        if (length == 0U) {
            continue;
        }

        std::wstring path(length, L'\0');
        if (DragQueryFileW(drop_handle, i, path.data(), length + 1U) == 0U) {
            continue;
        }
        paths.push_back(path);
    }

    return paths;
}

std::wstring JoinPathForDropTarget(const std::wstring& folder, const std::wstring& leaf_name) {
    std::wstring path = folder;
    if (!path.empty() && path.back() != L'\\') {
        path.push_back(L'\\');
    }
    path.append(leaf_name);
    return path;
}

bool TryFindDropNameConflict(
    const std::vector<std::wstring>& source_paths,
    const std::wstring& destination_path,
    std::wstring* conflict_name) {
    std::unordered_set<std::wstring> seen_names;
    for (const std::wstring& source_path : source_paths) {
        const wchar_t* leaf_ptr = PathFindFileNameW(source_path.c_str());
        if (leaf_ptr == nullptr || *leaf_ptr == L'\0') {
            continue;
        }

        const std::wstring leaf_name = leaf_ptr;
        const std::wstring normalized_leaf = ToLowerCopy(leaf_name);
        if (!seen_names.insert(normalized_leaf).second) {
            if (conflict_name != nullptr) {
                *conflict_name = leaf_name;
            }
            return true;
        }

        const std::wstring destination_candidate = JoinPathForDropTarget(destination_path, leaf_name);
        if (GetFileAttributesW(destination_candidate.c_str()) != INVALID_FILE_ATTRIBUTES) {
            if (conflict_name != nullptr) {
                *conflict_name = leaf_name;
            }
            return true;
        }
    }

    return false;
}

std::wstring ResolveDropLeafName(const std::wstring& candidate_name, UINT index) {
    const wchar_t* leaf_ptr = PathFindFileNameW(candidate_name.c_str());
    std::wstring leaf_name = (leaf_ptr != nullptr) ? std::wstring(leaf_ptr) : candidate_name;
    if (leaf_name.empty()) {
        wchar_t fallback_name[64] = {};
        swprintf_s(fallback_name, L"dropped_item_%u.bin", index + 1U);
        leaf_name = fallback_name;
    }
    return leaf_name;
}

bool WriteBufferToFile(const std::wstring& path, const BYTE* data, size_t bytes) {
    HANDLE file_handle = CreateFileW(
        path.c_str(),
        GENERIC_WRITE,
        0,
        nullptr,
        CREATE_NEW,
        FILE_ATTRIBUTE_NORMAL,
        nullptr);
    if (file_handle == INVALID_HANDLE_VALUE) {
        LogLastError(L"CreateFileW(DropWriteBuffer)");
        return false;
    }

    bool write_ok = true;
    size_t written_total = 0;
    while (written_total < bytes) {
        const size_t chunk = (std::min)(bytes - written_total, static_cast<size_t>(1U << 20));
        DWORD written_now = 0;
        if (!WriteFile(
                file_handle,
                data + written_total,
                static_cast<DWORD>(chunk),
                &written_now,
                nullptr)) {
            LogLastError(L"WriteFile(DropWriteBuffer)");
            write_ok = false;
            break;
        }
        if (written_now == 0U) {
            write_ok = false;
            break;
        }
        written_total += static_cast<size_t>(written_now);
    }

    CloseHandle(file_handle);
    if (!write_ok || written_total != bytes) {
        DeleteFileW(path.c_str());
        return false;
    }
    return true;
}

bool WriteIStreamToFile(IStream* stream, const std::wstring& path) {
    if (stream == nullptr) {
        return false;
    }

    LARGE_INTEGER origin = {};
    (void)stream->Seek(origin, STREAM_SEEK_SET, nullptr);

    HANDLE file_handle = CreateFileW(
        path.c_str(),
        GENERIC_WRITE,
        0,
        nullptr,
        CREATE_NEW,
        FILE_ATTRIBUTE_NORMAL,
        nullptr);
    if (file_handle == INVALID_HANDLE_VALUE) {
        LogLastError(L"CreateFileW(DropWriteIStream)");
        return false;
    }

    bool write_ok = true;
    constexpr ULONG kBufferBytes = 64U * 1024U;
    std::vector<BYTE> buffer(kBufferBytes);
    while (true) {
        ULONG bytes_read = 0;
        const HRESULT read_result = stream->Read(buffer.data(), kBufferBytes, &bytes_read);
        if (FAILED(read_result)) {
            LogHResult(L"IStream::Read(DropWriteIStream)", read_result);
            write_ok = false;
            break;
        }
        if (bytes_read == 0U) {
            break;
        }

        DWORD bytes_written = 0;
        if (!WriteFile(file_handle, buffer.data(), bytes_read, &bytes_written, nullptr) ||
            bytes_written != bytes_read) {
            LogLastError(L"WriteFile(DropWriteIStream)");
            write_ok = false;
            break;
        }
    }

    CloseHandle(file_handle);
    if (!write_ok) {
        DeleteFileW(path.c_str());
    }
    return write_ok;
}

bool WriteHGlobalToFile(HGLOBAL global_handle, const std::wstring& path) {
    if (global_handle == nullptr) {
        return false;
    }

    const SIZE_T bytes = GlobalSize(global_handle);
    const void* block = GlobalLock(global_handle);
    if (block == nullptr && bytes > 0U) {
        LogLastError(L"GlobalLock(DropWriteHGlobal)");
        return false;
    }

    const bool write_ok = WriteBufferToFile(
        path,
        static_cast<const BYTE*>(block),
        static_cast<size_t>(bytes));
    if (block != nullptr) {
        GlobalUnlock(global_handle);
    }
    return write_ok;
}

bool WriteIStorageToFile(IStorage* storage, const std::wstring& path) {
    if (storage == nullptr) {
        return false;
    }

    IStorage* destination_storage = nullptr;
    const HRESULT create_result = StgCreateDocfile(
        path.c_str(),
        STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_DIRECT,
        0,
        &destination_storage);
    if (FAILED(create_result) || destination_storage == nullptr) {
        LogHResult(L"StgCreateDocfile(DropWriteIStorage)", create_result);
        return false;
    }

    const HRESULT copy_result = storage->CopyTo(0, nullptr, nullptr, destination_storage);
    if (FAILED(copy_result)) {
        LogHResult(L"IStorage::CopyTo(DropWriteIStorage)", copy_result);
        destination_storage->Release();
        DeleteFileW(path.c_str());
        return false;
    }

    const HRESULT commit_result = destination_storage->Commit(STGC_DEFAULT);
    destination_storage->Release();
    if (FAILED(commit_result)) {
        LogHResult(L"IStorage::Commit(DropWriteIStorage)", commit_result);
        DeleteFileW(path.c_str());
        return false;
    }
    return true;
}

bool WriteStgMediumToFile(const STGMEDIUM& medium, const std::wstring& path) {
    switch (medium.tymed) {
    case TYMED_ISTREAM:
        return WriteIStreamToFile(medium.pstm, path);

    case TYMED_HGLOBAL:
        return WriteHGlobalToFile(medium.hGlobal, path);

    case TYMED_FILE:
        if (medium.lpszFileName == nullptr || medium.lpszFileName[0] == L'\0') {
            return false;
        }
        if (CopyFileW(medium.lpszFileName, path.c_str(), TRUE)) {
            return true;
        }
        LogLastError(L"CopyFileW(DropWriteTYMED_FILE)");
        return false;

    case TYMED_ISTORAGE:
        return WriteIStorageToFile(medium.pstg, path);

    default:
        return false;
    }
}

std::wstring TrimWhitespace(const std::wstring& value) {
    const size_t begin = value.find_first_not_of(L" \t\r\n");
    if (begin == std::wstring::npos) {
        return std::wstring();
    }
    const size_t end = value.find_last_not_of(L" \t\r\n");
    return value.substr(begin, (end - begin) + 1U);
}

void AppendExistingPathsFromTextBlock(const std::wstring& text, std::vector<std::wstring>* paths) {
    if (paths == nullptr || text.empty()) {
        return;
    }

    size_t start = 0;
    while (start < text.size()) {
        size_t end = text.find_first_of(L"\r\n", start);
        if (end == std::wstring::npos) {
            end = text.size();
        }

        std::wstring candidate = TrimWhitespace(text.substr(start, end - start));
        if (candidate.size() >= 2U && candidate.front() == L'\"' && candidate.back() == L'\"') {
            candidate = candidate.substr(1, candidate.size() - 2U);
            candidate = TrimWhitespace(candidate);
        }
        if (!candidate.empty() && GetFileAttributesW(candidate.c_str()) != INVALID_FILE_ATTRIBUTES) {
            paths->push_back(candidate);
        }

        if (end == text.size()) {
            break;
        }
        start = end + 1U;
        while (start < text.size() && (text[start] == L'\r' || text[start] == L'\n')) {
            ++start;
        }
    }
}

class FormatEtcEnumerator final : public IEnumFORMATETC {
public:
    explicit FormatEtcEnumerator(const std::vector<FORMATETC>& formats)
        : ref_count_(1), index_(0) {
        formats_.reserve(formats.size());
        for (const FORMATETC& format : formats) {
            formats_.push_back(CopyFormatEtc(format));
        }
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** object) override {
        if (object == nullptr) {
            return E_POINTER;
        }
        *object = nullptr;
        if (IsEqualIID(iid, IID_IUnknown) || IsEqualIID(iid, IID_IEnumFORMATETC)) {
            *object = static_cast<IEnumFORMATETC*>(this);
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override {
        return ++ref_count_;
    }

    ULONG STDMETHODCALLTYPE Release() override {
        const ULONG count = --ref_count_;
        if (count == 0U) {
            delete this;
        }
        return count;
    }

    HRESULT STDMETHODCALLTYPE Next(ULONG celt, FORMATETC* elements, ULONG* fetched) override {
        if (elements == nullptr) {
            return E_POINTER;
        }
        if (fetched != nullptr) {
            *fetched = 0;
        }
        if (celt == 0U) {
            return S_OK;
        }
        if (fetched == nullptr && celt > 1U) {
            return E_INVALIDARG;
        }

        ULONG copied = 0;
        while (copied < celt && index_ < formats_.size()) {
            elements[copied] = CopyFormatEtc(formats_[index_]);
            ++copied;
            ++index_;
        }

        if (fetched != nullptr) {
            *fetched = copied;
        }
        return (copied == celt) ? S_OK : S_FALSE;
    }

    HRESULT STDMETHODCALLTYPE Skip(ULONG celt) override {
        index_ = (std::min)(formats_.size(), index_ + static_cast<size_t>(celt));
        return (index_ < formats_.size()) ? S_OK : S_FALSE;
    }

    HRESULT STDMETHODCALLTYPE Reset() override {
        index_ = 0;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Clone(IEnumFORMATETC** clone) override {
        if (clone == nullptr) {
            return E_POINTER;
        }
        *clone = nullptr;

        auto* enumerator = new (std::nothrow) FormatEtcEnumerator(formats_);
        if (enumerator == nullptr) {
            return E_OUTOFMEMORY;
        }
        enumerator->index_ = index_;
        *clone = enumerator;
        return S_OK;
    }

private:
    ~FormatEtcEnumerator() {
        for (FORMATETC& format : formats_) {
            if (format.ptd != nullptr) {
                CoTaskMemFree(format.ptd);
                format.ptd = nullptr;
            }
        }
    }

    std::atomic<ULONG> ref_count_;
    std::vector<FORMATETC> formats_;
    size_t index_;
};

class ExternalDragDataObject final : public IDataObject {
public:
    ExternalDragDataObject(HGLOBAL drop_files_handle, HGLOBAL drop_effect_handle)
        : ref_count_(1) {
        AddHGlobalFormat(CF_HDROP, drop_files_handle);

        if (drop_effect_handle != nullptr) {
            const UINT drop_effect_format = RegisterClipboardFormatW(CFSTR_PREFERREDDROPEFFECT);
            if (drop_effect_format == 0U) {
                LogLastError(L"RegisterClipboardFormatW(CFSTR_PREFERREDDROPEFFECT)");
                GlobalFree(drop_effect_handle);
            } else {
                AddHGlobalFormat(static_cast<CLIPFORMAT>(drop_effect_format), drop_effect_handle);
            }
        }
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** object) override {
        if (object == nullptr) {
            return E_POINTER;
        }
        *object = nullptr;

        if (IsEqualIID(iid, IID_IUnknown) || IsEqualIID(iid, IID_IDataObject)) {
            *object = static_cast<IDataObject*>(this);
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override {
        return ++ref_count_;
    }

    ULONG STDMETHODCALLTYPE Release() override {
        const ULONG count = --ref_count_;
        if (count == 0U) {
            delete this;
        }
        return count;
    }

    HRESULT STDMETHODCALLTYPE GetData(FORMATETC* format_etc_in, STGMEDIUM* medium) override {
        if (format_etc_in == nullptr || medium == nullptr) {
            return E_POINTER;
        }

        for (size_t i = 0; i < formats_.size(); ++i) {
            const FORMATETC& stored_format = formats_[i];
            if (format_etc_in->cfFormat != stored_format.cfFormat) {
                continue;
            }
            if ((format_etc_in->tymed & stored_format.tymed) == 0) {
                continue;
            }
            if (format_etc_in->dwAspect != stored_format.dwAspect) {
                continue;
            }
            if (format_etc_in->lindex != -1 && format_etc_in->lindex != stored_format.lindex) {
                continue;
            }

            if (stored_media_[i].tymed != TYMED_HGLOBAL || stored_media_[i].hGlobal == nullptr) {
                return DV_E_TYMED;
            }

            HGLOBAL copy = DuplicateHGlobal(stored_media_[i].hGlobal);
            if (copy == nullptr) {
                return E_OUTOFMEMORY;
            }

            medium->tymed = TYMED_HGLOBAL;
            medium->hGlobal = copy;
            medium->pUnkForRelease = nullptr;
            return S_OK;
        }

        return DV_E_FORMATETC;
    }

    HRESULT STDMETHODCALLTYPE GetDataHere(FORMATETC* format_etc, STGMEDIUM* medium) override {
        (void)format_etc;
        (void)medium;
        return DATA_E_FORMATETC;
    }

    HRESULT STDMETHODCALLTYPE QueryGetData(FORMATETC* format_etc) override {
        if (format_etc == nullptr) {
            return E_POINTER;
        }

        for (const FORMATETC& stored_format : formats_) {
            if (format_etc->cfFormat != stored_format.cfFormat) {
                continue;
            }
            if ((format_etc->tymed & stored_format.tymed) == 0) {
                continue;
            }
            if (format_etc->dwAspect != stored_format.dwAspect) {
                continue;
            }
            if (format_etc->lindex != -1 && format_etc->lindex != stored_format.lindex) {
                continue;
            }
            return S_OK;
        }

        return DV_E_FORMATETC;
    }

    HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc(FORMATETC* format_etc_in, FORMATETC* format_etc_out) override {
        (void)format_etc_in;
        if (format_etc_out == nullptr) {
            return E_POINTER;
        }
        format_etc_out->ptd = nullptr;
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE SetData(FORMATETC* format_etc, STGMEDIUM* medium, BOOL release) override {
        (void)format_etc;
        (void)medium;
        (void)release;
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE EnumFormatEtc(DWORD direction, IEnumFORMATETC** enum_format_etc) override {
        if (enum_format_etc == nullptr) {
            return E_POINTER;
        }
        *enum_format_etc = nullptr;

        if (direction != DATADIR_GET) {
            return E_NOTIMPL;
        }

        auto* enumerator = new (std::nothrow) FormatEtcEnumerator(formats_);
        if (enumerator == nullptr) {
            return E_OUTOFMEMORY;
        }
        *enum_format_etc = enumerator;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE DAdvise(
        FORMATETC* format_etc,
        DWORD advf,
        IAdviseSink* advise_sink,
        DWORD* connection) override {
        (void)format_etc;
        (void)advf;
        (void)advise_sink;
        (void)connection;
        return OLE_E_ADVISENOTSUPPORTED;
    }

    HRESULT STDMETHODCALLTYPE DUnadvise(DWORD connection) override {
        (void)connection;
        return OLE_E_ADVISENOTSUPPORTED;
    }

    HRESULT STDMETHODCALLTYPE EnumDAdvise(IEnumSTATDATA** enum_advise) override {
        (void)enum_advise;
        return OLE_E_ADVISENOTSUPPORTED;
    }

private:
    ~ExternalDragDataObject() {
        for (STGMEDIUM& medium : stored_media_) {
            ReleaseStgMedium(&medium);
        }
    }

    void AddHGlobalFormat(CLIPFORMAT format, HGLOBAL global_handle) {
        if (global_handle == nullptr) {
            return;
        }

        FORMATETC format_etc = {};
        format_etc.cfFormat = format;
        format_etc.ptd = nullptr;
        format_etc.dwAspect = DVASPECT_CONTENT;
        format_etc.lindex = -1;
        format_etc.tymed = TYMED_HGLOBAL;
        formats_.push_back(format_etc);

        STGMEDIUM medium = {};
        medium.tymed = TYMED_HGLOBAL;
        medium.hGlobal = global_handle;
        medium.pUnkForRelease = nullptr;
        stored_media_.push_back(medium);
    }

    std::atomic<ULONG> ref_count_;
    std::vector<FORMATETC> formats_;
    std::vector<STGMEDIUM> stored_media_;
};

class ExternalDragDropSource final : public IDropSource {
public:
    ExternalDragDropSource() : ref_count_(1) {}

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** object) override {
        if (object == nullptr) {
            return E_POINTER;
        }
        *object = nullptr;

        if (IsEqualIID(iid, IID_IUnknown) || IsEqualIID(iid, IID_IDropSource)) {
            *object = static_cast<IDropSource*>(this);
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override {
        return ++ref_count_;
    }

    ULONG STDMETHODCALLTYPE Release() override {
        const ULONG count = --ref_count_;
        if (count == 0U) {
            delete this;
        }
        return count;
    }

    HRESULT STDMETHODCALLTYPE QueryContinueDrag(BOOL escape_pressed, DWORD key_state) override {
        if (escape_pressed != FALSE) {
            return DRAGDROP_S_CANCEL;
        }
        if ((key_state & MK_LBUTTON) == 0) {
            return DRAGDROP_S_DROP;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GiveFeedback(DWORD effect) override {
        (void)effect;
        return DRAGDROP_S_USEDEFAULTCURSORS;
    }

private:
    ~ExternalDragDropSource() = default;

    std::atomic<ULONG> ref_count_;
};

}  // namespace

namespace fileexplorer {

class FileListView::OleDropTarget final : public IDropTarget {
public:
    explicit OleDropTarget(FileListView* owner) : ref_count_(1), owner_(owner) {}

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** object) override {
        if (object == nullptr) {
            return E_POINTER;
        }
        *object = nullptr;
        if (IsEqualIID(iid, IID_IUnknown) || IsEqualIID(iid, IID_IDropTarget)) {
            *object = static_cast<IDropTarget*>(this);
            AddRef();
            return S_OK;
        }
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override {
        return ++ref_count_;
    }

    ULONG STDMETHODCALLTYPE Release() override {
        const ULONG count = --ref_count_;
        if (count == 0U) {
            delete this;
        }
        return count;
    }

    HRESULT STDMETHODCALLTYPE DragEnter(IDataObject* data_object, DWORD key_state, POINTL point, DWORD* effect) override {
        if (effect == nullptr) {
            return E_INVALIDARG;
        }
        (void)point;
        if (owner_ == nullptr || owner_->hwnd_ == nullptr || data_object == nullptr) {
            current_data_acceptable_ = false;
            *effect = DROPEFFECT_NONE;
            return S_OK;
        }

        const DWORD desired_effect = owner_->ComputeDropEffectForDataObject(data_object, key_state);
        const DWORD allowed_effects = (*effect) & (DROPEFFECT_COPY | DROPEFFECT_MOVE);
        if ((desired_effect & allowed_effects) != 0U) {
            current_data_acceptable_ = true;
            *effect = desired_effect & allowed_effects;
            return S_OK;
        }
        if ((allowed_effects & DROPEFFECT_COPY) != 0U && desired_effect != DROPEFFECT_NONE) {
            current_data_acceptable_ = true;
            *effect = DROPEFFECT_COPY;
            return S_OK;
        }
        if ((allowed_effects & DROPEFFECT_MOVE) != 0U && desired_effect != DROPEFFECT_NONE) {
            current_data_acceptable_ = true;
            *effect = DROPEFFECT_MOVE;
            return S_OK;
        }
        current_data_acceptable_ = false;
        *effect = DROPEFFECT_NONE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE DragOver(DWORD key_state, POINTL point, DWORD* effect) override {
        if (effect == nullptr) {
            return E_INVALIDARG;
        }
        (void)point;
        if (owner_ == nullptr || owner_->hwnd_ == nullptr || !current_data_acceptable_) {
            *effect = DROPEFFECT_NONE;
            return S_OK;
        }
        if ((*effect & (DROPEFFECT_COPY | DROPEFFECT_MOVE)) == 0U) {
            *effect = DROPEFFECT_NONE;
            return S_OK;
        }

        DWORD desired_effect = DROPEFFECT_COPY;
        if ((key_state & MK_SHIFT) != 0U && (key_state & MK_CONTROL) == 0U) {
            desired_effect = DROPEFFECT_MOVE;
        }
        if ((*effect & desired_effect) != 0U) {
            *effect = desired_effect;
            return S_OK;
        }
        if ((*effect & DROPEFFECT_COPY) != 0U) {
            *effect = DROPEFFECT_COPY;
            return S_OK;
        }
        if ((*effect & DROPEFFECT_MOVE) != 0U) {
            *effect = DROPEFFECT_MOVE;
            return S_OK;
        }
        *effect = DROPEFFECT_NONE;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE DragLeave() override {
        current_data_acceptable_ = false;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Drop(IDataObject* data_object, DWORD key_state, POINTL point, DWORD* effect) override {
        if (effect == nullptr) {
            return E_INVALIDARG;
        }
        current_data_acceptable_ = false;
        if (owner_ == nullptr || owner_->hwnd_ == nullptr || data_object == nullptr) {
            *effect = DROPEFFECT_NONE;
            return S_OK;
        }

        POINT client_point = {point.x, point.y};
        ScreenToClient(owner_->hwnd_, &client_point);

        DWORD performed_effect = DROPEFFECT_NONE;
        if (!owner_->HandleOleDataObjectDrop(data_object, key_state, &client_point, &performed_effect)) {
            *effect = DROPEFFECT_NONE;
            return S_OK;
        }

        const DWORD allowed_effects = (*effect) & (DROPEFFECT_COPY | DROPEFFECT_MOVE);
        *effect = performed_effect & allowed_effects;
        if (*effect == DROPEFFECT_NONE && (allowed_effects & DROPEFFECT_COPY) != 0U) {
            *effect = DROPEFFECT_COPY;
        }
        return S_OK;
    }

private:
    ~OleDropTarget() = default;

    std::atomic<ULONG> ref_count_;
    FileListView* owner_;
    bool current_data_acceptable_{false};
};

void FileListView::FontDeleter::operator()(HFONT font) const noexcept {
    if (font != nullptr) {
        DeleteObject(font);
    }
}

void FileListView::ImageListDeleter::operator()(HIMAGELIST image_list) const noexcept {
    if (image_list != nullptr) {
        ImageList_Destroy(image_list);
    }
}

FileListView::FileListView() = default;

FileListView::~FileListView() {
    EndInlineRenameControl();
    ClearShellContextMenu();
    RevokeOleDropTarget();
    if (header_hwnd_ != nullptr) {
        RemoveWindowSubclass(header_hwnd_, &FileListView::HeaderEraseSubclassProc, kHeaderEraseSubclassId);
        header_hwnd_ = nullptr;
    }
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
    EnsureUiFonts();
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
    if (hwnd_ != nullptr && detail_row_height_image_list_ != nullptr) {
        ListView_SetImageList(hwnd_, system_image_list_, LVSIL_SMALL);
    }
    detail_row_height_image_list_.reset();
    detail_row_height_image_height_ = 0;
    detail_row_height_icon_index_map_.clear();
    divider_font_.reset();
    divider_font_height_ = 0;
    list_font_.reset();
    list_font_height_ = 0;
    header_font_.reset();
    header_font_height_ = 0;
    EnsureUiFonts();
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

    // Incremental refresh keeps current content visible and should avoid spinner flicker.
    SetLoadingState(!incremental_refresh);
    PostStatusUpdate();
}

void FileListView::ApplyLoadedFolderEntries(std::vector<FileEntry> entries, bool incremental_refresh) {
    std::vector<std::wstring> selected_names;
    std::wstring focused_name;
    int top_index = (hwnd_ != nullptr) ? ListView_GetTopIndex(hwnd_) : 0;
    bool list_changed = true;
    if (incremental_refresh) {
        CaptureViewState(&selected_names, &focused_name, &top_index);
        list_changed = ApplyRefreshDelta(std::move(entries));
    } else {
        base_entries_ = std::move(entries);
    }

    if (!list_changed) {
        SetLoadingState(false);
        pending_incremental_refresh_ = false;
        PostStatusUpdate();
        return;
    }

    SortBaseEntries();
    BuildDisplayEntriesWithGroups();
    UpdateSortIndicators();
    UpdateItemCountAndRefresh(!incremental_refresh);
    if (incremental_refresh) {
        RestoreViewState(selected_names, focused_name, top_index);
    } else if (has_pending_view_state_snapshot_) {
        RestoreViewState(
            pending_view_state_snapshot_.selected_names,
            pending_view_state_snapshot_.focused_name,
            pending_view_state_snapshot_.top_index);
        has_pending_view_state_snapshot_ = false;
        pending_view_state_snapshot_ = ViewStateSnapshot{};
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
    if (has_pending_view_state_snapshot_) {
        RestoreViewState(
            pending_view_state_snapshot_.selected_names,
            pending_view_state_snapshot_.focused_name,
            pending_view_state_snapshot_.top_index);
        has_pending_view_state_snapshot_ = false;
        pending_view_state_snapshot_ = ViewStateSnapshot{};
    }
    SetLoadingState(false);
    PostStatusUpdate();
    return true;
}

bool FileListView::CaptureViewStateSnapshot(ViewStateSnapshot* snapshot) const {
    if (snapshot == nullptr || hwnd_ == nullptr) {
        return false;
    }

    snapshot->selected_names.clear();
    snapshot->focused_name.clear();
    snapshot->top_index = 0;
    CaptureViewState(&snapshot->selected_names, &snapshot->focused_name, &snapshot->top_index);
    return true;
}

void FileListView::SetPendingViewStateSnapshot(const ViewStateSnapshot* snapshot) {
    if (snapshot == nullptr) {
        has_pending_view_state_snapshot_ = false;
        pending_view_state_snapshot_ = ViewStateSnapshot{};
        return;
    }

    pending_view_state_snapshot_ = *snapshot;
    has_pending_view_state_snapshot_ = true;
}

void FileListView::ClearPendingViewStateSnapshot() {
    has_pending_view_state_snapshot_ = false;
    pending_view_state_snapshot_ = ViewStateSnapshot{};
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

            const bool is_hidden = (find_data.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN) != 0;
            const bool is_system = (find_data.dwFileAttributes & FILE_ATTRIBUTE_SYSTEM) != 0;
            if ((!show_hidden_files_ && is_hidden) || is_system) {
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

void FileListView::SetShowHiddenFiles(bool show_hidden_files) {
    show_hidden_files_ = show_hidden_files;
}

bool FileListView::show_hidden_files() const noexcept {
    return show_hidden_files_;
}

void FileListView::SetShowExtensions(bool show_extensions) {
    if (show_extensions_ == show_extensions) {
        return;
    }
    show_extensions_ = show_extensions;
    ApplyColumnMode();
}

bool FileListView::show_extensions() const noexcept {
    return show_extensions_;
}

void FileListView::SetDetailScalePercent(int detail_scale_percent) {
    const int clamped = ClampDetailScalePercent(detail_scale_percent);
    if (detail_scale_percent_ == clamped) {
        return;
    }

    detail_scale_percent_ = clamped;
    if (hwnd_ != nullptr && detail_row_height_image_list_ != nullptr) {
        ListView_SetImageList(hwnd_, system_image_list_, LVSIL_SMALL);
    }
    detail_row_height_image_list_.reset();
    detail_row_height_image_height_ = 0;
    detail_row_height_icon_index_map_.clear();

    list_font_.reset();
    list_font_height_ = 0;
    EnsureUiFonts();

    if (hwnd_ != nullptr) {
        InvalidateRect(hwnd_, nullptr, TRUE);
    }
}

int FileListView::detail_scale_percent() const noexcept {
    return detail_scale_percent_;
}

void FileListView::SetSortOrder(SortColumn sort_column, SortDirection sort_direction) {
    if (sort_column_ == sort_column && sort_direction_ == sort_direction) {
        return;
    }
    sort_column_ = sort_column;
    sort_direction_ = sort_direction;
    SortBaseEntries();
    BuildDisplayEntriesWithGroups();
    UpdateSortIndicators();
    UpdateItemCountAndRefresh(false);
}

SortColumn FileListView::sort_column() const noexcept {
    return sort_column_;
}

SortDirection FileListView::sort_direction() const noexcept {
    return sort_direction_;
}

void FileListView::SetColumnWidthsLogical(const ColumnWidthsLogical& widths) {
    for (int i = 0; i < kColumnCount; ++i) {
        column_widths_logical_[i] = ClampColumnWidthLogical(i, widths[static_cast<size_t>(i)]);
    }
    ApplyColumnMode();
}

FileListView::ColumnWidthsLogical FileListView::GetColumnWidthsLogical() const {
    return column_widths_logical_;
}

FileListView::ColumnWidthsLogical FileListView::DefaultColumnWidthsLogical() {
    return ColumnWidthsLogical{{260, 60, 140, 90, 320}};
}

bool FileListView::HandleNotify(LPARAM l_param, LRESULT* result) {
    auto* header = reinterpret_cast<NMHDR*>(l_param);
    if (header == nullptr) {
        return false;
    }

    HWND list_header = (hwnd_ != nullptr) ? ListView_GetHeader(hwnd_) : nullptr;
    if (list_header != nullptr && header->hwndFrom == list_header) {
        if (header->code == NM_CUSTOMDRAW) {
            return HandleHeaderCustomDraw(reinterpret_cast<NMCUSTOMDRAW*>(l_param), result);
        }

        if (header->code == HDN_ENDTRACKA ||
            header->code == HDN_ENDTRACKW ||
            header->code == HDN_ITEMCHANGEDA ||
            header->code == HDN_ITEMCHANGEDW) {
            CaptureCurrentColumnWidths();
            return false;
        }
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

    case LVN_BEGINDRAG:
        if (HandleBeginDrag(reinterpret_cast<NMLISTVIEW*>(l_param))) {
            if (result != nullptr) {
                *result = 0;
            }
            return true;
        }
        break;

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

LRESULT CALLBACK FileListView::RenameEditSubclassProc(
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

    switch (message) {
    case WM_KEYDOWN:
        if (w_param == VK_RETURN) {
            if (self->CommitInlineRename()) {
                return 0;
            }
            return 0;
        }
        if (w_param == VK_ESCAPE) {
            self->CancelInlineRename();
            return 0;
        }
        break;

    case WM_KILLFOCUS:
        self->CommitInlineRename();
        break;

    case WM_NCDESTROY:
        RemoveWindowSubclass(hwnd, &FileListView::RenameEditSubclassProc, kRenameEditSubclassId);
        break;

    default:
        break;
    }

    return DefSubclassProc(hwnd, message, w_param, l_param);
}

LRESULT CALLBACK FileListView::HeaderEraseSubclassProc(
    HWND hwnd,
    UINT message,
    WPARAM w_param,
    LPARAM l_param,
    UINT_PTR subclass_id,
    DWORD_PTR ref_data) {
    (void)subclass_id;
    (void)l_param;

    auto* self = reinterpret_cast<FileListView*>(ref_data);
    if (self == nullptr) {
        return DefSubclassProc(hwnd, message, w_param, l_param);
    }

    switch (message) {
    case WM_ERASEBKGND: {
        HDC hdc = reinterpret_cast<HDC>(w_param);
        if (hdc != nullptr) {
            RECT rect = {};
            if (GetClientRect(hwnd, &rect)) {
                HBRUSH brush = CreateSolidBrush(kHeaderBackgroundColor);
                if (brush != nullptr) {
                    FillRect(hdc, &rect, brush);
                    DeleteObject(brush);
                }
            }
        }
        return 1;
    }

    case WM_NCDESTROY:
        RemoveWindowSubclass(hwnd, &FileListView::HeaderEraseSubclassProc, kHeaderEraseSubclassId);
        if (self->header_hwnd_ == hwnd) {
            self->header_hwnd_ = nullptr;
        }
        break;

    default:
        break;
    }

    return DefSubclassProc(hwnd, message, w_param, l_param);
}

bool FileListView::CreateListViewControl() {
    constexpr DWORD style =
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | LVS_REPORT | LVS_OWNERDATA | LVS_SHOWSELALWAYS |
        LVS_SHAREIMAGELISTS;

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

    DragAcceptFiles(hwnd_, TRUE);
    (void)RegisterOleDropTarget();

    if (!SetWindowSubclass(hwnd_, &FileListView::ListViewSubclassProc, kListViewSubclassId, reinterpret_cast<DWORD_PTR>(this))) {
        LogLastError(L"SetWindowSubclass(FileListView)");
        return false;
    }

    const HRESULT list_theme_result = SetWindowTheme(hwnd_, L"DarkMode_Explorer", nullptr);
    if (FAILED(list_theme_result)) {
        LogHResult(L"SetWindowTheme(FileListView)", list_theme_result);
    }

    const DWORD ex_style = LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_HEADERDRAGDROP;
    ListView_SetExtendedListViewStyle(hwnd_, ex_style);
    ListView_SetBkColor(hwnd_, kListBackgroundColor);
    ListView_SetTextBkColor(hwnd_, kListBackgroundColor);
    ListView_SetTextColor(hwnd_, kListTextColor);
    return true;
}

bool FileListView::RegisterOleDropTarget() {
    if (hwnd_ == nullptr) {
        return false;
    }
    if (ole_drop_target_ != nullptr) {
        return true;
    }

    auto* drop_target = new (std::nothrow) OleDropTarget(this);
    if (drop_target == nullptr) {
        return false;
    }

    const HRESULT register_result = RegisterDragDrop(hwnd_, static_cast<IDropTarget*>(drop_target));
    if (FAILED(register_result)) {
        LogHResult(L"RegisterDragDrop(FileListView)", register_result);
        drop_target->Release();
        return false;
    }

    ole_drop_target_ = drop_target;
    return true;
}

void FileListView::RevokeOleDropTarget() {
    if (hwnd_ != nullptr) {
        const HRESULT revoke_result = RevokeDragDrop(hwnd_);
        if (FAILED(revoke_result) && revoke_result != DRAGDROP_E_NOTREGISTERED) {
            LogHResult(L"RevokeDragDrop(FileListView)", revoke_result);
        }
    }

    if (ole_drop_target_ != nullptr) {
        ole_drop_target_->Release();
        ole_drop_target_ = nullptr;
    }
}

void FileListView::CreateColumns() {
    const struct ColumnSpec {
        const wchar_t* title;
        int format;
    } columns[] = {
        {L"Name", LVCFMT_LEFT},
        {L"Extension", LVCFMT_LEFT},
        {L"Date Modified", LVCFMT_LEFT},
        {L"Size", LVCFMT_RIGHT},
        {L"Path", LVCFMT_LEFT},
    };

    for (int i = 0; i < static_cast<int>(std::size(columns)); ++i) {
        LVCOLUMNW column = {};
        column.mask = LVCF_FMT | LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
        column.fmt = columns[i].format;
        const int logical_width = column_widths_logical_[static_cast<size_t>(i)];
        const int width_px = MulDiv(logical_width, static_cast<int>(dpi_), 96);
        const bool hide_path = (i == 4 && !search_mode_);
        const bool hide_extension = (i == 1 && !show_extensions_);
        column.cx = (hide_path || hide_extension) ? 0 : width_px;
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

    for (int i = 0; i < kColumnCount; ++i) {
        const int logical_width = ClampColumnWidthLogical(i, column_widths_logical_[static_cast<size_t>(i)]);
        const int width_px = MulDiv(logical_width, static_cast<int>(dpi_), 96);
        const bool hide_path = (i == 4 && !search_mode_);
        const bool hide_extension = (i == 1 && !show_extensions_);
        const int applied_width = (hide_path || hide_extension) ? 0 : width_px;
        ListView_SetColumnWidth(hwnd_, i, applied_width);
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

    if (header_hwnd_ != nullptr && header_hwnd_ != header) {
        RemoveWindowSubclass(header_hwnd_, &FileListView::HeaderEraseSubclassProc, kHeaderEraseSubclassId);
        header_hwnd_ = nullptr;
    }
    header_hwnd_ = header;

    // Keep header visual styles disabled so our custom-draw path controls the full
    // dark surface and text colors consistently.
    const HRESULT theme_result = SetWindowTheme(header, L"", L"");
    if (FAILED(theme_result)) {
        LogHResult(L"SetWindowTheme(ListHeader)", theme_result);
    }

    if (!SetWindowSubclass(
            header,
            &FileListView::HeaderEraseSubclassProc,
            kHeaderEraseSubclassId,
            reinterpret_cast<DWORD_PTR>(this))) {
        LogLastError(L"SetWindowSubclass(ListHeaderErase)");
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
    ApplyDetailScaleToListMetrics();
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

void FileListView::EnsureUiFonts() {
    int desired_list_height = -MulDiv(
        kListBodyPointSize * detail_scale_percent_,
        static_cast<int>(dpi_),
        72 * 100);
    if (desired_list_height == 0) {
        desired_list_height = -1;
    }
    const int desired_header_height = -MulDiv(kListHeaderPointSize, static_cast<int>(dpi_), 72);

    if (list_font_ == nullptr || list_font_height_ != desired_list_height) {
        list_font_.reset(CreateUiFontFromPixelHeight(desired_list_height, FW_NORMAL));
        if (list_font_ == nullptr) {
            LogLastError(L"CreateFontW(FileListBody)");
            list_font_height_ = 0;
        } else {
            list_font_height_ = desired_list_height;
        }
    }

    if (header_font_ == nullptr || header_font_height_ != desired_header_height) {
        header_font_.reset(CreateUiFontFromPixelHeight(desired_header_height, FW_SEMIBOLD));
        if (header_font_ == nullptr) {
            LogLastError(L"CreateFontW(FileListHeader)");
            header_font_height_ = 0;
        } else {
            header_font_height_ = desired_header_height;
        }
    }

    if (hwnd_ != nullptr && list_font_ != nullptr) {
        SendMessageW(hwnd_, WM_SETFONT, reinterpret_cast<WPARAM>(list_font_.get()), TRUE);
    }

    const HWND header = (hwnd_ != nullptr) ? ListView_GetHeader(hwnd_) : nullptr;
    if (header != nullptr && header_font_ != nullptr) {
        SendMessageW(header, WM_SETFONT, reinterpret_cast<WPARAM>(header_font_.get()), TRUE);
    }

    ApplyDetailScaleToListMetrics();
}

void FileListView::ApplyDetailScaleToListMetrics() {
    if (hwnd_ == nullptr || system_image_list_ == nullptr) {
        return;
    }

    if (detail_scale_percent_ <= kDefaultDetailScalePercent) {
        if (detail_row_height_image_list_ != nullptr) {
            ListView_SetImageList(hwnd_, system_image_list_, LVSIL_SMALL);
            detail_row_height_image_list_.reset();
            detail_row_height_image_height_ = 0;
            detail_row_height_icon_index_map_.clear();
        }
        return;
    }

    int icon_width = 0;
    int icon_height = 16;
    if (!ImageList_GetIconSize(system_image_list_, &icon_width, &icon_height)) {
        icon_height = 16;
        icon_width = 16;
    }

    const int base_text_height = MulDiv(kListBodyPointSize, static_cast<int>(dpi_), 72);
    const int base_row_height = (std::max)(
        icon_height,
        base_text_height) + MulDiv(kListRowPaddingLogical, static_cast<int>(dpi_), 96);
    const int desired_row_height = (std::max)(
        base_row_height + 1,
        MulDiv(base_row_height, detail_scale_percent_, 100));

    if (detail_row_height_image_list_ != nullptr &&
        detail_row_height_image_height_ == desired_row_height) {
        ListView_SetImageList(hwnd_, detail_row_height_image_list_.get(), LVSIL_SMALL);
        return;
    }

    if (detail_row_height_image_list_ != nullptr) {
        detail_row_height_image_list_.reset();
        detail_row_height_image_height_ = 0;
        detail_row_height_icon_index_map_.clear();
    }

    const int cell_icon_width = (icon_width > 0) ? icon_width : 16;
    UniqueImageList scaled_image_list(ImageList_Create(
        cell_icon_width,
        desired_row_height,
        ILC_COLOR32 | ILC_MASK,
        1,
        1));
    if (scaled_image_list == nullptr) {
        LogLastError(L"ImageList_Create(ScaledSmallIconList)");
        return;
    }

    ListView_SetImageList(hwnd_, scaled_image_list.get(), LVSIL_SMALL);
    detail_row_height_image_list_ = std::move(scaled_image_list);
    detail_row_height_image_height_ = desired_row_height;
    detail_row_height_icon_index_map_.clear();
}

bool FileListView::HandleGetDispInfo(NMLVDISPINFOW* info) {
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
                text = BuildDisplayName(entry);
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
        info->item.iImage = entry.is_group_divider ? -1 : ResolveScaledSmallIconIndex(entry.icon_index);
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

    if (parent_hwnd_ != nullptr && !search_mode_) {
        PostMessageW(
            parent_hwnd_,
            WM_FE_FILELIST_SORT_CHANGED,
            static_cast<WPARAM>(sort_column_),
            static_cast<LPARAM>(sort_direction_ == SortDirection::Ascending ? 0 : 1));
    }
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

        const std::wstring full_path = BuildFullPath(entry);
        if (IsPathCutPending(full_path)) {
            custom_draw->clrText = kListCutPendingTextColor;
            *result = CDRF_DODEFAULT;
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

    const DWORD draw_stage = custom_draw->dwDrawStage;

    if (draw_stage == CDDS_PREPAINT) {
        RECT header_rect = {};
        if (GetClientRect(header, &header_rect)) {
            HBRUSH header_brush = CreateSolidBrush(kHeaderBackgroundColor);
            if (header_brush != nullptr) {
                FillRect(custom_draw->hdc, &header_rect, header_brush);
                DeleteObject(header_brush);
            }
        }

        *result = CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYPOSTPAINT;
        return true;
    }

    if ((draw_stage & CDDS_ITEMPREPAINT) == CDDS_ITEMPREPAINT) {
        const int column_index = static_cast<int>(custom_draw->dwItemSpec);
        RECT rect = custom_draw->rc;
        const bool hot = (custom_draw->uItemState & CDIS_HOT) != 0;

        HBRUSH column_brush = CreateSolidBrush(hot ? kHeaderHoverBackgroundColor : kHeaderBackgroundColor);
        if (column_brush != nullptr) {
            FillRect(custom_draw->hdc, &rect, column_brush);
            DeleteObject(column_brush);
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

    if (draw_stage == CDDS_POSTPAINT) {
        RECT header_rect = {};
        if (!GetClientRect(header, &header_rect)) {
            *result = CDRF_DODEFAULT;
            return true;
        }

        int right_edge = 0;
        const int column_count = Header_GetItemCount(header);
        for (int column = 0; column < column_count; ++column) {
            RECT column_rect = {};
            if (Header_GetItemRect(header, column, &column_rect)) {
                right_edge = (std::max)(right_edge, static_cast<int>(column_rect.right));
            }
        }

        if (right_edge < header_rect.right) {
            RECT trailing_rect = header_rect;
            trailing_rect.left = right_edge;
            HBRUSH trailing_brush = CreateSolidBrush(kHeaderBackgroundColor);
            if (trailing_brush != nullptr) {
                FillRect(custom_draw->hdc, &trailing_rect, trailing_brush);
                DeleteObject(trailing_brush);
            }
        }

        HPEN divider_pen = CreatePen(PS_SOLID, 1, kHeaderDividerColor);
        if (divider_pen != nullptr) {
            HGDIOBJ old_pen = SelectObject(custom_draw->hdc, divider_pen);
            MoveToEx(custom_draw->hdc, header_rect.left, header_rect.bottom - 1, nullptr);
            LineTo(custom_draw->hdc, header_rect.right, header_rect.bottom - 1);
            SelectObject(custom_draw->hdc, old_pen);
            DeleteObject(divider_pen);
        }

        *result = CDRF_DODEFAULT;
        return true;
    }

    return false;
}

bool FileListView::HandleDoubleClick(const NMITEMACTIVATE* item_activate) {
    if (item_activate == nullptr || item_activate->iItem < 0) {
        return false;
    }
    return OpenEntryAtIndex(item_activate->iItem, true);
}

bool FileListView::HandleBeginDrag(const NMLISTVIEW* list_view) {
    if (list_view == nullptr || hwnd_ == nullptr) {
        return false;
    }

    const int begin_index = list_view->iItem;
    if (IsSelectableIndex(begin_index) &&
        (ListView_GetItemState(hwnd_, begin_index, LVIS_SELECTED) & LVIS_SELECTED) == 0) {
        EnsureSingleSelectionAtIndex(begin_index);
    }

    std::vector<int> selected_indices = CollectSelectedSelectableIndices();
    if (selected_indices.empty() && IsSelectableIndex(begin_index)) {
        selected_indices.push_back(begin_index);
    }
    if (selected_indices.empty()) {
        return false;
    }

    const std::vector<std::wstring> selected_paths = CollectPathsForIndices(selected_indices);
    if (selected_paths.empty()) {
        return false;
    }

    EndInlineRenameControl();
    return StartExternalDrag(selected_paths);
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
    case WM_NOTIFY: {
        auto* notify = reinterpret_cast<NMHDR*>(l_param);
        if (notify != nullptr) {
            HWND header = (hwnd_ != nullptr) ? ListView_GetHeader(hwnd_) : nullptr;
            if (header != nullptr && notify->hwndFrom == header) {
                if (notify->code == NM_CUSTOMDRAW) {
                    LRESULT header_result = 0;
                    if (HandleHeaderCustomDraw(reinterpret_cast<NMCUSTOMDRAW*>(l_param), &header_result)) {
                        return header_result;
                    }
                }
                if (notify->code == HDN_ENDTRACKA ||
                    notify->code == HDN_ENDTRACKW ||
                    notify->code == HDN_ITEMCHANGEDA ||
                    notify->code == HDN_ITEMCHANGEDW) {
                    CaptureCurrentColumnWidths();
                }
            }
        }
        break;
    }

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

        if (ctrl_down && (w_param == 'C' || w_param == 'c')) {
            CopySelectionToClipboard(false, ResolveContextTargetIndex(-1));
            return 0;
        }

        if (ctrl_down && (w_param == 'X' || w_param == 'x')) {
            CopySelectionToClipboard(true, ResolveContextTargetIndex(-1));
            return 0;
        }

        if (ctrl_down && (w_param == 'V' || w_param == 'v')) {
            PasteFromClipboard();
            return 0;
        }

        if (w_param == VK_DELETE) {
            DeleteSelection(shift_down, ResolveContextTargetIndex(-1));
            return 0;
        }

        if (w_param == VK_F2) {
            BeginInlineRename(ResolveContextTargetIndex(-1));
            return 0;
        }

        if (w_param == VK_ESCAPE && !ctrl_down && !shift_down && !alt_down) {
            ClearSelection();
            if (hwnd_ != nullptr) {
                SetFocus(hwnd_);
            }
            PostStatusUpdate();
            return 0;
        }

        if (alt_down && w_param == VK_RETURN) {
            const int target_index = ResolveContextTargetIndex(-1);
            if (IsSelectableIndex(target_index)) {
                const std::wstring target_path = BuildFullPath(display_entries_[target_index]);
                FileOps::ShowProperties(hwnd_, target_path);
            }
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
        drag_candidate_active_ = false;
        drag_candidate_index_ = -1;
        drag_candidate_start_ = point;

        if (IsSelectableIndex(item)) {
            drag_candidate_active_ = true;
            drag_candidate_index_ = item;
        }

        if (item < 0) {
            ClearSelection();
            PostStatusUpdate();
        }
        break;
    }

    case WM_MOUSEMOVE: {
        if (drag_candidate_active_ && (w_param & MK_LBUTTON) != 0U) {
            const int delta_x = std::abs(GET_X_LPARAM(l_param) - drag_candidate_start_.x);
            const int delta_y = std::abs(GET_Y_LPARAM(l_param) - drag_candidate_start_.y);
            const int drag_width = (std::max)(1, GetSystemMetrics(SM_CXDRAG));
            const int drag_height = (std::max)(1, GetSystemMetrics(SM_CYDRAG));
            if (delta_x >= drag_width || delta_y >= drag_height) {
                const int begin_index = drag_candidate_index_;
                drag_candidate_active_ = false;
                drag_candidate_index_ = -1;

                if (IsSelectableIndex(begin_index) &&
                    (ListView_GetItemState(hwnd_, begin_index, LVIS_SELECTED) & LVIS_SELECTED) == 0) {
                    EnsureSingleSelectionAtIndex(begin_index);
                }

                std::vector<int> selected_indices = CollectSelectedSelectableIndices();
                if (selected_indices.empty() && IsSelectableIndex(begin_index)) {
                    selected_indices.push_back(begin_index);
                }
                const std::vector<std::wstring> selected_paths = CollectPathsForIndices(selected_indices);
                if (!selected_paths.empty()) {
                    EndInlineRenameControl();
                    StartExternalDrag(selected_paths);
                }

                return 0;
            }
        } else if ((w_param & MK_LBUTTON) == 0U) {
            drag_candidate_active_ = false;
            drag_candidate_index_ = -1;
        }
        break;
    }

    case WM_LBUTTONUP:
    case WM_CANCELMODE:
    case WM_CAPTURECHANGED:
        drag_candidate_active_ = false;
        drag_candidate_index_ = -1;
        break;

    case WM_DROPFILES: {
        HDROP drop_handle = reinterpret_cast<HDROP>(w_param);
        POINT drop_point = {};
        const bool has_client_point = DragQueryPoint(drop_handle, &drop_point) != FALSE;
        const std::vector<std::wstring> source_paths = ExtractDropPathsFromHandle(drop_handle);
        const bool shift_down = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        const bool ctrl_down = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        const bool move_drop = shift_down && !ctrl_down;
        DragFinish(drop_handle);

        const POINT* point_ptr = has_client_point ? &drop_point : nullptr;
        (void)HandleExternalPathsDrop(source_paths, move_drop, point_ptr);
        return 0;
    }

    case WM_XBUTTONUP: {
        if (parent_hwnd_ == nullptr) {
            return TRUE;
        }

        NavCommand command = NavCommand::Back;
        const WORD xbutton = GET_XBUTTON_WPARAM(w_param);
        if (xbutton == XBUTTON1) {
            command = NavCommand::Back;
        } else if (xbutton == XBUTTON2) {
            command = NavCommand::Forward;
        } else {
            return FALSE;
        }

        if (!PostMessageW(parent_hwnd_, WM_FE_NAV_COMMAND, static_cast<WPARAM>(command), 0)) {
            LogLastError(L"PostMessageW(WM_FE_NAV_COMMAND XBUTTON)");
            return FALSE;
        }
        return TRUE;
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
            loading_frame_ = (loading_frame_ + 1) % 8;
            InvalidateRect(hwnd_, nullptr, FALSE);
            return 0;
        }
        break;

    case WM_NCDESTROY:
        RevokeOleDropTarget();
        DragAcceptFiles(hwnd_, FALSE);
        drag_candidate_active_ = false;
        drag_candidate_index_ = -1;
        EndInlineRenameControl();
        ClearShellContextMenu();
        SetLoadingState(false);
        if (detail_row_height_image_list_ != nullptr) {
            ListView_SetImageList(hwnd_, system_image_list_, LVSIL_SMALL);
            detail_row_height_image_list_.reset();
            detail_row_height_image_height_ = 0;
            detail_row_height_icon_index_map_.clear();
        }
        RemoveWindowSubclass(hwnd_, &FileListView::ListViewSubclassProc, kListViewSubclassId);
        break;

    default:
        break;
    }

    return DefSubclassProc(hwnd_, message, w_param, l_param);
}

void FileListView::SortBaseEntries() {
    auto compare = [this](const FileEntry& left, const FileEntry& right) {
        const bool enforce_folders_first = sort_column_ != SortColumn::DateModified;
        if (enforce_folders_first && left.is_folder != right.is_folder) {
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

    const int overlay_width = MulDiv(210, static_cast<int>(dpi_), 96);
    const int overlay_height = MulDiv(44, static_cast<int>(dpi_), 96);
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

    struct SpinnerOffset {
        int x;
        int y;
    };
    static constexpr SpinnerOffset kSpinnerOffsets[8] = {
        {0, -9},
        {6, -6},
        {9, 0},
        {6, 6},
        {0, 9},
        {-6, 6},
        {-9, 0},
        {-6, -6},
    };

    const int spinner_center_x = overlay_rect.left + MulDiv(24, static_cast<int>(dpi_), 96);
    const int spinner_center_y = overlay_rect.top + ((overlay_rect.bottom - overlay_rect.top) / 2);
    const int dot_radius = (std::max)(2, MulDiv(2, static_cast<int>(dpi_), 96));
    const int leader_index = loading_frame_ % 8;

    for (int i = 0; i < 8; ++i) {
        const int offset_x = MulDiv(kSpinnerOffsets[i].x, static_cast<int>(dpi_), 96);
        const int offset_y = MulDiv(kSpinnerOffsets[i].y, static_cast<int>(dpi_), 96);
        const int distance = (i - leader_index + 8) % 8;
        const int intensity = 8 - distance;
        const COLORREF dot_color = LerpColor(kLoadingSpinnerBaseColor, kLoadingSpinnerActiveColor, intensity, 8);

        HBRUSH dot_brush = CreateSolidBrush(dot_color);
        if (dot_brush != nullptr) {
            HGDIOBJ old_pen = SelectObject(hdc, GetStockObject(NULL_PEN));
            HGDIOBJ old_brush = SelectObject(hdc, dot_brush);
            Ellipse(
                hdc,
                spinner_center_x + offset_x - dot_radius,
                spinner_center_y + offset_y - dot_radius,
                spinner_center_x + offset_x + dot_radius + 1,
                spinner_center_y + offset_y + dot_radius + 1);
            SelectObject(hdc, old_brush);
            SelectObject(hdc, old_pen);
            DeleteObject(dot_brush);
        }
    }

    RECT text_rect = overlay_rect;
    text_rect.left = spinner_center_x + MulDiv(16, static_cast<int>(dpi_), 96);
    text_rect.right -= MulDiv(10, static_cast<int>(dpi_), 96);
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, kLoadingOverlayTextColor);
    DrawTextW(hdc, L"Loading...", -1, &text_rect, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
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

bool FileListView::ApplyRefreshDelta(std::vector<FileEntry> incoming_entries) {
    std::unordered_map<std::wstring, FileEntry> incoming_by_key;
    incoming_by_key.reserve(incoming_entries.size());
    for (FileEntry& entry : incoming_entries) {
        incoming_by_key.insert_or_assign(EntryKey(entry), std::move(entry));
    }

    std::vector<FileEntry> merged_entries;
    merged_entries.reserve(incoming_by_key.size());
    bool changed = false;

    for (const FileEntry& current : base_entries_) {
        auto it = incoming_by_key.find(EntryKey(current));
        if (it == incoming_by_key.end()) {
            changed = true;
            continue;
        }

        if (EntriesEquivalent(current, it->second)) {
            merged_entries.push_back(current);
        } else {
            changed = true;
            merged_entries.push_back(std::move(it->second));
        }
        incoming_by_key.erase(it);
    }

    if (!incoming_by_key.empty()) {
        changed = true;
    }

    for (auto& [key, entry] : incoming_by_key) {
        (void)key;
        merged_entries.push_back(std::move(entry));
    }

    if (!changed && merged_entries.size() != base_entries_.size()) {
        changed = true;
    }

    if (!changed) {
        return false;
    }

    base_entries_ = std::move(merged_entries);
    return true;
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

    std::wstring launch_target = full_path;
    bool launch_target_is_folder = false;
    if (TryResolveShortcutTarget(full_path, &launch_target, &launch_target_is_folder)) {
        if (launch_target_is_folder) {
            auto payload = std::make_unique<std::wstring>(NormalizePath(std::move(launch_target)));
            if (!PostMessageW(parent_hwnd_, WM_FE_FILELIST_NAVIGATE, 0, reinterpret_cast<LPARAM>(payload.get()))) {
                LogLastError(L"PostMessageW(WM_FE_FILELIST_NAVIGATE Shortcut)");
                return false;
            }
            payload.release();
            return true;
        }
    }

    const HINSTANCE open_result = ShellExecuteW(
        nullptr,
        L"open",
        launch_target.c_str(),
        nullptr,
        nullptr,
        SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(open_result) <= 32) {
        LogLastError(L"ShellExecuteW(FileOpen)");
        return false;
    }
    return true;
}

bool FileListView::HandleExternalPathsDrop(
    const std::vector<std::wstring>& source_paths,
    bool move_drop,
    const POINT* client_point) {
    if (source_paths.empty() || search_mode_ || current_path_.empty()) {
        return false;
    }

    std::wstring destination_path = current_path_;
    if (client_point != nullptr && hwnd_ != nullptr) {
        LVHITTESTINFO hit = {};
        hit.pt = *client_point;
        const int hit_index = ListView_SubItemHitTest(hwnd_, &hit);
        if (IsSelectableIndex(hit_index) && display_entries_[hit_index].is_folder) {
            destination_path = BuildFullPath(display_entries_[hit_index]);
        }
    }

    std::wstring conflict_name;
    if (TryFindDropNameConflict(source_paths, destination_path, &conflict_name)) {
        std::wstring warning_message = L"Drop blocked: an item with the same name already exists in the destination folder.";
        if (!conflict_name.empty()) {
            warning_message = L"Drop blocked: \"";
            warning_message.append(conflict_name);
            warning_message.append(L"\" already exists in the destination folder.");
        }
        MessageBoxW(hwnd_, warning_message.c_str(), L"FileExplorer", MB_OK | MB_ICONWARNING);
        return false;
    }

    if (!FileOps::ImportPaths(hwnd_, source_paths, destination_path, move_drop)) {
        return false;
    }
    RequestRefresh();
    return true;
}

DWORD FileListView::ComputeDropEffectForDataObject(IDataObject* data_object, DWORD key_state) const {
    if (data_object == nullptr || search_mode_ || current_path_.empty()) {
        return DROPEFFECT_NONE;
    }

    std::vector<std::wstring> source_paths;
    if (ExtractSourcePathsFromDataObject(data_object, &source_paths) && !source_paths.empty()) {
        const bool shift_down = (key_state & MK_SHIFT) != 0U;
        const bool ctrl_down = (key_state & MK_CONTROL) != 0U;
        return (shift_down && !ctrl_down) ? DROPEFFECT_MOVE : DROPEFFECT_COPY;
    }

    const UINT file_descriptor_format = RegisterClipboardFormatW(CFSTR_FILEDESCRIPTORW);
    if (file_descriptor_format != 0U) {
        FORMATETC descriptor_format = {};
        descriptor_format.cfFormat = static_cast<CLIPFORMAT>(file_descriptor_format);
        descriptor_format.ptd = nullptr;
        descriptor_format.dwAspect = DVASPECT_CONTENT;
        descriptor_format.lindex = -1;
        descriptor_format.tymed = TYMED_HGLOBAL;
        if (SUCCEEDED(data_object->QueryGetData(&descriptor_format))) {
            return DROPEFFECT_COPY;
        }
    }
    return DROPEFFECT_NONE;
}

bool FileListView::ExtractSourcePathsFromDataObject(IDataObject* data_object, std::vector<std::wstring>* source_paths) const {
    if (data_object == nullptr || source_paths == nullptr) {
        return false;
    }
    source_paths->clear();

    FORMATETC drop_files_format = {};
    drop_files_format.cfFormat = CF_HDROP;
    drop_files_format.ptd = nullptr;
    drop_files_format.dwAspect = DVASPECT_CONTENT;
    drop_files_format.lindex = -1;
    drop_files_format.tymed = TYMED_HGLOBAL;

    STGMEDIUM medium = {};
    if (SUCCEEDED(data_object->GetData(&drop_files_format, &medium))) {
        if (medium.tymed == TYMED_HGLOBAL && medium.hGlobal != nullptr) {
            *source_paths = ExtractDropPathsFromHandle(reinterpret_cast<HDROP>(medium.hGlobal));
        }
        ReleaseStgMedium(&medium);
        if (!source_paths->empty()) {
            return true;
        }
    }

    const UINT shell_id_list_format = RegisterClipboardFormatW(CFSTR_SHELLIDLIST);
    if (shell_id_list_format != 0U) {
        FORMATETC pidl_format = {};
        pidl_format.cfFormat = static_cast<CLIPFORMAT>(shell_id_list_format);
        pidl_format.ptd = nullptr;
        pidl_format.dwAspect = DVASPECT_CONTENT;
        pidl_format.lindex = -1;
        pidl_format.tymed = TYMED_HGLOBAL;

        STGMEDIUM pidl_medium = {};
        if (SUCCEEDED(data_object->GetData(&pidl_format, &pidl_medium))) {
            if (pidl_medium.tymed == TYMED_HGLOBAL && pidl_medium.hGlobal != nullptr) {
                const SIZE_T bytes = GlobalSize(pidl_medium.hGlobal);
                const BYTE* block = static_cast<const BYTE*>(GlobalLock(pidl_medium.hGlobal));
                if (block != nullptr && bytes >= sizeof(UINT) * 2U) {
                    const CIDA* cida = reinterpret_cast<const CIDA*>(block);
                    const size_t offset_count = static_cast<size_t>(cida->cidl) + 1U;
                    const size_t min_bytes = sizeof(UINT) + (offset_count * sizeof(UINT));
                    if (offset_count >= 1U && min_bytes <= bytes) {
                        bool offsets_valid = true;
                        for (size_t offset_index = 0; offset_index < offset_count; ++offset_index) {
                            if (cida->aoffset[offset_index] >= bytes) {
                                offsets_valid = false;
                                break;
                            }
                        }

                        if (offsets_valid) {
                            const PCIDLIST_ABSOLUTE parent_pidl = reinterpret_cast<PCIDLIST_ABSOLUTE>(block + cida->aoffset[0]);
                            for (UINT i = 0; i < cida->cidl; ++i) {
                                const PCUIDLIST_RELATIVE child_pidl =
                                    reinterpret_cast<PCUIDLIST_RELATIVE>(block + cida->aoffset[i + 1U]);
                                PIDLIST_ABSOLUTE absolute_pidl = ILCombine(parent_pidl, child_pidl);
                                if (absolute_pidl == nullptr) {
                                    continue;
                                }
                                wchar_t path_buffer[MAX_PATH] = {};
                                if (SHGetPathFromIDListW(absolute_pidl, path_buffer) && path_buffer[0] != L'\0') {
                                    source_paths->push_back(path_buffer);
                                }
                                CoTaskMemFree(absolute_pidl);
                            }
                        }
                    }
                }
                if (block != nullptr) {
                    GlobalUnlock(pidl_medium.hGlobal);
                }
            }
            ReleaseStgMedium(&pidl_medium);
            if (!source_paths->empty()) {
                return true;
            }
        }
    }

    const UINT file_name_w_format = RegisterClipboardFormatW(CFSTR_FILENAMEW);
    if (file_name_w_format != 0U) {
        FORMATETC file_name_format = {};
        file_name_format.cfFormat = static_cast<CLIPFORMAT>(file_name_w_format);
        file_name_format.ptd = nullptr;
        file_name_format.dwAspect = DVASPECT_CONTENT;
        file_name_format.lindex = -1;
        file_name_format.tymed = TYMED_HGLOBAL;

        STGMEDIUM file_name_medium = {};
        if (SUCCEEDED(data_object->GetData(&file_name_format, &file_name_medium))) {
            if (file_name_medium.tymed == TYMED_HGLOBAL && file_name_medium.hGlobal != nullptr) {
                const wchar_t* text = static_cast<const wchar_t*>(GlobalLock(file_name_medium.hGlobal));
                if (text != nullptr) {
                    AppendExistingPathsFromTextBlock(text, source_paths);
                    GlobalUnlock(file_name_medium.hGlobal);
                }
            }
            ReleaseStgMedium(&file_name_medium);
            if (!source_paths->empty()) {
                return true;
            }
        }
    }

    FORMATETC unicode_text_format = {};
    unicode_text_format.cfFormat = CF_UNICODETEXT;
    unicode_text_format.ptd = nullptr;
    unicode_text_format.dwAspect = DVASPECT_CONTENT;
    unicode_text_format.lindex = -1;
    unicode_text_format.tymed = TYMED_HGLOBAL;

    STGMEDIUM text_medium = {};
    if (SUCCEEDED(data_object->GetData(&unicode_text_format, &text_medium))) {
        if (text_medium.tymed == TYMED_HGLOBAL && text_medium.hGlobal != nullptr) {
            const wchar_t* text = static_cast<const wchar_t*>(GlobalLock(text_medium.hGlobal));
            if (text != nullptr) {
                AppendExistingPathsFromTextBlock(text, source_paths);
                GlobalUnlock(text_medium.hGlobal);
            }
        }
        ReleaseStgMedium(&text_medium);
        if (!source_paths->empty()) {
            return true;
        }
    }

    return !source_paths->empty();
}

bool FileListView::ImportVirtualFilesFromDataObject(IDataObject* data_object, const std::wstring& destination_path) {
    if (data_object == nullptr || destination_path.empty()) {
        return false;
    }

    const UINT descriptor_format_w = RegisterClipboardFormatW(CFSTR_FILEDESCRIPTORW);
    if (descriptor_format_w == 0U) {
        return false;
    }

    FORMATETC descriptor_format = {};
    descriptor_format.cfFormat = static_cast<CLIPFORMAT>(descriptor_format_w);
    descriptor_format.ptd = nullptr;
    descriptor_format.dwAspect = DVASPECT_CONTENT;
    descriptor_format.lindex = -1;
    descriptor_format.tymed = TYMED_HGLOBAL;

    STGMEDIUM descriptor_medium = {};
    if (FAILED(data_object->GetData(&descriptor_format, &descriptor_medium))) {
        return false;
    }
    if (descriptor_medium.tymed != TYMED_HGLOBAL || descriptor_medium.hGlobal == nullptr) {
        ReleaseStgMedium(&descriptor_medium);
        return false;
    }

    struct VirtualDescriptorItem {
        std::wstring leaf_name;
        DWORD flags = 0;
        DWORD attributes = 0;
    };
    std::vector<VirtualDescriptorItem> items;
    std::vector<std::wstring> source_name_checks;

    const SIZE_T descriptor_bytes = GlobalSize(descriptor_medium.hGlobal);
    const FILEGROUPDESCRIPTORW* group =
        static_cast<const FILEGROUPDESCRIPTORW*>(GlobalLock(descriptor_medium.hGlobal));
    if (group == nullptr) {
        ReleaseStgMedium(&descriptor_medium);
        return false;
    }

    const size_t max_items = (descriptor_bytes > sizeof(UINT))
        ? (descriptor_bytes - sizeof(UINT)) / sizeof(FILEDESCRIPTORW)
        : 0U;
    const size_t item_count = (std::min)(static_cast<size_t>(group->cItems), max_items);
    items.reserve(item_count);
    source_name_checks.reserve(item_count);

    for (size_t i = 0; i < item_count; ++i) {
        const FILEDESCRIPTORW& descriptor = group->fgd[i];
        VirtualDescriptorItem item = {};
        item.leaf_name = ResolveDropLeafName(descriptor.cFileName, static_cast<UINT>(i));
        item.flags = descriptor.dwFlags;
        item.attributes = descriptor.dwFileAttributes;
        items.push_back(item);
        source_name_checks.push_back(item.leaf_name);
    }

    GlobalUnlock(descriptor_medium.hGlobal);
    ReleaseStgMedium(&descriptor_medium);
    if (items.empty()) {
        return false;
    }

    std::wstring conflict_name;
    if (TryFindDropNameConflict(source_name_checks, destination_path, &conflict_name)) {
        std::wstring warning_message = L"Drop blocked: an item with the same name already exists in the destination folder.";
        if (!conflict_name.empty()) {
            warning_message = L"Drop blocked: \"";
            warning_message.append(conflict_name);
            warning_message.append(L"\" already exists in the destination folder.");
        }
        MessageBoxW(hwnd_, warning_message.c_str(), L"FileExplorer", MB_OK | MB_ICONWARNING);
        return false;
    }

    const UINT contents_format = RegisterClipboardFormatW(CFSTR_FILECONTENTS);
    if (contents_format == 0U) {
        return false;
    }

    bool imported_any = false;
    for (size_t index = 0; index < items.size(); ++index) {
        const VirtualDescriptorItem& item = items[index];
        const std::wstring target_path = JoinPathForDropTarget(destination_path, item.leaf_name);

        const bool is_directory =
            (item.flags & FD_ATTRIBUTES) != 0U &&
            (item.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0U;
        if (is_directory) {
            if (!CreateDirectoryW(target_path.c_str(), nullptr)) {
                const DWORD create_error = GetLastError();
                if (create_error != ERROR_ALREADY_EXISTS) {
                    LogLastError(L"CreateDirectoryW(DropVirtualDirectory)");
                    return false;
                }
            }
            imported_any = true;
            continue;
        }

        FORMATETC contents_request = {};
        contents_request.cfFormat = static_cast<CLIPFORMAT>(contents_format);
        contents_request.ptd = nullptr;
        contents_request.dwAspect = DVASPECT_CONTENT;
        contents_request.lindex = static_cast<LONG>(index);
        contents_request.tymed = TYMED_ISTREAM | TYMED_HGLOBAL | TYMED_FILE | TYMED_ISTORAGE;

        STGMEDIUM contents_medium = {};
        const HRESULT get_contents_result = data_object->GetData(&contents_request, &contents_medium);
        if (FAILED(get_contents_result)) {
            LogHResult(L"IDataObject::GetData(CFSTR_FILECONTENTS)", get_contents_result);
            return false;
        }

        const bool write_ok = WriteStgMediumToFile(contents_medium, target_path);
        ReleaseStgMedium(&contents_medium);
        if (!write_ok) {
            return false;
        }

        imported_any = true;
    }

    if (imported_any) {
        RequestRefresh();
    }
    return imported_any;
}

bool FileListView::HandleOleDataObjectDrop(
    IDataObject* data_object,
    DWORD key_state,
    const POINT* client_point,
    DWORD* performed_effect) {
    if (performed_effect != nullptr) {
        *performed_effect = DROPEFFECT_NONE;
    }
    if (data_object == nullptr || search_mode_ || current_path_.empty()) {
        return false;
    }

    std::wstring destination_path = current_path_;
    if (client_point != nullptr && hwnd_ != nullptr) {
        LVHITTESTINFO hit = {};
        hit.pt = *client_point;
        const int hit_index = ListView_SubItemHitTest(hwnd_, &hit);
        if (IsSelectableIndex(hit_index) && display_entries_[hit_index].is_folder) {
            destination_path = BuildFullPath(display_entries_[hit_index]);
        }
    }

    std::vector<std::wstring> source_paths;
    if (ExtractSourcePathsFromDataObject(data_object, &source_paths) && !source_paths.empty()) {
        const bool shift_down = (key_state & MK_SHIFT) != 0U;
        const bool ctrl_down = (key_state & MK_CONTROL) != 0U;
        const bool move_drop = shift_down && !ctrl_down;
        const bool imported = HandleExternalPathsDrop(source_paths, move_drop, client_point);
        if (imported && performed_effect != nullptr) {
            *performed_effect = move_drop ? DROPEFFECT_MOVE : DROPEFFECT_COPY;
        }
        return imported;
    }

    const bool imported_virtual = ImportVirtualFilesFromDataObject(data_object, destination_path);
    if (imported_virtual && performed_effect != nullptr) {
        *performed_effect = DROPEFFECT_COPY;
    }
    return imported_virtual;
}

void FileListView::ShowContextMenu(POINT screen_point, int hit_index) {
    EndInlineRenameControl();
    ClearShellContextMenu();

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
    AppendMenuW(menu, MF_STRING, kMenuNewFolder, L"New Folder");
    AppendMenuW(menu, MF_STRING, kMenuProperties, L"Properties");
    AppendMenuW(menu, MF_STRING, kMenuCopyPath, L"Copy Path");
    if (search_mode_) {
        AppendMenuW(menu, MF_STRING, kMenuOpenFileLocation, L"Open file location");
    }
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING, kMenuMakeFavourite, L"Make Favourite");
    AppendMenuW(menu, MF_STRING, kMenuMakeFlyoutFavourite, L"Make Flyout Favourite");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(
        menu,
        MF_STRING | (show_hidden_files_ ? MF_CHECKED : MF_UNCHECKED),
        kMenuToggleShowHidden,
        L"Show hidden files");
    AppendMenuW(
        menu,
        MF_STRING | (show_extensions_ ? MF_CHECKED : MF_UNCHECKED),
        kMenuToggleShowExtensions,
        L"Show file extensions");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING, kMenuIncreaseDetailScale, L"Increase detail font + row spacing (+5%)");
    AppendMenuW(menu, MF_STRING, kMenuResetDetailScale, L"Default detail font + row spacing");

    const int target_index = ResolveContextTargetIndex(hit_index);
    const bool has_target = IsSelectableIndex(target_index);
    const bool can_create_folder = !search_mode_ && !current_path_.empty();
    const bool can_paste = FileOps::ClipboardHasDropFiles();

    if (has_target && (ListView_GetItemState(hwnd_, target_index, LVIS_SELECTED) & LVIS_SELECTED) == 0) {
        EnsureSingleSelectionAtIndex(target_index);
    }

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

    if (!can_create_folder) {
        EnableMenuItem(menu, kMenuNewFolder, MF_BYCOMMAND | MF_GRAYED);
    }

    if (!can_paste) {
        EnableMenuItem(menu, kMenuPaste, MF_BYCOMMAND | MF_GRAYED);
    }

    if (detail_scale_percent_ >= kDetailScaleMaxPercent) {
        EnableMenuItem(menu, kMenuIncreaseDetailScale, MF_BYCOMMAND | MF_GRAYED);
    }
    if (detail_scale_percent_ <= kDefaultDetailScalePercent) {
        EnableMenuItem(menu, kMenuResetDetailScale, MF_BYCOMMAND | MF_GRAYED);
    }

    if (has_target && !display_entries_[target_index].is_folder) {
        const UINT disable_flags = MF_BYCOMMAND | MF_GRAYED;
        EnableMenuItem(menu, kMenuMakeFavourite, disable_flags);
        EnableMenuItem(menu, kMenuMakeFlyoutFavourite, disable_flags);
    }

    if (has_target) {
        const FileEntry& entry = display_entries_[target_index];
        (void)AppendSevenZipContextMenu(menu, BuildFullPath(entry));
    }

    const UINT command = static_cast<UINT>(TrackPopupMenuEx(
        menu,
        TPM_RIGHTBUTTON | TPM_RETURNCMD,
        screen_point.x,
        screen_point.y,
        hwnd_,
        nullptr));

    if (command != 0U) {
        if (has_target) {
            const FileEntry& entry = display_entries_[target_index];
            const std::wstring full_path = BuildFullPath(entry);

            switch (command) {
            case kMenuCopy:
                CopySelectionToClipboard(false, target_index);
                break;

            case kMenuCut:
                CopySelectionToClipboard(true, target_index);
                break;

            case kMenuDelete:
                DeleteSelection(false, target_index);
                break;

            case kMenuCopyPath: {
                std::vector<int> indices = CollectSelectedSelectableIndices();
                if (indices.empty()) {
                    indices.push_back(target_index);
                }
                const std::vector<std::wstring> paths = CollectPathsForIndices(indices);
                std::wstring text;
                for (size_t i = 0; i < paths.size(); ++i) {
                    if (i > 0) {
                        text.append(L"\r\n");
                    }
                    text.append(paths[i]);
                }
                FileOps::CopyTextToClipboard(hwnd_, text);
                break;
            }

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
                FileOps::ShowProperties(hwnd_, full_path);
                break;

            case kMenuRename:
                BeginInlineRename(target_index);
                break;

            default:
                (void)InvokeShellContextMenu(command);
                break;
            }
        }

        switch (command) {
        case kMenuPaste:
            PasteFromClipboard();
            break;

        case kMenuNewFolder:
            CreateNewFolderAndBeginRename();
            break;

        case kMenuToggleShowHidden:
            if (parent_hwnd_ != nullptr &&
                !PostMessageW(parent_hwnd_, WM_FE_FILELIST_TOGGLE_SHOW_HIDDEN, 0, 0)) {
                LogLastError(L"PostMessageW(WM_FE_FILELIST_TOGGLE_SHOW_HIDDEN)");
            }
            break;

        case kMenuToggleShowExtensions:
            if (parent_hwnd_ != nullptr &&
                !PostMessageW(parent_hwnd_, WM_FE_FILELIST_TOGGLE_SHOW_EXTENSIONS, 0, 0)) {
                LogLastError(L"PostMessageW(WM_FE_FILELIST_TOGGLE_SHOW_EXTENSIONS)");
            }
            break;

        case kMenuIncreaseDetailScale:
            if (parent_hwnd_ != nullptr &&
                !PostMessageW(parent_hwnd_, WM_FE_FILELIST_INCREASE_DETAIL_SCALE, 0, 0)) {
                LogLastError(L"PostMessageW(WM_FE_FILELIST_INCREASE_DETAIL_SCALE)");
            }
            break;

        case kMenuResetDetailScale:
            if (parent_hwnd_ != nullptr &&
                !PostMessageW(parent_hwnd_, WM_FE_FILELIST_RESET_DETAIL_SCALE, 0, 0)) {
                LogLastError(L"PostMessageW(WM_FE_FILELIST_RESET_DETAIL_SCALE)");
            }
            break;

        default:
            break;
        }
    }

    DestroyMenu(menu);
    ClearShellContextMenu();
    if (rename_edit_hwnd_ == nullptr && hwnd_ != nullptr) {
        SetFocus(hwnd_);
    }
}

int FileListView::ResolveContextTargetIndex(int hit_index) const {
    if (IsSelectableIndex(hit_index)) {
        return hit_index;
    }

    const int focused = FocusedIndex();
    if (IsSelectableIndex(focused)) {
        return focused;
    }

    if (hwnd_ != nullptr) {
        int selected = -1;
        while (true) {
            selected = ListView_GetNextItem(hwnd_, selected, LVNI_SELECTED);
            if (selected < 0) {
                break;
            }
            if (IsSelectableIndex(selected)) {
                return selected;
            }
        }
    }

    return FindNextSelectableIndex(-1, 1);
}

std::vector<int> FileListView::CollectSelectedSelectableIndices() const {
    std::vector<int> indices;
    if (hwnd_ == nullptr) {
        return indices;
    }

    int index = -1;
    while (true) {
        index = ListView_GetNextItem(hwnd_, index, LVNI_SELECTED);
        if (index < 0) {
            break;
        }
        if (IsSelectableIndex(index)) {
            indices.push_back(index);
        }
    }
    return indices;
}

std::vector<std::wstring> FileListView::CollectPathsForIndices(const std::vector<int>& indices) const {
    std::vector<std::wstring> paths;
    paths.reserve(indices.size());
    for (int index : indices) {
        if (!IsSelectableIndex(index)) {
            continue;
        }
        paths.push_back(BuildFullPath(display_entries_[index]));
    }
    return paths;
}

bool FileListView::StartExternalDrag(const std::vector<std::wstring>& paths) {
    if (paths.empty()) {
        return false;
    }

    const bool shift_down = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
    const bool ctrl_down = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const DWORD preferred_effect =
        (shift_down && !ctrl_down) ? DROPEFFECT_MOVE : DROPEFFECT_COPY;
    constexpr DWORD allowed_effects = DROPEFFECT_COPY | DROPEFFECT_MOVE;

    HGLOBAL drop_files_handle = CreateDropFilesHandle(paths);
    if (drop_files_handle == nullptr) {
        return false;
    }

    HGLOBAL drop_effect_handle = CreateDropEffectHandle(preferred_effect);
    auto* data_object = new (std::nothrow) ExternalDragDataObject(drop_files_handle, drop_effect_handle);
    if (data_object == nullptr) {
        GlobalFree(drop_files_handle);
        if (drop_effect_handle != nullptr) {
            GlobalFree(drop_effect_handle);
        }
        return false;
    }

    auto* drop_source = new (std::nothrow) ExternalDragDropSource();
    if (drop_source == nullptr) {
        data_object->Release();
        return false;
    }

    DWORD performed_effect = DROPEFFECT_NONE;
    const HRESULT drag_result = DoDragDrop(
        static_cast<IDataObject*>(data_object),
        static_cast<IDropSource*>(drop_source),
        allowed_effects,
        &performed_effect);
    (void)performed_effect;
    if (FAILED(drag_result)) {
        wchar_t buffer[256] = {};
        swprintf_s(
            buffer,
            L"[FileExplorer] DoDragDrop failed (HRESULT=0x%08lX).\r\n",
            static_cast<unsigned long>(drag_result));
        OutputDebugStringW(buffer);
    }

    drop_source->Release();
    data_object->Release();

    return drag_result == DRAGDROP_S_DROP || drag_result == DRAGDROP_S_CANCEL;
}

void FileListView::EnsureSingleSelectionAtIndex(int index) {
    if (!IsSelectableIndex(index) || hwnd_ == nullptr) {
        return;
    }
    ClearSelection();
    ListView_SetItemState(
        hwnd_,
        index,
        LVIS_SELECTED | LVIS_FOCUSED,
        LVIS_SELECTED | LVIS_FOCUSED);
    ListView_SetSelectionMark(hwnd_, index);
    ListView_EnsureVisible(hwnd_, index, FALSE);
    PostStatusUpdate();
}

bool FileListView::CopySelectionToClipboard(bool cut, int fallback_index) {
    std::vector<int> indices = CollectSelectedSelectableIndices();
    if (indices.empty() && IsSelectableIndex(fallback_index)) {
        EnsureSingleSelectionAtIndex(fallback_index);
        indices.push_back(fallback_index);
    }
    if (indices.empty()) {
        return false;
    }

    const std::vector<std::wstring> paths = CollectPathsForIndices(indices);
    if (!FileOps::CopyPathsToClipboard(hwnd_, paths, cut)) {
        return false;
    }

    if (cut) {
        SetCutPendingPaths(paths);
    } else {
        ClearCutPendingState();
    }
    return true;
}

bool FileListView::PasteFromClipboard() {
    if (current_path_.empty()) {
        return false;
    }

    bool was_move = false;
    if (!FileOps::PasteFromClipboard(hwnd_, current_path_, &was_move)) {
        return false;
    }

    if (was_move) {
        ClearCutPendingState();
    }

    if (!search_mode_) {
        RequestRefresh();
    }
    PostStatusUpdate();
    if (hwnd_ != nullptr) {
        SetFocus(hwnd_);
    }
    return true;
}

bool FileListView::DeleteSelection(bool permanent, int fallback_index) {
    std::vector<int> indices = CollectSelectedSelectableIndices();
    if (indices.empty() && IsSelectableIndex(fallback_index)) {
        EnsureSingleSelectionAtIndex(fallback_index);
        indices.push_back(fallback_index);
    }
    if (indices.empty()) {
        return false;
    }
    const int top_deleted_index = indices.front();

    const std::vector<std::wstring> paths = CollectPathsForIndices(indices);
    if (paths.empty()) {
        return false;
    }

    if (permanent) {
        wchar_t prompt[256] = {};
        swprintf_s(
            prompt,
            L"Permanently delete %zu selected item(s)?",
            static_cast<size_t>(paths.size()));
        if (MessageBoxW(hwnd_, prompt, L"FileExplorer", MB_ICONWARNING | MB_YESNO | MB_DEFBUTTON2) != IDYES) {
            return false;
        }
    }

    const bool deleted = permanent
        ? FileOps::DeletePermanently(hwnd_, paths)
        : FileOps::DeleteToRecycleBin(hwnd_, paths);
    if (!deleted) {
        return false;
    }

    RemoveCutPendingPaths(paths);

    std::unordered_set<std::wstring> removed_paths;
    removed_paths.reserve(paths.size());
    for (const std::wstring& path : paths) {
        removed_paths.insert(ToLowerCopy(NormalizePath(path)));
    }

    std::vector<FileEntry> filtered_entries;
    filtered_entries.reserve(base_entries_.size());
    for (const FileEntry& entry : base_entries_) {
        const std::wstring normalized_entry = ToLowerCopy(NormalizePath(BuildFullPath(entry)));
        if (removed_paths.find(normalized_entry) == removed_paths.end()) {
            filtered_entries.push_back(entry);
        }
    }
    base_entries_ = std::move(filtered_entries);

    SortBaseEntries();
    BuildDisplayEntriesWithGroups();
    UpdateSortIndicators();
    UpdateItemCountAndRefresh(false);

    int focus_index = -1;
    if (!display_entries_.empty()) {
        int anchor_index = top_deleted_index - 1;
        if (anchor_index >= static_cast<int>(display_entries_.size())) {
            anchor_index = static_cast<int>(display_entries_.size()) - 1;
        }

        focus_index = FindNextSelectableIndex(anchor_index + 1, -1);
        if (focus_index < 0) {
            const int downward_start = (anchor_index < 0) ? -1 : (anchor_index - 1);
            focus_index = FindNextSelectableIndex(downward_start, 1);
        }
    }

    if (focus_index >= 0) {
        EnsureSingleSelectionAtIndex(focus_index);
    } else {
        ClearSelection();
    }

    PostStatusUpdate();

    if (!search_mode_) {
        RequestRefresh();
    }
    if (hwnd_ != nullptr) {
        SetFocus(hwnd_);
    }
    return true;
}

bool FileListView::BeginInlineRename(int index) {
    if (!IsSelectableIndex(index) || hwnd_ == nullptr) {
        return false;
    }

    EndInlineRenameControl();

    RECT label_rect = {};
    if (!ListView_GetSubItemRect(hwnd_, index, 0, LVIR_LABEL, &label_rect)) {
        if (!ListView_GetItemRect(hwnd_, index, &label_rect, LVIR_LABEL)) {
            return false;
        }
    }

    const int min_height = MulDiv(24, static_cast<int>(dpi_), 96);
    const int rect_height = static_cast<int>(label_rect.bottom - label_rect.top);
    const int rect_width = static_cast<int>(label_rect.right - label_rect.left);
    const int height = (std::max)(min_height, rect_height);
    const int width = (std::max)(MulDiv(120, static_cast<int>(dpi_), 96), rect_width);

    const FileEntry& entry = display_entries_[index];
    rename_item_index_ = index;
    rename_original_name_ = entry.name;
    rename_original_full_path_ = BuildFullPath(entry);

    rename_edit_hwnd_ = CreateWindowExW(
        WS_EX_CLIENTEDGE,
        WC_EDITW,
        entry.name.c_str(),
        WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
        label_rect.left,
        label_rect.top,
        width,
        height,
        hwnd_,
        nullptr,
        instance_,
        nullptr);
    if (rename_edit_hwnd_ == nullptr) {
        LogLastError(L"CreateWindowExW(FileList RenameEdit)");
        rename_item_index_ = -1;
        rename_original_name_.clear();
        rename_original_full_path_.clear();
        return false;
    }

    if (!SetWindowSubclass(
            rename_edit_hwnd_,
            &FileListView::RenameEditSubclassProc,
            kRenameEditSubclassId,
            reinterpret_cast<DWORD_PTR>(this))) {
        LogLastError(L"SetWindowSubclass(FileList RenameEdit)");
        DestroyWindow(rename_edit_hwnd_);
        rename_edit_hwnd_ = nullptr;
        rename_item_index_ = -1;
        rename_original_name_.clear();
        rename_original_full_path_.clear();
        return false;
    }

    SendMessageW(rename_edit_hwnd_, WM_SETFONT, SendMessageW(hwnd_, WM_GETFONT, 0, 0), TRUE);
    SetFocus(rename_edit_hwnd_);

    int selection_end = static_cast<int>(entry.name.size());
    if (!entry.is_folder) {
        const std::wstring::size_type dot = entry.name.find_last_of(L'.');
        if (dot != std::wstring::npos && dot > 0) {
            selection_end = static_cast<int>(dot);
        }
    }
    SendMessageW(rename_edit_hwnd_, EM_SETSEL, 0, selection_end);
    return true;
}

bool FileListView::CommitInlineRename() {
    if (rename_edit_hwnd_ == nullptr || !IsWindow(rename_edit_hwnd_)) {
        return false;
    }

    const int length = GetWindowTextLengthW(rename_edit_hwnd_);
    std::wstring new_name((std::max)(0, length) + 1, L'\0');
    if (length > 0) {
        GetWindowTextW(rename_edit_hwnd_, new_name.data(), length + 1);
        new_name.resize(static_cast<size_t>(length));
    } else {
        new_name.clear();
    }

    const std::wstring::size_type first = new_name.find_first_not_of(L" \t");
    if (first == std::wstring::npos) {
        MessageBoxW(rename_edit_hwnd_, L"Name cannot be empty.", L"Rename", MB_ICONWARNING | MB_OK);
        SetFocus(rename_edit_hwnd_);
        return false;
    }
    const std::wstring::size_type last = new_name.find_last_not_of(L" \t");
    new_name = new_name.substr(first, last - first + 1);

    if (new_name.find_first_of(L"\\/:*?\"<>|") != std::wstring::npos) {
        MessageBoxW(
            rename_edit_hwnd_,
            L"Name contains invalid characters.",
            L"Rename",
            MB_ICONWARNING | MB_OK);
        SetFocus(rename_edit_hwnd_);
        return false;
    }

    if (new_name == rename_original_name_) {
        EndInlineRenameControl();
        return true;
    }

    if (!ApplyInlineRenameChange(rename_original_full_path_, new_name)) {
        SetFocus(rename_edit_hwnd_);
        return false;
    }

    EndInlineRenameControl();
    return true;
}

void FileListView::CancelInlineRename() {
    EndInlineRenameControl();
}

void FileListView::EndInlineRenameControl() {
    const bool had_rename_control = rename_edit_hwnd_ != nullptr;
    if (rename_edit_hwnd_ != nullptr) {
        RemoveWindowSubclass(rename_edit_hwnd_, &FileListView::RenameEditSubclassProc, kRenameEditSubclassId);
        if (IsWindow(rename_edit_hwnd_)) {
            DestroyWindow(rename_edit_hwnd_);
        }
        rename_edit_hwnd_ = nullptr;
    }

    rename_item_index_ = -1;
    rename_original_name_.clear();
    rename_original_full_path_.clear();

    if (had_rename_control && hwnd_ != nullptr) {
        SetFocus(hwnd_);
    }
}

bool FileListView::ApplyInlineRenameChange(const std::wstring& old_full_path, const std::wstring& new_name) {
    const std::wstring parent_folder = ParentPath(old_full_path);
    if (parent_folder.empty()) {
        return false;
    }

    std::wstring new_full_path = parent_folder;
    if (!new_full_path.empty() && new_full_path.back() != L'\\') {
        new_full_path.push_back(L'\\');
    }
    new_full_path.append(new_name);

    const std::wstring normalized_old = NormalizePath(old_full_path);
    const std::wstring normalized_new = NormalizePath(new_full_path);

    if (CompareStringOrdinal(
            normalized_old.c_str(),
            static_cast<int>(normalized_old.size()),
            normalized_new.c_str(),
            static_cast<int>(normalized_new.size()),
            FALSE) == CSTR_EQUAL) {
        return true;
    }

    DWORD rename_error = ERROR_SUCCESS;
    if (!FileOps::RenamePath(normalized_old, normalized_new, &rename_error)) {
        std::wstring message = L"Rename failed.\n\n";
        if (rename_error == ERROR_SHARING_VIOLATION || rename_error == ERROR_LOCK_VIOLATION) {
            message.append(L"The file is currently open in another program.\nClose it and try again.");
        } else {
            const std::wstring error_text = FormatWin32ErrorMessage(rename_error);
            message.append(error_text);
        }
        MessageBoxW(rename_edit_hwnd_, message.c_str(), L"Rename", MB_ICONERROR | MB_OK);
        return false;
    }

    if (IsPathCutPending(normalized_old)) {
        RemoveCutPendingPaths({normalized_old});
        SetCutPendingPaths({normalized_new});
    }

    bool updated = false;
    for (FileEntry& entry : base_entries_) {
        const std::wstring entry_path = NormalizePath(BuildFullPath(entry));
        if (CompareStringOrdinal(
                entry_path.c_str(),
                static_cast<int>(entry_path.size()),
                normalized_old.c_str(),
                static_cast<int>(normalized_old.size()),
                TRUE) != CSTR_EQUAL) {
            continue;
        }

        entry.name = new_name;
        entry.extension = entry.is_folder ? L"" : GetExtensionFromName(new_name);
        entry.full_path = normalized_new;
        updated = true;
        break;
    }

    if (!updated && IsSelectableIndex(rename_item_index_)) {
        FileEntry& entry = display_entries_[rename_item_index_];
        entry.name = new_name;
        entry.extension = entry.is_folder ? L"" : GetExtensionFromName(new_name);
        entry.full_path = normalized_new;
    }

    SortBaseEntries();
    BuildDisplayEntriesWithGroups();
    UpdateSortIndicators();
    UpdateItemCountAndRefresh(false);

    int renamed_index = -1;
    for (int i = 0; i < static_cast<int>(display_entries_.size()); ++i) {
        if (!IsSelectableIndex(i)) {
            continue;
        }
        const std::wstring entry_path = NormalizePath(BuildFullPath(display_entries_[i]));
        if (CompareStringOrdinal(
                entry_path.c_str(),
                static_cast<int>(entry_path.size()),
                normalized_new.c_str(),
                static_cast<int>(normalized_new.size()),
                TRUE) == CSTR_EQUAL) {
            renamed_index = i;
            break;
        }
    }

    if (renamed_index >= 0) {
        EnsureSingleSelectionAtIndex(renamed_index);
    } else {
        PostStatusUpdate();
    }

    if (!search_mode_) {
        RequestRefresh();
    }
    return true;
}

bool FileListView::CreateNewFolderAndBeginRename() {
    if (search_mode_ || current_path_.empty()) {
        return false;
    }

    const std::wstring folder_path = BuildUniqueNewFolderPath();
    if (folder_path.empty()) {
        return false;
    }

    DWORD create_error = ERROR_SUCCESS;
    if (!FileOps::CreateDirectoryPath(folder_path, &create_error)) {
        const std::wstring error_text = FormatWin32ErrorMessage(create_error);
        std::wstring message = L"Failed to create folder.\n\n";
        message.append(error_text);
        MessageBoxW(hwnd_, message.c_str(), L"New Folder", MB_ICONERROR | MB_OK);
        return false;
    }

    WIN32_FILE_ATTRIBUTE_DATA attributes = {};
    FILETIME modified_time = {};
    if (GetFileAttributesExW(folder_path.c_str(), GetFileExInfoStandard, &attributes)) {
        modified_time = attributes.ftLastWriteTime;
    } else {
        GetSystemTimeAsFileTime(&modified_time);
    }

    FileEntry entry = {};
    entry.name = PathFindFileNameW(folder_path.c_str());
    entry.full_path = folder_path;
    entry.modified_time = modified_time;
    entry.size_bytes = 0;
    entry.attributes = FILE_ATTRIBUTE_DIRECTORY;
    entry.is_folder = true;
    entry.icon_index = ResolveIconIndex(folder_path, true);

    base_entries_.push_back(std::move(entry));
    SortBaseEntries();
    BuildDisplayEntriesWithGroups();
    UpdateSortIndicators();
    UpdateItemCountAndRefresh(false);

    int new_index = -1;
    for (int i = 0; i < static_cast<int>(display_entries_.size()); ++i) {
        if (!IsSelectableIndex(i)) {
            continue;
        }
        const std::wstring entry_path = NormalizePath(BuildFullPath(display_entries_[i]));
        if (CompareStringOrdinal(
                entry_path.c_str(),
                static_cast<int>(entry_path.size()),
                folder_path.c_str(),
                static_cast<int>(folder_path.size()),
                TRUE) == CSTR_EQUAL) {
            new_index = i;
            break;
        }
    }

    if (new_index >= 0) {
        EnsureSingleSelectionAtIndex(new_index);
        BeginInlineRename(new_index);
    } else {
        PostStatusUpdate();
    }

    if (!search_mode_) {
        RequestRefresh();
    }
    return true;
}

std::wstring FileListView::BuildUniqueNewFolderPath() const {
    if (current_path_.empty()) {
        return L"";
    }

    constexpr wchar_t kBaseName[] = L"New folder";
    for (int suffix = 1; suffix <= 9999; ++suffix) {
        std::wstring candidate = current_path_;
        if (candidate.back() != L'\\') {
            candidate.push_back(L'\\');
        }
        candidate.append(kBaseName);
        if (suffix > 1) {
            wchar_t suffix_text[32] = {};
            swprintf_s(suffix_text, L" (%d)", suffix);
            candidate.append(suffix_text);
        }

        const DWORD attributes = GetFileAttributesW(candidate.c_str());
        if (attributes == INVALID_FILE_ATTRIBUTES) {
            return candidate;
        }
    }

    return L"";
}

void FileListView::CaptureCurrentColumnWidths() {
    if (hwnd_ == nullptr) {
        return;
    }

    for (int i = 0; i < kColumnCount; ++i) {
        if ((i == 4 && !search_mode_) || (i == 1 && !show_extensions_)) {
            continue;
        }

        const int width_px = ListView_GetColumnWidth(hwnd_, i);
        if (width_px <= 0) {
            continue;
        }

        const int logical_width = MulDiv(width_px, 96, static_cast<int>(dpi_));
        column_widths_logical_[static_cast<size_t>(i)] = ClampColumnWidthLogical(i, logical_width);
    }
}

void FileListView::RequestRefresh() {
    if (parent_hwnd_ != nullptr) {
        if (!PostMessageW(parent_hwnd_, WM_FE_FILELIST_REFRESH, 0, 0)) {
            LogLastError(L"PostMessageW(WM_FE_FILELIST_REFRESH)");
        }
    }
}

void FileListView::ClearCutPendingState() {
    if (cut_pending_paths_.empty()) {
        return;
    }
    cut_pending_paths_.clear();
    if (hwnd_ != nullptr) {
        InvalidateRect(hwnd_, nullptr, FALSE);
    }
}

void FileListView::SetCutPendingPaths(const std::vector<std::wstring>& paths) {
    cut_pending_paths_.clear();
    cut_pending_paths_.reserve(paths.size());
    for (const std::wstring& path : paths) {
        if (!path.empty()) {
            cut_pending_paths_.insert(ToLowerCopy(NormalizePath(path)));
        }
    }
    if (hwnd_ != nullptr) {
        InvalidateRect(hwnd_, nullptr, FALSE);
    }
}

void FileListView::RemoveCutPendingPaths(const std::vector<std::wstring>& paths) {
    if (cut_pending_paths_.empty() || paths.empty()) {
        return;
    }
    for (const std::wstring& path : paths) {
        if (!path.empty()) {
            cut_pending_paths_.erase(ToLowerCopy(NormalizePath(path)));
        }
    }
    if (hwnd_ != nullptr) {
        InvalidateRect(hwnd_, nullptr, FALSE);
    }
}

bool FileListView::IsPathCutPending(const std::wstring& full_path) const {
    if (full_path.empty() || cut_pending_paths_.empty()) {
        return false;
    }
    const std::wstring key = ToLowerCopy(NormalizePath(full_path));
    return cut_pending_paths_.find(key) != cut_pending_paths_.end();
}

bool FileListView::AppendSevenZipContextMenu(HMENU menu, const std::wstring& path) {
    if (menu == nullptr || path.empty()) {
        return false;
    }

    HMENU shell_menu = CreatePopupMenu();
    if (shell_menu == nullptr) {
        LogLastError(L"CreatePopupMenu(AppendSevenZipContextMenu)");
        return false;
    }

    if (!AppendShellContextMenu(shell_menu, path, kShellMenuCommandFirst, kShellMenuCommandLast)) {
        DestroyMenu(shell_menu);
        return false;
    }

    const int shell_item_count = GetMenuItemCount(shell_menu);
    if (shell_item_count <= 0) {
        DestroyMenu(shell_menu);
        ClearShellContextMenu();
        return false;
    }

    int seven_zip_index = -1;
    HMENU seven_zip_submenu = nullptr;
    UINT seven_zip_command_id = 0;
    std::wstring seven_zip_label;

    for (int index = 0; index < shell_item_count; ++index) {
        std::wstring menu_label;
        if (!TryGetMenuItemTextByPosition(shell_menu, static_cast<UINT>(index), &menu_label)) {
            continue;
        }
        if (!IsSevenZipMenuLabel(menu_label)) {
            continue;
        }

        MENUITEMINFOW item_info = {};
        item_info.cbSize = sizeof(item_info);
        item_info.fMask = MIIM_ID | MIIM_SUBMENU;
        if (!GetMenuItemInfoW(shell_menu, static_cast<UINT>(index), TRUE, &item_info)) {
            continue;
        }

        seven_zip_index = index;
        seven_zip_submenu = item_info.hSubMenu;
        seven_zip_command_id = item_info.wID;
        seven_zip_label = std::move(menu_label);
        break;
    }

    if (seven_zip_index < 0) {
        DestroyMenu(shell_menu);
        ClearShellContextMenu();
        return false;
    }

    if (!RemoveMenu(shell_menu, static_cast<UINT>(seven_zip_index), MF_BYPOSITION)) {
        LogLastError(L"RemoveMenu(AppendSevenZipContextMenu)");
        DestroyMenu(shell_menu);
        ClearShellContextMenu();
        return false;
    }

    if (GetMenuItemCount(menu) > 0) {
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    }

    const std::wstring visible_label = seven_zip_label.empty() ? std::wstring(L"7-Zip") : seven_zip_label;
    bool appended = false;
    if (seven_zip_submenu != nullptr) {
        appended = AppendMenuW(
            menu,
            MF_POPUP | MF_STRING,
            reinterpret_cast<UINT_PTR>(seven_zip_submenu),
            visible_label.c_str()) != FALSE;
    } else if (seven_zip_command_id != 0U) {
        appended = AppendMenuW(
            menu,
            MF_STRING,
            static_cast<UINT_PTR>(seven_zip_command_id),
            visible_label.c_str()) != FALSE;
    }

    DestroyMenu(shell_menu);
    if (!appended) {
        if (seven_zip_submenu != nullptr) {
            DestroyMenu(seven_zip_submenu);
        }
        ClearShellContextMenu();
        return false;
    }
    return true;
}

bool FileListView::AppendShellContextMenu(HMENU menu, const std::wstring& path, UINT first_id, UINT last_id) {
    if (menu == nullptr || path.empty()) {
        return false;
    }

    PIDLIST_ABSOLUTE absolute_item_id = nullptr;
    IShellFolder* parent_folder = nullptr;
    IContextMenu* context_menu = nullptr;
    PCUITEMID_CHILD child_item_id = nullptr;

    const HRESULT parse_result = SHParseDisplayName(path.c_str(), nullptr, &absolute_item_id, 0, nullptr);
    if (FAILED(parse_result) || absolute_item_id == nullptr) {
        return false;
    }

    const HRESULT bind_result =
        SHBindToParent(absolute_item_id, IID_IShellFolder, reinterpret_cast<void**>(&parent_folder), &child_item_id);
    if (FAILED(bind_result) || parent_folder == nullptr || child_item_id == nullptr) {
        CoTaskMemFree(absolute_item_id);
        return false;
    }

    const HRESULT context_result = parent_folder->GetUIObjectOf(
        hwnd_,
        1,
        &child_item_id,
        IID_IContextMenu,
        nullptr,
        reinterpret_cast<void**>(&context_menu));
    parent_folder->Release();
    CoTaskMemFree(absolute_item_id);
    if (FAILED(context_result) || context_menu == nullptr) {
        return false;
    }

    const UINT insert_index = static_cast<UINT>(GetMenuItemCount(menu));
    const HRESULT query_result =
        context_menu->QueryContextMenu(menu, insert_index, first_id, last_id, CMF_NORMAL);
    const UINT added_items = static_cast<UINT>(HRESULT_CODE(query_result));
    if (FAILED(query_result) || added_items == 0U) {
        context_menu->Release();
        return false;
    }

    shell_context_menu_ = context_menu;
    shell_context_menu_first_id_ = first_id;
    shell_context_menu_last_id_ = last_id;
    return true;
}

bool FileListView::InvokeShellContextMenu(UINT command_id) {
    if (shell_context_menu_ == nullptr) {
        return false;
    }
    if (command_id < shell_context_menu_first_id_ || command_id > shell_context_menu_last_id_) {
        return false;
    }

    CMINVOKECOMMANDINFOEX invoke = {};
    invoke.cbSize = sizeof(invoke);
    invoke.fMask = CMIC_MASK_UNICODE;
    invoke.hwnd = hwnd_;
    invoke.lpVerb = MAKEINTRESOURCEA(command_id - shell_context_menu_first_id_);
    invoke.lpVerbW = MAKEINTRESOURCEW(command_id - shell_context_menu_first_id_);
    invoke.nShow = SW_SHOWNORMAL;

    const HRESULT invoke_result = shell_context_menu_->InvokeCommand(
        reinterpret_cast<LPCMINVOKECOMMANDINFO>(&invoke));
    if (FAILED(invoke_result)) {
        LogHResult(L"IContextMenu::InvokeCommand", invoke_result);
        return false;
    }
    return true;
}

void FileListView::ClearShellContextMenu() {
    if (shell_context_menu_ != nullptr) {
        shell_context_menu_->Release();
        shell_context_menu_ = nullptr;
    }
    shell_context_menu_first_id_ = 0;
    shell_context_menu_last_id_ = 0;
}

int FileListView::ResolveScaledSmallIconIndex(int system_icon_index) {
    if (system_icon_index < 0) {
        return -1;
    }

    if (detail_row_height_image_list_ == nullptr || detail_scale_percent_ <= kDefaultDetailScalePercent) {
        return system_icon_index;
    }
    if (system_image_list_ == nullptr) {
        return system_icon_index;
    }

    const auto it = detail_row_height_icon_index_map_.find(system_icon_index);
    if (it != detail_row_height_icon_index_map_.end()) {
        return it->second;
    }

    int target_width = 16;
    int target_height = (detail_row_height_image_height_ > 0) ? detail_row_height_image_height_ : 16;
    if (!ImageList_GetIconSize(detail_row_height_image_list_.get(), &target_width, &target_height)) {
        target_width = 16;
        target_height = (detail_row_height_image_height_ > 0) ? detail_row_height_image_height_ : 16;
    }

    int source_width = 16;
    int source_height = 16;
    if (!ImageList_GetIconSize(system_image_list_, &source_width, &source_height)) {
        source_width = 16;
        source_height = 16;
    }

    HBITMAP composed_bitmap = CreateTransparentDibBitmap(target_width, target_height);
    if (composed_bitmap == nullptr) {
        LogLastError(L"CreateTransparentDibBitmap(ScaledSmallIconList)");
        return -1;
    }

    HDC memory_dc = CreateCompatibleDC(nullptr);
    if (memory_dc == nullptr) {
        LogLastError(L"CreateCompatibleDC(ScaledSmallIconList)");
        DeleteObject(composed_bitmap);
        return -1;
    }

    HGDIOBJ old_bitmap = SelectObject(memory_dc, composed_bitmap);
    if (old_bitmap == HGDI_ERROR) {
        LogLastError(L"SelectObject(ScaledSmallIconList)");
        DeleteDC(memory_dc);
        DeleteObject(composed_bitmap);
        return -1;
    }

    HICON icon = ImageList_GetIcon(system_image_list_, system_icon_index, ILD_NORMAL);
    if (icon == nullptr) {
        SelectObject(memory_dc, old_bitmap);
        DeleteDC(memory_dc);
        DeleteObject(composed_bitmap);
        return -1;
    }

    const int draw_x = (std::max)(0, (target_width - source_width) / 2);
    const int draw_y = (std::max)(0, (target_height - source_height) / 2);
    if (!DrawIconEx(
            memory_dc,
            draw_x,
            draw_y,
            icon,
            source_width,
            source_height,
            0,
            nullptr,
            DI_NORMAL)) {
        LogLastError(L"DrawIconEx(ScaledSmallIconList)");
    }
    DestroyIcon(icon);

    SelectObject(memory_dc, old_bitmap);
    DeleteDC(memory_dc);

    const int scaled_index = ImageList_Add(detail_row_height_image_list_.get(), composed_bitmap, nullptr);
    DeleteObject(composed_bitmap);
    if (scaled_index < 0) {
        LogLastError(L"ImageList_Add(ScaledSmallIconList)");
        return -1;
    }

    detail_row_height_icon_index_map_.insert_or_assign(system_icon_index, scaled_index);
    return scaled_index;
}

std::wstring FileListView::BuildDisplayName(const FileEntry& entry) const {
    if (entry.is_folder || show_extensions_) {
        return entry.name;
    }
    return TrimExtensionForDisplay(entry.name);
}

void FileListView::DrawEmptyState(HDC hdc) const {
    if (hdc == nullptr || hwnd_ == nullptr || !base_entries_.empty()) {
        return;
    }

    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        return;
    }

    const bool search_empty = search_mode_;
    std::wstring title_text = L"This folder is empty";
    std::wstring subtitle_text = L"Drag files here or navigate to a folder";
    if (search_empty) {
        title_text = L"No files found matching \"";
        title_text.append(search_pattern_);
        title_text.push_back(L'"');
        subtitle_text = L"Try a broader wildcard pattern";
    }

    const int client_width = static_cast<int>(client_rect.right - client_rect.left);
    const int client_height = static_cast<int>(client_rect.bottom - client_rect.top);
    const int center_x = client_rect.left + (client_width / 2);
    const int center_y = client_rect.top + (client_height / 2);

    const int icon_size = MulDiv(48, static_cast<int>(dpi_), 96);
    const int icon_y = center_y - MulDiv(62, static_cast<int>(dpi_), 96);
    const int title_top = icon_y + icon_size + MulDiv(14, static_cast<int>(dpi_), 96);
    const int subtitle_top = title_top + MulDiv(26, static_cast<int>(dpi_), 96);

    bool drew_empty_icon = false;
    HFONT icon_font = CreateMdl2Font(36, dpi_, FW_NORMAL);
    if (icon_font != nullptr) {
        HFONT old_icon_font = static_cast<HFONT>(SelectObject(hdc, icon_font));
        const int icon_half_width = MulDiv(32, static_cast<int>(dpi_), 96);
        RECT icon_rect = {};
        icon_rect.left = center_x - icon_half_width;
        icon_rect.right = center_x + icon_half_width;
        icon_rect.top = icon_y;
        icon_rect.bottom = icon_y + icon_size;
        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, kEmptyStateTextColor);
        DrawTextW(hdc, L"\uE8B7", -1, &icon_rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
        SelectObject(hdc, old_icon_font);
        DeleteObject(icon_font);
        drew_empty_icon = true;
    }

    if (!drew_empty_icon) {
        SHFILEINFOW shell_info = {};
        if (SHGetFileInfoW(
                L"C:\\",
                FILE_ATTRIBUTE_DIRECTORY,
                &shell_info,
                sizeof(shell_info),
                SHGFI_ICON | SHGFI_LARGEICON | SHGFI_USEFILEATTRIBUTES) != 0 &&
            shell_info.hIcon != nullptr) {
            const int icon_x = center_x - (icon_size / 2);
            DrawIconEx(hdc, icon_x, icon_y, shell_info.hIcon, icon_size, icon_size, 0, nullptr, DI_NORMAL);
            DestroyIcon(shell_info.hIcon);
        }
    }

    HFONT title_font = CreateEmptyStateFont(search_empty ? 16 : 18, dpi_, FW_NORMAL);
    HFONT subtitle_font = CreateEmptyStateFont(12, dpi_, FW_NORMAL);
    HFONT old_font = nullptr;

    SetBkMode(hdc, TRANSPARENT);
    if (title_font != nullptr) {
        old_font = static_cast<HFONT>(SelectObject(hdc, title_font));
    }
    SetTextColor(hdc, kEmptyStateTitleColor);

    RECT title_rect = {};
    title_rect.left = client_rect.left + MulDiv(24, static_cast<int>(dpi_), 96);
    title_rect.right = client_rect.right - MulDiv(24, static_cast<int>(dpi_), 96);
    title_rect.top = title_top;
    title_rect.bottom = title_top + MulDiv(28, static_cast<int>(dpi_), 96);
    DrawTextW(hdc, title_text.c_str(), -1, &title_rect, DT_CENTER | DT_TOP | DT_SINGLELINE | DT_END_ELLIPSIS);

    HFONT subtitle_restore_font = old_font;
    if (subtitle_font != nullptr) {
        subtitle_restore_font = static_cast<HFONT>(SelectObject(hdc, subtitle_font));
    }
    SetTextColor(hdc, kEmptyStateTextColor);

    RECT subtitle_rect = {};
    subtitle_rect.left = client_rect.left + MulDiv(24, static_cast<int>(dpi_), 96);
    subtitle_rect.right = client_rect.right - MulDiv(24, static_cast<int>(dpi_), 96);
    subtitle_rect.top = subtitle_top;
    subtitle_rect.bottom = subtitle_top + MulDiv(20, static_cast<int>(dpi_), 96);
    DrawTextW(
        hdc,
        subtitle_text.c_str(),
        -1,
        &subtitle_rect,
        DT_CENTER | DT_TOP | DT_SINGLELINE | DT_END_ELLIPSIS);

    if (subtitle_font != nullptr) {
        SelectObject(hdc, subtitle_restore_font);
        DeleteObject(subtitle_font);
    }

    if (title_font != nullptr) {
        if (old_font != nullptr) {
            SelectObject(hdc, old_font);
        }
        DeleteObject(title_font);
    }
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

std::wstring FileListView::FormatWin32ErrorMessage(DWORD error_code) {
    wchar_t message_buffer[512] = {};
    const DWORD written = FormatMessageW(
        FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
        nullptr,
        error_code,
        0,
        message_buffer,
        static_cast<DWORD>(std::size(message_buffer)),
        nullptr);
    if (written == 0U) {
        wchar_t fallback[64] = {};
        swprintf_s(fallback, L"Error code: %lu", error_code);
        return fallback;
    }

    std::wstring message = message_buffer;
    while (!message.empty() &&
           (message.back() == L'\r' || message.back() == L'\n' || message.back() == L' ')) {
        message.pop_back();
    }
    return message;
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
        return fileexplorer::colors::kFilePdf;
    }
    if (lower == L".xls" || lower == L".xlsx" || lower == L".xlsm") {
        return fileexplorer::colors::kFileExcel;
    }
    if (lower == L".doc" || lower == L".docx") {
        return fileexplorer::colors::kFileWord;
    }
    if (lower == L".zip" || lower == L".7z" || lower == L".rar" || lower == L".tar" || lower == L".gz") {
        return fileexplorer::colors::kFileArchive;
    }
    if (lower == L".txt" || lower == L".md" || lower == L".log") {
        return fileexplorer::colors::kFileText;
    }
    return CLR_INVALID;
}

int FileListView::ClampColumnWidthLogical(int column_index, int value) noexcept {
    int min_width = kMinColumnWidthLogical;
    if (column_index == 0) {
        min_width = 120;
    } else if (column_index == 4) {
        min_width = 120;
    }
    return (std::max)(min_width, (std::min)(kMaxColumnWidthLogical, value));
}

int FileListView::ClampDetailScalePercent(int value) noexcept {
    return (std::max)(kDetailScaleMinPercent, (std::min)(kDetailScaleMaxPercent, value));
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
