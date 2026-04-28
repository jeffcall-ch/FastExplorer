#include "FileOps.h"

#include <Shellapi.h>
#include <ShlObj.h>
#include <Shlwapi.h>

#include <cstring>
#include <cwchar>
#include <utility>

namespace {

void LogLastError(const wchar_t* context) {
    const DWORD error_code = GetLastError();
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (GetLastError=%lu).\r\n", context, error_code);
    OutputDebugStringW(buffer);
}

void LogShellError(const wchar_t* context, int error_code) {
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (SHFileOperation=%d).\r\n", context, error_code);
    OutputDebugStringW(buffer);
}

bool FillDropFilesBlock(
    HGLOBAL global_handle,
    const std::vector<std::wstring>& paths,
    bool wide,
    DROPFILES* out_header) {
    if (global_handle == nullptr || out_header == nullptr) {
        return false;
    }

    void* block = GlobalLock(global_handle);
    if (block == nullptr) {
        LogLastError(L"GlobalLock(DROPFILES)");
        return false;
    }

    DROPFILES* drop = reinterpret_cast<DROPFILES*>(block);
    ZeroMemory(drop, sizeof(*drop));
    drop->pFiles = sizeof(DROPFILES);
    drop->fWide = wide ? TRUE : FALSE;

    auto* cursor = reinterpret_cast<wchar_t*>(reinterpret_cast<BYTE*>(block) + sizeof(DROPFILES));
    for (const std::wstring& path : paths) {
        if (path.empty()) {
            continue;
        }
        memcpy(cursor, path.c_str(), path.size() * sizeof(wchar_t));
        cursor += path.size();
        *cursor++ = L'\0';
    }
    *cursor = L'\0';

    *out_header = *drop;
    GlobalUnlock(global_handle);
    return true;
}

bool SetClipboardDropEffect(DWORD effect) {
    const UINT format = RegisterClipboardFormatW(L"Preferred DropEffect");
    if (format == 0U) {
        LogLastError(L"RegisterClipboardFormatW(PreferredDropEffect)");
        return false;
    }

    HGLOBAL effect_handle = GlobalAlloc(GMEM_MOVEABLE, sizeof(DWORD));
    if (effect_handle == nullptr) {
        LogLastError(L"GlobalAlloc(PreferredDropEffect)");
        return false;
    }

    void* block = GlobalLock(effect_handle);
    if (block == nullptr) {
        LogLastError(L"GlobalLock(PreferredDropEffect)");
        GlobalFree(effect_handle);
        return false;
    }
    *reinterpret_cast<DWORD*>(block) = effect;
    GlobalUnlock(effect_handle);

    if (SetClipboardData(format, effect_handle) == nullptr) {
        LogLastError(L"SetClipboardData(PreferredDropEffect)");
        GlobalFree(effect_handle);
        return false;
    }

    return true;
}

std::vector<std::wstring> ExtractClipboardDropPaths(HDROP drop_handle) {
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

DWORD ReadPreferredDropEffectFromClipboard() {
    const UINT format = RegisterClipboardFormatW(L"Preferred DropEffect");
    if (format == 0U) {
        return DROPEFFECT_COPY;
    }

    HANDLE effect_handle = GetClipboardData(format);
    if (effect_handle == nullptr) {
        return DROPEFFECT_COPY;
    }

    const void* block = GlobalLock(effect_handle);
    if (block == nullptr) {
        return DROPEFFECT_COPY;
    }

    const DWORD effect = *reinterpret_cast<const DWORD*>(block);
    GlobalUnlock(effect_handle);
    return effect;
}

std::wstring JoinPath(const std::wstring& folder, const std::wstring& leaf_name) {
    std::wstring result = folder;
    if (!result.empty() && result.back() != L'\\') {
        result.push_back(L'\\');
    }
    result.append(leaf_name);
    return result;
}

bool PathExists(const std::wstring& path) {
    return GetFileAttributesW(path.c_str()) != INVALID_FILE_ATTRIBUTES;
}

std::wstring BuildCopyLeafName(const std::wstring& leaf_name, bool is_directory, int index) {
    wchar_t index_text[32] = {};
    swprintf_s(index_text, L"%d", index);

    if (is_directory) {
        std::wstring result = leaf_name;
        result.append(L" copy");
        result.append(index_text);
        return result;
    }

    const wchar_t* extension_ptr = PathFindExtensionW(leaf_name.c_str());
    if (extension_ptr == nullptr || *extension_ptr == L'\0') {
        std::wstring result = leaf_name;
        result.append(L" copy");
        result.append(index_text);
        return result;
    }

    const size_t extension_offset = static_cast<size_t>(extension_ptr - leaf_name.c_str());
    if (extension_offset == 0) {
        std::wstring result = leaf_name;
        result.append(L" copy");
        result.append(index_text);
        return result;
    }
    std::wstring result = leaf_name.substr(0, extension_offset);
    result.append(L" copy");
    result.append(index_text);
    result.push_back(L' ');
    result.append(extension_ptr);
    return result;
}

std::wstring BuildCollisionFreeDestinationPath(
    const std::wstring& destination_folder,
    const std::wstring& source_path) {
    const DWORD source_attributes = GetFileAttributesW(source_path.c_str());
    const bool source_is_directory =
        (source_attributes != INVALID_FILE_ATTRIBUTES) &&
        ((source_attributes & FILE_ATTRIBUTE_DIRECTORY) != 0);

    const std::wstring source_leaf = PathFindFileNameW(source_path.c_str());
    if (source_leaf.empty()) {
        return L"";
    }

    const std::wstring base_target_path = JoinPath(destination_folder, source_leaf);
    if (!PathExists(base_target_path)) {
        return base_target_path;
    }

    for (int index = 1; index <= 9999; ++index) {
        const std::wstring candidate_leaf = BuildCopyLeafName(source_leaf, source_is_directory, index);
        const std::wstring candidate_path = JoinPath(destination_folder, candidate_leaf);
        if (!PathExists(candidate_path)) {
            return candidate_path;
        }
    }

    return L"";
}

}  // namespace

namespace fileexplorer {

std::wstring FileOps::BuildMultiStringPaths(const std::vector<std::wstring>& paths) {
    std::wstring multi;
    size_t total_chars = 1;
    for (const std::wstring& path : paths) {
        if (!path.empty()) {
            total_chars += path.size() + 1;
        }
    }

    multi.reserve(total_chars);
    for (const std::wstring& path : paths) {
        if (path.empty()) {
            continue;
        }
        multi.append(path);
        multi.push_back(L'\0');
    }
    multi.push_back(L'\0');
    return multi;
}

bool FileOps::CopyPathsToClipboard(HWND owner, const std::vector<std::wstring>& paths, bool cut) {
    if (paths.empty()) {
        return false;
    }

    if (!OpenClipboard(owner)) {
        LogLastError(L"OpenClipboard(CF_HDROP)");
        return false;
    }

    if (!EmptyClipboard()) {
        LogLastError(L"EmptyClipboard(CF_HDROP)");
        CloseClipboard();
        return false;
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
        LogLastError(L"GlobalAlloc(CF_HDROP)");
        CloseClipboard();
        return false;
    }

    DROPFILES header = {};
    if (!FillDropFilesBlock(drop_handle, paths, true, &header)) {
        GlobalFree(drop_handle);
        CloseClipboard();
        return false;
    }
    (void)header;

    if (SetClipboardData(CF_HDROP, drop_handle) == nullptr) {
        LogLastError(L"SetClipboardData(CF_HDROP)");
        GlobalFree(drop_handle);
        CloseClipboard();
        return false;
    }

    if (cut) {
        if (!SetClipboardDropEffect(DROPEFFECT_MOVE)) {
            CloseClipboard();
            return false;
        }
    }

    CloseClipboard();
    return true;
}

bool FileOps::ClipboardHasDropFiles() {
    return IsClipboardFormatAvailable(CF_HDROP) != FALSE;
}

bool FileOps::PerformShellOperation(
    HWND owner,
    UINT operation,
    const std::vector<std::wstring>& from_paths,
    const std::wstring& to_path,
    FILEOP_FLAGS flags,
    bool* any_aborted) {
    if (from_paths.empty()) {
        return false;
    }

    const std::wstring from_multi = BuildMultiStringPaths(from_paths);
    std::wstring to_multi;
    LPCWSTR to_buffer = nullptr;
    if (operation == FO_COPY || operation == FO_MOVE) {
        to_multi = to_path;
        to_multi.push_back(L'\0');
        to_multi.push_back(L'\0');
        to_buffer = to_multi.c_str();
    }

    SHFILEOPSTRUCTW file_operation = {};
    file_operation.hwnd = owner;
    file_operation.wFunc = operation;
    file_operation.pFrom = from_multi.c_str();
    file_operation.pTo = to_buffer;
    file_operation.fFlags = flags;

    const int result = SHFileOperationW(&file_operation);
    if (any_aborted != nullptr) {
        *any_aborted = file_operation.fAnyOperationsAborted != FALSE;
    }
    if (result != 0) {
        LogShellError(L"SHFileOperationW", result);
        return false;
    }
    if (file_operation.fAnyOperationsAborted != FALSE) {
        return false;
    }
    return true;
}

bool FileOps::PasteFromClipboard(HWND owner, const std::wstring& destination_folder, bool* was_move) {
    if (!OpenClipboard(owner)) {
        LogLastError(L"OpenClipboard(Paste)");
        return false;
    }

    HANDLE drop_data = GetClipboardData(CF_HDROP);
    if (drop_data == nullptr) {
        CloseClipboard();
        return false;
    }

    const std::vector<std::wstring> paths = ExtractClipboardDropPaths(reinterpret_cast<HDROP>(drop_data));
    const DWORD effect = ReadPreferredDropEffectFromClipboard();
    CloseClipboard();

    if (paths.empty()) {
        return false;
    }

    const bool move = (effect & DROPEFFECT_MOVE) != 0;
    if (was_move != nullptr) {
        *was_move = move;
    }

    if (move) {
        const FILEOP_FLAGS flags = FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR;
        bool aborted = false;
        return PerformShellOperation(owner, FO_MOVE, paths, destination_folder, flags, &aborted);
    }

    // Copy operations are executed item-by-item so name collisions can be auto-renamed
    // to "<name> copy <x>" without prompting.
    bool copied_any = false;
    for (const std::wstring& source_path : paths) {
        const std::wstring target_path = BuildCollisionFreeDestinationPath(destination_folder, source_path);
        if (target_path.empty()) {
            return false;
        }

        const FILEOP_FLAGS flags = FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOF_NOCONFIRMATION | FOF_NOERRORUI;
        bool aborted = false;
        if (!PerformShellOperation(owner, FO_COPY, {source_path}, target_path, flags, &aborted)) {
            return false;
        }
        copied_any = true;
    }

    return copied_any;
}

bool FileOps::ImportPaths(
    HWND owner,
    const std::vector<std::wstring>& source_paths,
    const std::wstring& destination_folder,
    bool move) {
    if (source_paths.empty() || destination_folder.empty()) {
        return false;
    }

    const UINT operation = move ? FO_MOVE : FO_COPY;
    const FILEOP_FLAGS flags = FOF_ALLOWUNDO | FOF_NOCONFIRMMKDIR | FOF_NOCONFIRMATION;
    bool aborted = false;
    return PerformShellOperation(owner, operation, source_paths, destination_folder, flags, &aborted);
}

bool FileOps::DeleteToRecycleBin(HWND owner, const std::vector<std::wstring>& paths) {
    bool aborted = false;
    const FILEOP_FLAGS flags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION;
    return PerformShellOperation(owner, FO_DELETE, paths, L"", flags, &aborted);
}

bool FileOps::DeletePermanently(HWND owner, const std::vector<std::wstring>& paths) {
    bool aborted = false;
    const FILEOP_FLAGS flags = FOF_NOCONFIRMATION | FOF_WANTNUKEWARNING;
    return PerformShellOperation(owner, FO_DELETE, paths, L"", flags, &aborted);
}

bool FileOps::RenamePath(const std::wstring& from_path, const std::wstring& to_path, DWORD* error_code) {
    const DWORD source_attributes = GetFileAttributesW(from_path.c_str());
    const bool source_is_file =
        (source_attributes != INVALID_FILE_ATTRIBUTES) &&
        ((source_attributes & FILE_ATTRIBUTE_DIRECTORY) == 0);

    if (source_is_file) {
        HANDLE probe_handle = CreateFileW(
            from_path.c_str(),
            0,
            0,
            nullptr,
            OPEN_EXISTING,
            FILE_ATTRIBUTE_NORMAL,
            nullptr);
        if (probe_handle == INVALID_HANDLE_VALUE) {
            const DWORD probe_error = GetLastError();
            if (probe_error == ERROR_SHARING_VIOLATION || probe_error == ERROR_LOCK_VIOLATION) {
                if (error_code != nullptr) {
                    *error_code = probe_error;
                }
                return false;
            }
        } else {
            CloseHandle(probe_handle);
        }
    }

    if (MoveFileW(from_path.c_str(), to_path.c_str())) {
        if (error_code != nullptr) {
            *error_code = ERROR_SUCCESS;
        }
        return true;
    }

    if (error_code != nullptr) {
        *error_code = GetLastError();
    } else {
        LogLastError(L"MoveFileW(Rename)");
    }
    return false;
}

bool FileOps::CreateDirectoryPath(const std::wstring& path, DWORD* error_code) {
    if (CreateDirectoryW(path.c_str(), nullptr)) {
        if (error_code != nullptr) {
            *error_code = ERROR_SUCCESS;
        }
        return true;
    }

    if (error_code != nullptr) {
        *error_code = GetLastError();
    } else {
        LogLastError(L"CreateDirectoryW");
    }
    return false;
}

bool FileOps::CopyTextToClipboard(HWND owner, const std::wstring& text) {
    if (!OpenClipboard(owner)) {
        LogLastError(L"OpenClipboard(CF_UNICODETEXT)");
        return false;
    }

    if (!EmptyClipboard()) {
        LogLastError(L"EmptyClipboard(CF_UNICODETEXT)");
        CloseClipboard();
        return false;
    }

    const SIZE_T bytes = (text.size() + 1) * sizeof(wchar_t);
    HGLOBAL global_handle = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, bytes);
    if (global_handle == nullptr) {
        LogLastError(L"GlobalAlloc(CF_UNICODETEXT)");
        CloseClipboard();
        return false;
    }

    void* block = GlobalLock(global_handle);
    if (block == nullptr) {
        LogLastError(L"GlobalLock(CF_UNICODETEXT)");
        GlobalFree(global_handle);
        CloseClipboard();
        return false;
    }

    memcpy(block, text.c_str(), bytes);
    GlobalUnlock(global_handle);

    if (SetClipboardData(CF_UNICODETEXT, global_handle) == nullptr) {
        LogLastError(L"SetClipboardData(CF_UNICODETEXT)");
        GlobalFree(global_handle);
        CloseClipboard();
        return false;
    }

    CloseClipboard();
    return true;
}

bool FileOps::ShowProperties(HWND owner, const std::wstring& path) {
    if (path.empty()) {
        return false;
    }

    if (!SHObjectProperties(owner, SHOP_FILEPATH, path.c_str(), nullptr)) {
        LogLastError(L"SHObjectProperties");
        return false;
    }
    return true;
}

}  // namespace fileexplorer
