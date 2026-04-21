#include "Settings.h"

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>
#include <ShlObj.h>

#include <algorithm>
#include <cwchar>
#include <utility>

namespace {

constexpr wchar_t kSettingsFileName[] = L"settings.ini";
constexpr wchar_t kGeneralSection[] = L"General";
constexpr wchar_t kSidebarWidthKey[] = L"SidebarWidth";
constexpr wchar_t kFileListColumnWidthKeys[fileexplorer::Settings::kFileListColumnCount][32] = {
    L"FileListNameWidth",
    L"FileListExtensionWidth",
    L"FileListDateModifiedWidth",
    L"FileListSizeWidth",
    L"FileListPathWidth",
};

constexpr int kSidebarWidthMinLogical = 180;
constexpr int kSidebarWidthMaxLogical = 900;
constexpr int kColumnWidthMinLogical = 40;
constexpr int kColumnWidthMaxLogical = 1800;

void LogLastError(const wchar_t* context) {
    const DWORD error_code = GetLastError();
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (GetLastError=%lu).\r\n", context, error_code);
    OutputDebugStringW(buffer);
}

}  // namespace

namespace fileexplorer {

Settings::Settings() {
    UseDefaultStoragePath();
}

void Settings::SetStoragePath(std::wstring storage_path) {
    storage_path_ = std::move(storage_path);
    if (storage_path_.empty()) {
        UseDefaultStoragePath();
    }
}

void Settings::UseDefaultStoragePath() {
    storage_path_ = ResolveDefaultStoragePath();
}

bool Settings::Load(Values* values) const {
    if (values == nullptr) {
        return false;
    }

    Values loaded = *values;
    if (storage_path_.empty()) {
        loaded.sidebar_width_logical = ClampSidebarWidthLogical(loaded.sidebar_width_logical);
        for (int i = 0; i < kFileListColumnCount; ++i) {
            loaded.file_list_column_widths_logical[static_cast<size_t>(i)] = ClampFileListColumnWidthLogical(
                i,
                loaded.file_list_column_widths_logical[static_cast<size_t>(i)]);
        }
        *values = loaded;
        return true;
    }

    const int stored_sidebar_width = static_cast<int>(GetPrivateProfileIntW(
        kGeneralSection,
        kSidebarWidthKey,
        loaded.sidebar_width_logical,
        storage_path_.c_str()));
    loaded.sidebar_width_logical = ClampSidebarWidthLogical(stored_sidebar_width);
    for (int i = 0; i < kFileListColumnCount; ++i) {
        const int stored_width = static_cast<int>(GetPrivateProfileIntW(
            kGeneralSection,
            kFileListColumnWidthKeys[i],
            loaded.file_list_column_widths_logical[static_cast<size_t>(i)],
            storage_path_.c_str()));
        loaded.file_list_column_widths_logical[static_cast<size_t>(i)] =
            ClampFileListColumnWidthLogical(i, stored_width);
    }
    *values = loaded;
    return true;
}

bool Settings::Save(const Values& values) const {
    if (storage_path_.empty()) {
        return false;
    }

    const std::wstring parent = ParentDirectory(storage_path_);
    if (!parent.empty() && !EnsureDirectoryExists(parent)) {
        return false;
    }

    wchar_t buffer[32] = {};
    swprintf_s(buffer, L"%d", ClampSidebarWidthLogical(values.sidebar_width_logical));
    if (!WritePrivateProfileStringW(kGeneralSection, kSidebarWidthKey, buffer, storage_path_.c_str())) {
        LogLastError(L"WritePrivateProfileStringW(Settings)");
        return false;
    }

    for (int i = 0; i < kFileListColumnCount; ++i) {
        swprintf_s(
            buffer,
            L"%d",
            ClampFileListColumnWidthLogical(
                i,
                values.file_list_column_widths_logical[static_cast<size_t>(i)]));
        if (!WritePrivateProfileStringW(
                kGeneralSection,
                kFileListColumnWidthKeys[i],
                buffer,
                storage_path_.c_str())) {
            LogLastError(L"WritePrivateProfileStringW(Settings ColumnWidth)");
            return false;
        }
    }
    return true;
}

const std::wstring& Settings::storage_path() const noexcept {
    return storage_path_;
}

int Settings::ClampSidebarWidthLogical(int value) noexcept {
    return (std::max)(kSidebarWidthMinLogical, (std::min)(kSidebarWidthMaxLogical, value));
}

int Settings::ClampFileListColumnWidthLogical(int column_index, int value) noexcept {
    int min_width = kColumnWidthMinLogical;
    if (column_index == 0 || column_index == 4) {
        min_width = 120;
    }
    return (std::max)(min_width, (std::min)(kColumnWidthMaxLogical, value));
}

std::wstring Settings::ResolveDefaultStorageRoot() {
    const std::wstring exe_directory = ExecutableDirectory();
    if (!exe_directory.empty() && IsDirectoryWritable(exe_directory)) {
        return exe_directory;
    }

    wchar_t appdata[MAX_PATH] = {};
    if (SHGetFolderPathW(
            nullptr,
            CSIDL_APPDATA | CSIDL_FLAG_CREATE,
            nullptr,
            SHGFP_TYPE_CURRENT,
            appdata) != S_OK) {
        LogLastError(L"SHGetFolderPathW(Settings)");
        return exe_directory;
    }

    std::wstring root = JoinPath(appdata, L"FileExplorer");
    if (!EnsureDirectoryExists(root)) {
        return exe_directory.empty() ? root : exe_directory;
    }
    return root;
}

std::wstring Settings::ResolveDefaultStoragePath() {
    const std::wstring root = ResolveDefaultStorageRoot();
    return JoinPath(root, kSettingsFileName);
}

std::wstring Settings::ExecutableDirectory() {
    wchar_t path[MAX_PATH] = {};
    const DWORD path_capacity = static_cast<DWORD>(_countof(path));
    const DWORD length = GetModuleFileNameW(nullptr, path, path_capacity);
    if (length == 0 || length >= path_capacity) {
        LogLastError(L"GetModuleFileNameW(Settings)");
        return L"";
    }

    std::wstring full_path(path, path + length);
    const std::wstring::size_type separator = full_path.find_last_of(L'\\');
    if (separator == std::wstring::npos) {
        return L"";
    }
    return full_path.substr(0, separator);
}

std::wstring Settings::ParentDirectory(const std::wstring& path) {
    const std::wstring::size_type separator = path.find_last_of(L'\\');
    if (separator == std::wstring::npos) {
        return L"";
    }
    return path.substr(0, separator);
}

std::wstring Settings::JoinPath(const std::wstring& base, const std::wstring& child) {
    std::wstring joined = base;
    if (joined.empty()) {
        return child;
    }
    if (!joined.empty() && joined.back() != L'\\') {
        joined.push_back(L'\\');
    }
    joined.append(child);
    return joined;
}

bool Settings::EnsureDirectoryExists(const std::wstring& path) {
    if (path.empty()) {
        return false;
    }

    const DWORD attributes = GetFileAttributesW(path.c_str());
    if (attributes != INVALID_FILE_ATTRIBUTES) {
        return (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    }

    const int result = SHCreateDirectoryExW(nullptr, path.c_str(), nullptr);
    return result == ERROR_SUCCESS || result == ERROR_FILE_EXISTS || result == ERROR_ALREADY_EXISTS;
}

bool Settings::IsDirectoryWritable(const std::wstring& directory) {
    if (directory.empty()) {
        return false;
    }

    std::wstring probe = JoinPath(directory, L".fe_settings_probe.tmp");
    HANDLE handle = CreateFileW(
        probe.c_str(),
        GENERIC_WRITE,
        0,
        nullptr,
        CREATE_ALWAYS,
        FILE_ATTRIBUTE_TEMPORARY | FILE_FLAG_DELETE_ON_CLOSE,
        nullptr);
    if (handle == INVALID_HANDLE_VALUE) {
        return false;
    }

    const bool closed = CloseHandle(handle) != FALSE;
    if (!closed) {
        LogLastError(L"CloseHandle(Settings Probe)");
    }
    DeleteFileW(probe.c_str());
    return true;
}

}  // namespace fileexplorer
