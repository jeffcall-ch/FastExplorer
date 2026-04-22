#include "SessionStore.h"

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>
#include <ShlObj.h>

#include <algorithm>
#include <cwchar>
#include <utility>

namespace {

constexpr wchar_t kSessionFileName[] = L"session.ini";
constexpr wchar_t kSessionSection[] = L"Session";
constexpr wchar_t kTabCountKey[] = L"TabCount";
constexpr wchar_t kActiveTabKey[] = L"ActiveTab";

void LogLastError(const wchar_t* context) {
    const DWORD error_code = GetLastError();
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (GetLastError=%lu).\r\n", context, error_code);
    OutputDebugStringW(buffer);
}

std::wstring BuildTabSectionName(int index) {
    wchar_t section[32] = {};
    swprintf_s(section, L"Tab%d", index);
    return section;
}

}  // namespace

namespace fileexplorer {

SessionStore::SessionStore() {
    UseDefaultStoragePath();
}

void SessionStore::SetStoragePath(std::wstring storage_path) {
    storage_path_ = std::move(storage_path);
    if (storage_path_.empty()) {
        UseDefaultStoragePath();
    }
}

void SessionStore::UseDefaultStoragePath() {
    storage_path_ = ResolveDefaultStoragePath();
}

bool SessionStore::Save(const std::vector<SessionTab>& tabs, int active_index) const {
    if (storage_path_.empty()) {
        return false;
    }

    const std::wstring parent = ParentDirectory(storage_path_);
    if (!parent.empty() && !EnsureDirectoryExists(parent)) {
        return false;
    }

    const int tab_count = static_cast<int>(tabs.size());
    const int clamped_active = tab_count > 0
        ? (std::max)(0, (std::min)(tab_count - 1, active_index))
        : 0;

    wchar_t value_buffer[32] = {};
    swprintf_s(value_buffer, L"%d", tab_count);
    if (!WritePrivateProfileStringW(kSessionSection, kTabCountKey, value_buffer, storage_path_.c_str())) {
        LogLastError(L"WritePrivateProfileStringW(Session TabCount)");
        return false;
    }

    swprintf_s(value_buffer, L"%d", clamped_active);
    if (!WritePrivateProfileStringW(kSessionSection, kActiveTabKey, value_buffer, storage_path_.c_str())) {
        LogLastError(L"WritePrivateProfileStringW(Session ActiveTab)");
        return false;
    }

    for (int i = 0; i < tab_count; ++i) {
        const std::wstring section = BuildTabSectionName(i);
        const std::wstring path = NormalizePath(tabs[static_cast<size_t>(i)].path);
        if (!WritePrivateProfileStringW(section.c_str(), L"Path", path.c_str(), storage_path_.c_str())) {
            LogLastError(L"WritePrivateProfileStringW(Session Tab Path)");
            return false;
        }

        const wchar_t* pinned_text = tabs[static_cast<size_t>(i)].pinned ? L"1" : L"0";
        if (!WritePrivateProfileStringW(section.c_str(), L"Pinned", pinned_text, storage_path_.c_str())) {
            LogLastError(L"WritePrivateProfileStringW(Session Tab Pinned)");
            return false;
        }
    }

    return true;
}

std::vector<SessionStore::SessionTab> SessionStore::Load(int* active_index) const {
    if (active_index != nullptr) {
        *active_index = 0;
    }

    std::vector<SessionTab> tabs;
    if (storage_path_.empty()) {
        return tabs;
    }

    const DWORD attributes = GetFileAttributesW(storage_path_.c_str());
    if (attributes == INVALID_FILE_ATTRIBUTES || (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
        return tabs;
    }

    const int tab_count = static_cast<int>(GetPrivateProfileIntW(
        kSessionSection,
        kTabCountKey,
        0,
        storage_path_.c_str()));
    if (tab_count <= 0 || tab_count > 64) {
        return tabs;
    }

    int loaded_active_index = static_cast<int>(GetPrivateProfileIntW(
        kSessionSection,
        kActiveTabKey,
        0,
        storage_path_.c_str()));

    tabs.reserve(static_cast<size_t>(tab_count));
    for (int i = 0; i < tab_count; ++i) {
        const std::wstring section = BuildTabSectionName(i);

        wchar_t path_buffer[4096] = {};
        const DWORD chars = GetPrivateProfileStringW(
            section.c_str(),
            L"Path",
            L"",
            path_buffer,
            static_cast<DWORD>(_countof(path_buffer)),
            storage_path_.c_str());
        if (chars == 0) {
            continue;
        }

        SessionTab tab = {};
        tab.path = NormalizePath(std::wstring(path_buffer, path_buffer + chars));
        if (tab.path.empty()) {
            continue;
        }
        tab.pinned = GetPrivateProfileIntW(section.c_str(), L"Pinned", 0, storage_path_.c_str()) != 0;
        tabs.push_back(std::move(tab));
    }

    if (tabs.empty()) {
        if (active_index != nullptr) {
            *active_index = 0;
        }
        return tabs;
    }

    loaded_active_index = (std::max)(0, (std::min)(loaded_active_index, static_cast<int>(tabs.size()) - 1));
    if (active_index != nullptr) {
        *active_index = loaded_active_index;
    }

    return tabs;
}

const std::wstring& SessionStore::storage_path() const noexcept {
    return storage_path_;
}

std::wstring SessionStore::ResolveDefaultStorageRoot() {
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
        LogLastError(L"SHGetFolderPathW(SessionStore)");
        return exe_directory;
    }

    std::wstring root = JoinPath(appdata, L"FileExplorer");
    if (!EnsureDirectoryExists(root)) {
        return exe_directory.empty() ? root : exe_directory;
    }
    return root;
}

std::wstring SessionStore::ResolveDefaultStoragePath() {
    return JoinPath(ResolveDefaultStorageRoot(), kSessionFileName);
}

std::wstring SessionStore::ExecutableDirectory() {
    wchar_t path[MAX_PATH] = {};
    const DWORD path_capacity = static_cast<DWORD>(_countof(path));
    const DWORD length = GetModuleFileNameW(nullptr, path, path_capacity);
    if (length == 0 || length >= path_capacity) {
        LogLastError(L"GetModuleFileNameW(SessionStore)");
        return L"";
    }

    std::wstring full_path(path, path + length);
    const std::wstring::size_type separator = full_path.find_last_of(L'\\');
    if (separator == std::wstring::npos) {
        return L"";
    }
    return full_path.substr(0, separator);
}

std::wstring SessionStore::ParentDirectory(const std::wstring& path) {
    const std::wstring::size_type separator = path.find_last_of(L'\\');
    if (separator == std::wstring::npos) {
        return L"";
    }
    return path.substr(0, separator);
}

std::wstring SessionStore::JoinPath(const std::wstring& base, const std::wstring& child) {
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

bool SessionStore::EnsureDirectoryExists(const std::wstring& path) {
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

bool SessionStore::IsDirectoryWritable(const std::wstring& directory) {
    if (directory.empty()) {
        return false;
    }

    std::wstring probe = JoinPath(directory, L".fe_session_probe.tmp");
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
        LogLastError(L"CloseHandle(SessionStore Probe)");
    }
    DeleteFileW(probe.c_str());
    return true;
}

std::wstring SessionStore::NormalizePath(std::wstring path) {
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

}  // namespace fileexplorer
