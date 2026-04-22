#include "SortSettings.h"

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>
#include <ShlObj.h>

#include <algorithm>
#include <cwchar>
#include <utility>

#include "CsvParser.h"

namespace {

constexpr wchar_t kSortSettingsFileName[] = L"sort_settings.csv";

void LogLastError(const wchar_t* context) {
    const DWORD error_code = GetLastError();
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (GetLastError=%lu).\r\n", context, error_code);
    OutputDebugStringW(buffer);
}

}  // namespace

namespace fileexplorer {

SortSettings::SortSettings() {
    UseDefaultStoragePath();
}

void SortSettings::SetStoragePath(std::wstring storage_path) {
    storage_path_ = std::move(storage_path);
    if (storage_path_.empty()) {
        UseDefaultStoragePath();
    }
}

void SortSettings::UseDefaultStoragePath() {
    storage_path_ = ResolveDefaultStoragePath();
}

bool SortSettings::Load() {
    entries_.clear();

    if (storage_path_.empty()) {
        return true;
    }

    CsvParser::CsvRows rows;
    if (!CsvParser::ReadFile(storage_path_, &rows, nullptr)) {
        return false;
    }

    entries_.reserve((std::min)(rows.size(), kMaxEntries));
    for (const auto& row : rows) {
        if (row.size() < 3) {
            continue;
        }

        const std::wstring folder_path = NormalizePath(row[0]);
        if (folder_path.empty()) {
            continue;
        }

        SortColumn sort_column = SortColumn::Name;
        SortDirection sort_direction = SortDirection::Ascending;
        if (!TryParseColumn(row[1], &sort_column) || !TryParseDirection(row[2], &sort_direction)) {
            continue;
        }

        bool duplicate = false;
        for (const Entry& existing : entries_) {
            if (ComparePathInsensitive(existing.folder_path, folder_path) == 0) {
                duplicate = true;
                break;
            }
        }
        if (duplicate) {
            continue;
        }

        entries_.push_back({folder_path, sort_column, sort_direction});
        if (entries_.size() >= kMaxEntries) {
            break;
        }
    }

    return true;
}

bool SortSettings::Save() const {
    if (storage_path_.empty()) {
        return false;
    }

    const std::wstring parent = ParentDirectory(storage_path_);
    if (!parent.empty() && !EnsureDirectoryExists(parent)) {
        return false;
    }

    CsvParser::CsvRows rows;
    rows.reserve(entries_.size());
    for (const Entry& entry : entries_) {
        rows.push_back({entry.folder_path, ColumnToString(entry.sort_column), DirectionToString(entry.sort_direction)});
    }
    return CsvParser::WriteFile(storage_path_, rows);
}

void SortSettings::Set(const std::wstring& folder_path, SortColumn sort_column, SortDirection sort_direction) {
    const std::wstring normalized_path = NormalizePath(folder_path);
    if (normalized_path.empty()) {
        return;
    }

    auto existing = std::find_if(
        entries_.begin(),
        entries_.end(),
        [&normalized_path](const Entry& entry) {
            return ComparePathInsensitive(entry.folder_path, normalized_path) == 0;
        });
    if (existing != entries_.end()) {
        existing->sort_column = sort_column;
        existing->sort_direction = sort_direction;
        Entry updated = *existing;
        entries_.erase(existing);
        entries_.insert(entries_.begin(), std::move(updated));
    } else {
        entries_.insert(entries_.begin(), Entry{normalized_path, sort_column, sort_direction});
    }

    if (entries_.size() > kMaxEntries) {
        entries_.resize(kMaxEntries);
    }
}

SortSettings::Entry SortSettings::Get(const std::wstring& folder_path) const {
    const std::wstring normalized_path = NormalizePath(folder_path);
    if (normalized_path.empty()) {
        return {};
    }

    const auto it = std::find_if(
        entries_.begin(),
        entries_.end(),
        [&normalized_path](const Entry& entry) {
            return ComparePathInsensitive(entry.folder_path, normalized_path) == 0;
        });
    if (it == entries_.end()) {
        return {};
    }
    return *it;
}

const std::wstring& SortSettings::storage_path() const noexcept {
    return storage_path_;
}

std::wstring SortSettings::ResolveDefaultStorageRoot() {
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
        LogLastError(L"SHGetFolderPathW(SortSettings)");
        return exe_directory;
    }

    std::wstring root = JoinPath(appdata, L"FileExplorer");
    if (!EnsureDirectoryExists(root)) {
        return exe_directory.empty() ? root : exe_directory;
    }
    return root;
}

std::wstring SortSettings::ResolveDefaultStoragePath() {
    return JoinPath(ResolveDefaultStorageRoot(), kSortSettingsFileName);
}

std::wstring SortSettings::ExecutableDirectory() {
    wchar_t path[MAX_PATH] = {};
    const DWORD path_capacity = static_cast<DWORD>(_countof(path));
    const DWORD length = GetModuleFileNameW(nullptr, path, path_capacity);
    if (length == 0 || length >= path_capacity) {
        LogLastError(L"GetModuleFileNameW(SortSettings)");
        return L"";
    }

    std::wstring full_path(path, path + length);
    const std::wstring::size_type separator = full_path.find_last_of(L'\\');
    if (separator == std::wstring::npos) {
        return L"";
    }
    return full_path.substr(0, separator);
}

std::wstring SortSettings::ParentDirectory(const std::wstring& path) {
    const std::wstring::size_type separator = path.find_last_of(L'\\');
    if (separator == std::wstring::npos) {
        return L"";
    }
    return path.substr(0, separator);
}

std::wstring SortSettings::JoinPath(const std::wstring& base, const std::wstring& child) {
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

bool SortSettings::EnsureDirectoryExists(const std::wstring& path) {
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

bool SortSettings::IsDirectoryWritable(const std::wstring& directory) {
    if (directory.empty()) {
        return false;
    }

    std::wstring probe = JoinPath(directory, L".fe_sort_probe.tmp");
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
        LogLastError(L"CloseHandle(SortSettings Probe)");
    }
    DeleteFileW(probe.c_str());
    return true;
}

std::wstring SortSettings::NormalizePath(std::wstring path) {
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

int SortSettings::ComparePathInsensitive(const std::wstring& left, const std::wstring& right) {
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

bool SortSettings::TryParseColumn(const std::wstring& value, SortColumn* sort_column) {
    if (sort_column == nullptr) {
        return false;
    }

    wchar_t* end = nullptr;
    const long parsed = wcstol(value.c_str(), &end, 10);
    if (end == value.c_str() || parsed < 0L || parsed > 4L) {
        return false;
    }
    *sort_column = static_cast<SortColumn>(parsed);
    return true;
}

bool SortSettings::TryParseDirection(const std::wstring& value, SortDirection* sort_direction) {
    if (sort_direction == nullptr) {
        return false;
    }

    wchar_t* end = nullptr;
    const long parsed = wcstol(value.c_str(), &end, 10);
    if (end == value.c_str() || (parsed != 0L && parsed != 1L)) {
        return false;
    }
    *sort_direction = (parsed == 0L) ? SortDirection::Ascending : SortDirection::Descending;
    return true;
}

std::wstring SortSettings::ColumnToString(SortColumn sort_column) {
    return std::to_wstring(static_cast<int>(sort_column));
}

std::wstring SortSettings::DirectionToString(SortDirection sort_direction) {
    return std::to_wstring(sort_direction == SortDirection::Ascending ? 0 : 1);
}

}  // namespace fileexplorer
