#include "FavouritesStore.h"

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>
#include <ShlObj.h>

#include <algorithm>
#include <cwchar>
#include <cwctype>
#include <memory>

#include "CsvParser.h"

namespace {

constexpr wchar_t kFlyoutsFileName[] = L"flyout_favourites.csv";
constexpr wchar_t kRegularsFileName[] = L"regular_favourites.csv";

void LogLastError(const wchar_t* context) {
    const DWORD error_code = GetLastError();
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (GetLastError=%lu).\r\n", context, error_code);
    OutputDebugStringW(buffer);
}

std::wstring ParentDirectory(const std::wstring& path) {
    const std::wstring::size_type separator = path.find_last_of(L'\\');
    if (separator == std::wstring::npos) {
        return L"";
    }
    return path.substr(0, separator);
}

bool EnsureDirectoryExists(const std::wstring& path) {
    if (path.empty()) {
        return false;
    }

    const DWORD attributes = GetFileAttributesW(path.c_str());
    if (attributes != INVALID_FILE_ATTRIBUTES) {
        return (attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    }

    const int result = SHCreateDirectoryExW(nullptr, path.c_str(), nullptr);
    if (result == ERROR_SUCCESS || result == ERROR_FILE_EXISTS || result == ERROR_ALREADY_EXISTS) {
        return true;
    }
    return false;
}

}  // namespace

namespace fileexplorer {

FavouritesStore::FavouritesStore() {
    UseDefaultStoragePaths();
}

void FavouritesStore::SetStoragePaths(std::wstring flyouts_path, std::wstring regulars_path) {
    flyouts_path_ = std::move(flyouts_path);
    regulars_path_ = std::move(regulars_path);
}

void FavouritesStore::UseDefaultStoragePaths() {
    const std::wstring root = ResolveDefaultStorageRoot();
    flyouts_path_ = JoinPath(root, kFlyoutsFileName);
    regulars_path_ = JoinPath(root, kRegularsFileName);
}

bool FavouritesStore::Load() {
    had_existing_files_on_last_load_ = false;
    flyouts_.clear();
    regulars_.clear();

    bool flyouts_file_exists = false;
    bool regulars_file_exists = false;
    CsvParser::CsvRows flyout_rows;
    CsvParser::CsvRows regular_rows;
    if (!CsvParser::ReadFile(flyouts_path_, &flyout_rows, &flyouts_file_exists)) {
        return false;
    }
    if (!CsvParser::ReadFile(regulars_path_, &regular_rows, &regulars_file_exists)) {
        return false;
    }

    had_existing_files_on_last_load_ = flyouts_file_exists || regulars_file_exists;

    flyouts_ = ParseEntries(flyout_rows);
    regulars_ = ParseEntries(regular_rows);
    SortAndDedupe(&flyouts_);
    SortAndDedupe(&regulars_);
    return true;
}

bool FavouritesStore::Save() const {
    const std::wstring flyout_parent = ParentDirectory(flyouts_path_);
    const std::wstring regular_parent = ParentDirectory(regulars_path_);
    if (!flyout_parent.empty() && !EnsureDirectoryExists(flyout_parent)) {
        return false;
    }
    if (!regular_parent.empty() && !EnsureDirectoryExists(regular_parent)) {
        return false;
    }

    const CsvParser::CsvRows flyout_rows = SerializeEntries(flyouts_);
    const CsvParser::CsvRows regular_rows = SerializeEntries(regulars_);
    if (!CsvParser::WriteFile(flyouts_path_, flyout_rows)) {
        return false;
    }
    if (!CsvParser::WriteFile(regulars_path_, regular_rows)) {
        return false;
    }
    return true;
}

bool FavouritesStore::AddFlyout(const std::wstring& path, const std::wstring& friendly_name) {
    return AddInternal(&flyouts_, path, friendly_name);
}

bool FavouritesStore::AddRegular(const std::wstring& path, const std::wstring& friendly_name) {
    return AddInternal(&regulars_, path, friendly_name);
}

bool FavouritesStore::Remove(FavouriteType type, size_t index) {
    std::vector<FavouriteEntry>* target = (type == FavouriteType::Flyout) ? &flyouts_ : &regulars_;
    if (index >= target->size()) {
        return false;
    }

    target->erase(target->begin() + static_cast<std::ptrdiff_t>(index));
    SortAndDedupe(target);
    return Save();
}

bool FavouritesStore::Rename(FavouriteType type, size_t index, const std::wstring& new_name) {
    std::vector<FavouriteEntry>* target = (type == FavouriteType::Flyout) ? &flyouts_ : &regulars_;
    if (index >= target->size()) {
        return false;
    }

    std::wstring friendly = TrimWhitespace(new_name);
    if (friendly.empty()) {
        friendly = FriendlyNameFromPath((*target)[index].path);
    }
    (*target)[index].friendly_name = std::move(friendly);
    SortAndDedupe(target);
    return Save();
}

const std::vector<FavouriteEntry>& FavouritesStore::GetFlyouts() const noexcept {
    return flyouts_;
}

const std::vector<FavouriteEntry>& FavouritesStore::GetRegulars() const noexcept {
    return regulars_;
}

const std::wstring& FavouritesStore::flyouts_path() const noexcept {
    return flyouts_path_;
}

const std::wstring& FavouritesStore::regulars_path() const noexcept {
    return regulars_path_;
}

bool FavouritesStore::had_existing_files_on_last_load() const noexcept {
    return had_existing_files_on_last_load_;
}

bool FavouritesStore::AddInternal(
    std::vector<FavouriteEntry>* entries,
    const std::wstring& path,
    const std::wstring& friendly_name) {
    if (entries == nullptr) {
        return false;
    }

    const std::wstring normalized_path = NormalizePath(path);
    if (normalized_path.empty()) {
        return false;
    }

    std::wstring resolved_name = TrimWhitespace(friendly_name);
    if (resolved_name.empty()) {
        resolved_name = FriendlyNameFromPath(normalized_path);
    }

    for (FavouriteEntry& existing : *entries) {
        if (PathsEqualInsensitive(existing.path, normalized_path)) {
            if (!resolved_name.empty() && CompareTextInsensitive(existing.friendly_name, resolved_name) != 0) {
                existing.friendly_name = std::move(resolved_name);
                SortAndDedupe(entries);
                return Save();
            }
            return true;
        }
    }

    FavouriteEntry entry = {};
    entry.path = normalized_path;
    entry.friendly_name = std::move(resolved_name);
    entries->push_back(std::move(entry));
    SortAndDedupe(entries);
    return Save();
}

std::vector<FavouriteEntry> FavouritesStore::ParseEntries(const std::vector<std::vector<std::wstring>>& rows) {
    std::vector<FavouriteEntry> entries;
    entries.reserve(rows.size());

    for (const auto& row : rows) {
        if (row.empty()) {
            continue;
        }

        const std::wstring path = NormalizePath(row[0]);
        if (path.empty()) {
            continue;
        }

        std::wstring friendly_name;
        if (row.size() >= 2) {
            friendly_name = TrimWhitespace(row[1]);
        }
        if (friendly_name.empty()) {
            friendly_name = FriendlyNameFromPath(path);
        }

        entries.push_back({path, friendly_name});
    }

    return entries;
}

std::vector<std::vector<std::wstring>> FavouritesStore::SerializeEntries(const std::vector<FavouriteEntry>& entries) {
    std::vector<std::vector<std::wstring>> rows;
    rows.reserve(entries.size());
    for (const FavouriteEntry& entry : entries) {
        rows.push_back({entry.path, entry.friendly_name});
    }
    return rows;
}

std::wstring FavouritesStore::NormalizePath(std::wstring path) {
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

std::wstring FavouritesStore::TrimWhitespace(std::wstring value) {
    const auto not_space = [](wchar_t ch) { return !iswspace(ch); };
    const auto begin = std::find_if(value.begin(), value.end(), not_space);
    if (begin == value.end()) {
        return L"";
    }
    const auto reverse_begin = std::find_if(value.rbegin(), value.rend(), not_space).base();
    return std::wstring(begin, reverse_begin);
}

std::wstring FavouritesStore::FriendlyNameFromPath(const std::wstring& path) {
    const std::wstring normalized = NormalizePath(path);
    if (normalized.empty()) {
        return L"";
    }
    if (normalized.size() == 3 && normalized[1] == L':') {
        return normalized.substr(0, 2);
    }

    const std::wstring::size_type separator = normalized.find_last_of(L'\\');
    if (separator == std::wstring::npos || separator + 1 >= normalized.size()) {
        return normalized;
    }
    return normalized.substr(separator + 1);
}

int FavouritesStore::CompareTextInsensitive(const std::wstring& left, const std::wstring& right) {
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

bool FavouritesStore::PathsEqualInsensitive(const std::wstring& left, const std::wstring& right) {
    return CompareTextInsensitive(NormalizePath(left), NormalizePath(right)) == 0;
}

void FavouritesStore::SortAndDedupe(std::vector<FavouriteEntry>* entries) {
    if (entries == nullptr) {
        return;
    }

    std::sort(
        entries->begin(),
        entries->end(),
        [](const FavouriteEntry& left, const FavouriteEntry& right) {
            const int by_name = CompareTextInsensitive(left.friendly_name, right.friendly_name);
            if (by_name != 0) {
                return by_name < 0;
            }
            return CompareTextInsensitive(left.path, right.path) < 0;
        });

    std::vector<FavouriteEntry> deduped;
    deduped.reserve(entries->size());
    for (const FavouriteEntry& entry : *entries) {
        bool duplicate = false;
        for (const FavouriteEntry& kept : deduped) {
            if (PathsEqualInsensitive(kept.path, entry.path)) {
                duplicate = true;
                break;
            }
        }
        if (!duplicate) {
            deduped.push_back(entry);
        }
    }
    *entries = std::move(deduped);
}

std::wstring FavouritesStore::JoinPath(const std::wstring& base, const std::wstring& child) {
    std::wstring result = NormalizePath(base);
    if (!result.empty() && result.back() != L'\\') {
        result.push_back(L'\\');
    }
    result.append(child);
    return NormalizePath(std::move(result));
}

std::wstring FavouritesStore::ExecutableDirectory() {
    wchar_t path[MAX_PATH] = {};
    const DWORD length = GetModuleFileNameW(nullptr, path, static_cast<DWORD>(std::size(path)));
    if (length == 0 || length >= std::size(path)) {
        LogLastError(L"GetModuleFileNameW(FavouritesStore)");
        return L"";
    }

    std::wstring full_path(path, path + length);
    const std::wstring::size_type separator = full_path.find_last_of(L'\\');
    if (separator == std::wstring::npos) {
        return L"";
    }
    return full_path.substr(0, separator);
}

bool FavouritesStore::IsDirectoryWritable(const std::wstring& directory) {
    if (directory.empty()) {
        return false;
    }

    std::wstring probe = JoinPath(directory, L".fe_write_probe.tmp");
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
        LogLastError(L"CloseHandle(FavouritesStore Probe)");
    }
    DeleteFileW(probe.c_str());
    return true;
}

std::wstring FavouritesStore::ResolveDefaultStorageRoot() {
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
        LogLastError(L"SHGetFolderPathW(FavouritesStore)");
        return exe_directory;
    }

    std::wstring root = JoinPath(appdata, L"FileExplorer");
    if (!EnsureDirectoryExists(root)) {
        return exe_directory.empty() ? root : exe_directory;
    }
    return root;
}

}  // namespace fileexplorer
