#pragma once

#include <string>
#include <vector>

#include "FileListView.h"

namespace fileexplorer {

class SortSettings final {
public:
    struct Entry {
        std::wstring folder_path;
        SortColumn sort_column = SortColumn::Name;
        SortDirection sort_direction = SortDirection::Ascending;
    };

    SortSettings();

    void SetStoragePath(std::wstring storage_path);
    void UseDefaultStoragePath();

    bool Load();
    bool Save() const;

    void Set(const std::wstring& folder_path, SortColumn sort_column, SortDirection sort_direction);
    Entry Get(const std::wstring& folder_path) const;

    const std::wstring& storage_path() const noexcept;

    static constexpr size_t kMaxEntries = 500;

private:
    static std::wstring ResolveDefaultStorageRoot();
    static std::wstring ResolveDefaultStoragePath();
    static std::wstring ExecutableDirectory();
    static std::wstring ParentDirectory(const std::wstring& path);
    static std::wstring JoinPath(const std::wstring& base, const std::wstring& child);
    static bool EnsureDirectoryExists(const std::wstring& path);
    static bool IsDirectoryWritable(const std::wstring& directory);

    static std::wstring NormalizePath(std::wstring path);
    static int ComparePathInsensitive(const std::wstring& left, const std::wstring& right);
    static bool TryParseColumn(const std::wstring& value, SortColumn* sort_column);
    static bool TryParseDirection(const std::wstring& value, SortDirection* sort_direction);
    static std::wstring ColumnToString(SortColumn sort_column);
    static std::wstring DirectionToString(SortDirection sort_direction);

    std::wstring storage_path_{};
    std::vector<Entry> entries_{};
};

}  // namespace fileexplorer
