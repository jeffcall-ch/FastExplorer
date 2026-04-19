#pragma once

#include <cstddef>
#include <string>
#include <vector>

namespace fileexplorer {

enum class FavouriteType {
    Flyout,
    Regular,
};

struct FavouriteEntry {
    std::wstring path;
    std::wstring friendly_name;
};

class FavouritesStore final {
public:
    FavouritesStore();

    void SetStoragePaths(std::wstring flyouts_path, std::wstring regulars_path);
    void UseDefaultStoragePaths();

    bool Load();
    bool Save() const;

    bool AddFlyout(const std::wstring& path, const std::wstring& friendly_name = L"");
    bool AddRegular(const std::wstring& path, const std::wstring& friendly_name = L"");
    bool Remove(FavouriteType type, size_t index);
    bool Rename(FavouriteType type, size_t index, const std::wstring& new_name);

    const std::vector<FavouriteEntry>& GetFlyouts() const noexcept;
    const std::vector<FavouriteEntry>& GetRegulars() const noexcept;

    const std::wstring& flyouts_path() const noexcept;
    const std::wstring& regulars_path() const noexcept;

    bool had_existing_files_on_last_load() const noexcept;

private:
    bool AddInternal(std::vector<FavouriteEntry>* entries, const std::wstring& path, const std::wstring& friendly_name);
    static std::vector<FavouriteEntry> ParseEntries(const std::vector<std::vector<std::wstring>>& rows);
    static std::vector<std::vector<std::wstring>> SerializeEntries(const std::vector<FavouriteEntry>& entries);

    static std::wstring NormalizePath(std::wstring path);
    static std::wstring TrimWhitespace(std::wstring value);
    static std::wstring FriendlyNameFromPath(const std::wstring& path);
    static int CompareTextInsensitive(const std::wstring& left, const std::wstring& right);
    static bool PathsEqualInsensitive(const std::wstring& left, const std::wstring& right);
    static void SortAndDedupe(std::vector<FavouriteEntry>* entries);

    static std::wstring JoinPath(const std::wstring& base, const std::wstring& child);
    static std::wstring ExecutableDirectory();
    static bool IsDirectoryWritable(const std::wstring& directory);
    static std::wstring ResolveDefaultStorageRoot();

    std::vector<FavouriteEntry> flyouts_{};
    std::vector<FavouriteEntry> regulars_{};
    std::wstring flyouts_path_{};
    std::wstring regulars_path_{};
    bool had_existing_files_on_last_load_{false};
};

}  // namespace fileexplorer
