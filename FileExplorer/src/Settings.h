#pragma once

#include <array>
#include <string>

namespace fileexplorer {

class Settings final {
public:
    static constexpr int kFileListColumnCount = 5;
    using FileListColumnWidthsLogical = std::array<int, kFileListColumnCount>;

    struct Values {
        int sidebar_width_logical = 320;
        FileListColumnWidthsLogical file_list_column_widths_logical{{260, 60, 140, 90, 320}};
    };

    Settings();

    void SetStoragePath(std::wstring storage_path);
    void UseDefaultStoragePath();

    bool Load(Values* values) const;
    bool Save(const Values& values) const;

    const std::wstring& storage_path() const noexcept;

    static int ClampSidebarWidthLogical(int value) noexcept;
    static int ClampFileListColumnWidthLogical(int column_index, int value) noexcept;

private:
    static std::wstring ResolveDefaultStorageRoot();
    static std::wstring ResolveDefaultStoragePath();
    static std::wstring ExecutableDirectory();
    static std::wstring ParentDirectory(const std::wstring& path);
    static std::wstring JoinPath(const std::wstring& base, const std::wstring& child);
    static bool EnsureDirectoryExists(const std::wstring& path);
    static bool IsDirectoryWritable(const std::wstring& directory);

    std::wstring storage_path_{};
};

}  // namespace fileexplorer
