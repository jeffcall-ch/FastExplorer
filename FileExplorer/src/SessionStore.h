#pragma once

#include <string>
#include <vector>

namespace fileexplorer {

class SessionStore final {
public:
    struct SessionTab {
        std::wstring path;
        bool pinned = false;
    };

    SessionStore();

    void SetStoragePath(std::wstring storage_path);
    void UseDefaultStoragePath();

    bool Save(const std::vector<SessionTab>& tabs, int active_index) const;
    std::vector<SessionTab> Load(int* active_index) const;

    const std::wstring& storage_path() const noexcept;

private:
    static std::wstring ResolveDefaultStorageRoot();
    static std::wstring ResolveDefaultStoragePath();
    static std::wstring ExecutableDirectory();
    static std::wstring ParentDirectory(const std::wstring& path);
    static std::wstring JoinPath(const std::wstring& base, const std::wstring& child);
    static bool EnsureDirectoryExists(const std::wstring& path);
    static bool IsDirectoryWritable(const std::wstring& directory);
    static std::wstring NormalizePath(std::wstring path);

    std::wstring storage_path_{};
};

}  // namespace fileexplorer
