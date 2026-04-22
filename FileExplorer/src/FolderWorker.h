#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>

#include <atomic>
#include <memory>
#include <string>
#include <vector>

#include "FileListView.h"

namespace fileexplorer {

class FolderWorker final {
public:
    struct Request {
        std::wstring path;
        bool show_hidden_files = false;
        uint64_t generation = 0;
        HWND hwnd_target = nullptr;
        std::shared_ptr<const std::atomic<uint64_t>> generation_source;
    };

    static void Start(Request request);

private:
    static void Run(Request request);
    static std::vector<FileEntry> EnumerateEntries(const std::wstring& path, bool show_hidden_files);
    static std::wstring NormalizePath(std::wstring path);
    static std::wstring BuildFullPath(const std::wstring& parent, const std::wstring& name);
    static std::wstring GetExtensionFromName(const std::wstring& name);
    static int ResolveIconIndex(const std::wstring& full_path, bool is_folder);
    static int CompareTextInsensitive(const std::wstring& left, const std::wstring& right);
};

}  // namespace fileexplorer
