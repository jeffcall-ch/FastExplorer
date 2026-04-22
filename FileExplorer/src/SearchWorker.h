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

class SearchWorker final {
public:
    struct Request {
        std::wstring root_path;
        std::wstring pattern;
        bool show_hidden_files = false;
        uint64_t generation = 0;
        HWND hwnd_target = nullptr;
        std::shared_ptr<const std::atomic<uint64_t>> generation_source;
        std::shared_ptr<const std::atomic<bool>> cancel_token;
    };

    static void Start(Request request);

private:
    static void Run(Request request);
    static bool IsCancelled(
        const std::shared_ptr<const std::atomic<uint64_t>>& generation_source,
        uint64_t expected_generation,
        const std::shared_ptr<const std::atomic<bool>>& cancel_token);
    static bool FlushBatch(HWND hwnd_target, uint64_t generation, std::vector<FileEntry>* batch);
    static std::wstring NormalizePath(std::wstring path);
    static std::wstring BuildFullPath(const std::wstring& parent, const std::wstring& name);
    static std::wstring GetExtensionFromName(const std::wstring& name);
    static int ResolveIconIndex(const std::wstring& full_path, bool is_folder);
};

}  // namespace fileexplorer
