#include "SearchWorker.h"

#include <Shellapi.h>

#include <algorithm>
#include <cwchar>
#include <cwctype>
#include <thread>
#include <vector>

#include "WildcardMatch.h"
#include "WorkerMessages.h"

namespace {

constexpr size_t kSearchBatchSize = 50;

void LogLastError(const wchar_t* context) {
    const DWORD error_code = GetLastError();
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (GetLastError=%lu).\r\n", context, error_code);
    OutputDebugStringW(buffer);
}

struct FindHandle final {
    HANDLE value = INVALID_HANDLE_VALUE;

    FindHandle() = default;
    explicit FindHandle(HANDLE handle) : value(handle) {}

    FindHandle(const FindHandle&) = delete;
    FindHandle& operator=(const FindHandle&) = delete;

    FindHandle(FindHandle&& other) noexcept : value(other.value) {
        other.value = INVALID_HANDLE_VALUE;
    }

    FindHandle& operator=(FindHandle&& other) noexcept {
        if (this == &other) {
            return *this;
        }
        if (value != INVALID_HANDLE_VALUE) {
            FindClose(value);
        }
        value = other.value;
        other.value = INVALID_HANDLE_VALUE;
        return *this;
    }

    ~FindHandle() {
        if (value != INVALID_HANDLE_VALUE) {
            FindClose(value);
        }
    }

    bool valid() const noexcept {
        return value != INVALID_HANDLE_VALUE;
    }
};

std::wstring ToLowerCopy(std::wstring text) {
    std::transform(
        text.begin(),
        text.end(),
        text.begin(),
        [](wchar_t ch) { return static_cast<wchar_t>(towlower(ch)); });
    return text;
}

}  // namespace

namespace fileexplorer {

void SearchWorker::Start(Request request) {
    std::thread worker_thread(&SearchWorker::Run, std::move(request));
    worker_thread.detach();
}

void SearchWorker::Run(Request request) {
    request.root_path = NormalizePath(std::move(request.root_path));
    request.pattern = request.pattern.empty() ? L"*" : request.pattern;

    if (request.hwnd_target == nullptr || request.root_path.empty()) {
        return;
    }
    if (IsCancelled(request.generation_source, request.generation, request.cancel_token)) {
        return;
    }

    std::vector<std::wstring> directory_stack;
    directory_stack.push_back(request.root_path);

    auto batch = std::make_unique<std::vector<FileEntry>>();
    batch->reserve(kSearchBatchSize);

    while (!directory_stack.empty()) {
        if (IsCancelled(request.generation_source, request.generation, request.cancel_token)) {
            return;
        }

        std::wstring current_dir = std::move(directory_stack.back());
        directory_stack.pop_back();

        std::wstring wildcard = current_dir;
        if (wildcard.back() != L'\\') {
            wildcard.push_back(L'\\');
        }
        wildcard.push_back(L'*');

        WIN32_FIND_DATAW find_data = {};
        FindHandle find_handle(FindFirstFileW(wildcard.c_str(), &find_data));
        if (!find_handle.valid()) {
            continue;
        }

        do {
            if (IsCancelled(request.generation_source, request.generation, request.cancel_token)) {
                return;
            }

            const wchar_t* name = find_data.cFileName;
            if (name[0] == L'.' && (name[1] == L'\0' || (name[1] == L'.' && name[2] == L'\0'))) {
                continue;
            }

            const bool is_hidden = (find_data.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN) != 0;
            const bool is_system = (find_data.dwFileAttributes & FILE_ATTRIBUTE_SYSTEM) != 0;
            if ((!request.show_hidden_files && is_hidden) || is_system) {
                continue;
            }

            const bool is_folder = (find_data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
            const std::wstring full_path = BuildFullPath(current_dir, name);
            if (is_folder) {
                directory_stack.push_back(full_path);
            }

            if (is_folder || !WildcardMatch(request.pattern.c_str(), name)) {
                continue;
            }

            FileEntry entry = {};
            entry.name = name;
            entry.attributes = find_data.dwFileAttributes;
            entry.is_folder = false;
            entry.modified_time = find_data.ftLastWriteTime;
            entry.size_bytes =
                (static_cast<ULONGLONG>(find_data.nFileSizeHigh) << 32ULL) |
                static_cast<ULONGLONG>(find_data.nFileSizeLow);
            entry.extension = GetExtensionFromName(entry.name);
            entry.full_path = full_path;
            entry.icon_index = ResolveIconIndex(full_path, entry.is_folder);
            batch->push_back(std::move(entry));

            if (batch->size() >= kSearchBatchSize) {
                if (!FlushBatch(request.hwnd_target, request.generation, batch.get())) {
                    return;
                }
                batch = std::make_unique<std::vector<FileEntry>>();
                batch->reserve(kSearchBatchSize);
            }
        } while (FindNextFileW(find_handle.value, &find_data) != 0);

        const DWORD find_error = GetLastError();
        if (find_error != ERROR_NO_MORE_FILES) {
            LogLastError(L"FindNextFileW(SearchWorker)");
        }
    }

    if (IsCancelled(request.generation_source, request.generation, request.cancel_token)) {
        return;
    }

    if (batch != nullptr && !batch->empty()) {
        if (!FlushBatch(request.hwnd_target, request.generation, batch.get())) {
            return;
        }
    }

    if (!PostMessageW(
            request.hwnd_target,
            WM_FE_SEARCH_DONE,
            static_cast<WPARAM>(request.generation),
            0)) {
        LogLastError(L"PostMessageW(WM_FE_SEARCH_DONE)");
    }
}

bool SearchWorker::IsCancelled(
    const std::shared_ptr<const std::atomic<uint64_t>>& generation_source,
    uint64_t expected_generation,
    const std::shared_ptr<const std::atomic<bool>>& cancel_token) {
    if (!generation_source || generation_source->load(std::memory_order_acquire) != expected_generation) {
        return true;
    }
    return cancel_token && cancel_token->load(std::memory_order_acquire);
}

bool SearchWorker::FlushBatch(HWND hwnd_target, uint64_t generation, std::vector<FileEntry>* batch) {
    if (hwnd_target == nullptr || batch == nullptr || batch->empty()) {
        return true;
    }

    auto payload = std::make_unique<std::vector<FileEntry>>();
    payload->swap(*batch);
    if (!PostMessageW(
            hwnd_target,
            WM_FE_SEARCH_RESULT,
            static_cast<WPARAM>(generation),
            reinterpret_cast<LPARAM>(payload.get()))) {
        LogLastError(L"PostMessageW(WM_FE_SEARCH_RESULT)");
        return false;
    }

    payload.release();
    return true;
}

std::wstring SearchWorker::NormalizePath(std::wstring path) {
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

std::wstring SearchWorker::BuildFullPath(const std::wstring& parent, const std::wstring& name) {
    if (parent.empty()) {
        return name;
    }

    std::wstring result = parent;
    if (result.back() != L'\\') {
        result.push_back(L'\\');
    }
    result.append(name);
    return result;
}

std::wstring SearchWorker::GetExtensionFromName(const std::wstring& name) {
    const std::wstring::size_type dot = name.find_last_of(L'.');
    if (dot == std::wstring::npos || dot == 0 || dot + 1 >= name.size()) {
        return L"";
    }
    return ToLowerCopy(name.substr(dot));
}

int SearchWorker::ResolveIconIndex(const std::wstring& full_path, bool is_folder) {
    SHFILEINFOW shell_info = {};
    const DWORD attributes = is_folder ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL;
    if (SHGetFileInfoW(
            full_path.c_str(),
            attributes,
            &shell_info,
            sizeof(shell_info),
            SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES) == 0) {
        return 0;
    }
    return shell_info.iIcon;
}

}  // namespace fileexplorer
