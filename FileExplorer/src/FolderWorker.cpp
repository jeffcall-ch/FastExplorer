#include "FolderWorker.h"

#include <Shellapi.h>

#include <algorithm>
#include <cwchar>
#include <cwctype>
#include <thread>

#include "WorkerMessages.h"

namespace {

void LogLastError(const wchar_t* context) {
    const DWORD error_code = GetLastError();
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (GetLastError=%lu).\r\n", context, error_code);
    OutputDebugStringW(buffer);
}

bool IsCurrentGeneration(
    const std::shared_ptr<const std::atomic<uint64_t>>& generation_source,
    uint64_t expected_generation) {
    return generation_source && generation_source->load(std::memory_order_acquire) == expected_generation;
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

void FolderWorker::Start(Request request) {
    std::thread worker_thread(&FolderWorker::Run, std::move(request));
    worker_thread.detach();
}

void FolderWorker::Run(Request request) {
    request.path = NormalizePath(std::move(request.path));
    if (!IsCurrentGeneration(request.generation_source, request.generation)) {
        return;
    }

    auto entries = std::make_unique<std::vector<FileEntry>>(EnumerateEntries(request.path));
    std::sort(
        entries->begin(),
        entries->end(),
        [](const FileEntry& left, const FileEntry& right) {
            if (left.is_folder != right.is_folder) {
                return left.is_folder;
            }
            return FolderWorker::CompareTextInsensitive(left.name, right.name) < 0;
        });

    if (!IsCurrentGeneration(request.generation_source, request.generation)) {
        return;
    }

    if (!PostMessageW(
            request.hwnd_target,
            WM_FE_FOLDER_LOADED,
            static_cast<WPARAM>(request.generation),
            reinterpret_cast<LPARAM>(entries.get()))) {
        LogLastError(L"PostMessageW(WM_FE_FOLDER_LOADED)");
        return;
    }

    entries.release();
}

std::vector<FileEntry> FolderWorker::EnumerateEntries(const std::wstring& path) {
    std::vector<FileEntry> entries;
    if (path.empty()) {
        return entries;
    }

    std::wstring wildcard = path;
    if (wildcard.back() != L'\\') {
        wildcard.push_back(L'\\');
    }
    wildcard.push_back(L'*');

    WIN32_FIND_DATAW find_data = {};
    FindHandle find_handle(FindFirstFileW(wildcard.c_str(), &find_data));
    if (!find_handle.valid()) {
        return entries;
    }

    do {
        const wchar_t* name = find_data.cFileName;
        if (name[0] == L'.' && (name[1] == L'\0' || (name[1] == L'.' && name[2] == L'\0'))) {
            continue;
        }

        if ((find_data.dwFileAttributes & (FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM)) != 0) {
            continue;
        }

        FileEntry entry = {};
        entry.name = name;
        entry.attributes = find_data.dwFileAttributes;
        entry.is_folder = (find_data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        entry.modified_time = find_data.ftLastWriteTime;
        entry.size_bytes =
            (static_cast<ULONGLONG>(find_data.nFileSizeHigh) << 32ULL) |
            static_cast<ULONGLONG>(find_data.nFileSizeLow);
        if (entry.is_folder) {
            entry.size_bytes = 0;
        } else {
            entry.extension = GetExtensionFromName(entry.name);
        }

        const std::wstring full_path = BuildFullPath(path, entry.name);
        entry.icon_index = ResolveIconIndex(full_path, entry.is_folder);
        entries.push_back(std::move(entry));
    } while (FindNextFileW(find_handle.value, &find_data) != 0);

    const DWORD find_error = GetLastError();
    if (find_error != ERROR_NO_MORE_FILES) {
        LogLastError(L"FindNextFileW(FolderWorker)");
    }

    return entries;
}

std::wstring FolderWorker::NormalizePath(std::wstring path) {
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

std::wstring FolderWorker::BuildFullPath(const std::wstring& parent, const std::wstring& name) {
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

std::wstring FolderWorker::GetExtensionFromName(const std::wstring& name) {
    const std::wstring::size_type dot = name.find_last_of(L'.');
    if (dot == std::wstring::npos || dot == 0 || dot + 1 >= name.size()) {
        return L"";
    }
    return ToLowerCopy(name.substr(dot));
}

int FolderWorker::ResolveIconIndex(const std::wstring& full_path, bool is_folder) {
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

int FolderWorker::CompareTextInsensitive(const std::wstring& left, const std::wstring& right) {
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

}  // namespace fileexplorer
