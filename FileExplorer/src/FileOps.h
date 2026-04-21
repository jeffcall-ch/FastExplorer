#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>

#include <string>
#include <vector>

namespace fileexplorer {

class FileOps final {
public:
    static bool CopyPathsToClipboard(HWND owner, const std::vector<std::wstring>& paths, bool cut);
    static bool ClipboardHasDropFiles();
    static bool PasteFromClipboard(HWND owner, const std::wstring& destination_folder, bool* was_move);

    static bool DeleteToRecycleBin(HWND owner, const std::vector<std::wstring>& paths);
    static bool DeletePermanently(HWND owner, const std::vector<std::wstring>& paths);

    static bool RenamePath(const std::wstring& from_path, const std::wstring& to_path, DWORD* error_code);
    static bool CreateDirectoryPath(const std::wstring& path, DWORD* error_code);

    static bool CopyTextToClipboard(HWND owner, const std::wstring& text);
    static bool ShowProperties(HWND owner, const std::wstring& path);

private:
    static std::wstring BuildMultiStringPaths(const std::vector<std::wstring>& paths);
    static bool PerformShellOperation(
        HWND owner,
        UINT operation,
        const std::vector<std::wstring>& from_paths,
        const std::wstring& to_path,
        FILEOP_FLAGS flags,
        bool* any_aborted);
};

}  // namespace fileexplorer
