#include <gtest/gtest.h>

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>

#include <cwchar>
#include <string>
#include <vector>

#include "SessionStore.h"

namespace {

std::wstring MakeTempSessionPath(const wchar_t* prefix) {
    wchar_t temp_path[MAX_PATH] = {};
    if (GetTempPathW(static_cast<DWORD>(_countof(temp_path)), temp_path) == 0U) {
        return L"";
    }

    std::wstring path = temp_path;
    path.append(prefix);
    path.append(std::to_wstring(GetCurrentProcessId()));
    path.push_back(L'_');
    path.append(std::to_wstring(GetTickCount64()));
    path.append(L".ini");
    return path;
}

void WriteTextFile(const std::wstring& path, const wchar_t* content) {
    HANDLE handle = CreateFileW(
        path.c_str(),
        GENERIC_WRITE,
        0,
        nullptr,
        CREATE_ALWAYS,
        FILE_ATTRIBUTE_NORMAL,
        nullptr);
    if (handle == INVALID_HANDLE_VALUE) {
        return;
    }

    const DWORD bytes = static_cast<DWORD>(wcslen(content) * sizeof(wchar_t));
    DWORD written = 0;
    if (bytes > 0) {
        WriteFile(handle, content, bytes, &written, nullptr);
    }
    CloseHandle(handle);
}

}  // namespace

namespace fileexplorer::tests {

TEST(Session, SaveAndLoad) {
    const std::wstring file_path = MakeTempSessionPath(L"FileExplorer_Session_SaveLoad_");
    if (file_path.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    DeleteFileW(file_path.c_str());

    SessionStore writer;
    writer.SetStoragePath(file_path);
    const std::vector<SessionStore::SessionTab> tabs = {
        {L"C:\\Users\\Administrator", false},
        {L"D:\\Work\\Project", true},
        {L"E:\\Archive", false},
    };
    EXPECT_TRUE(writer.Save(tabs, 1));

    SessionStore reader;
    reader.SetStoragePath(file_path);
    int active_index = -1;
    const auto loaded_tabs = reader.Load(&active_index);
    EXPECT_TRUE(loaded_tabs.size() == tabs.size());
    EXPECT_TRUE(active_index == 1);
    EXPECT_TRUE(loaded_tabs[0].path == tabs[0].path && loaded_tabs[0].pinned == tabs[0].pinned);
    EXPECT_TRUE(loaded_tabs[1].path == tabs[1].path && loaded_tabs[1].pinned == tabs[1].pinned);
    EXPECT_TRUE(loaded_tabs[2].path == tabs[2].path && loaded_tabs[2].pinned == tabs[2].pinned);

    DeleteFileW(file_path.c_str());
}

TEST(Session, LoadMissingFile) {
    const std::wstring file_path = MakeTempSessionPath(L"FileExplorer_Session_Missing_");
    if (file_path.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    DeleteFileW(file_path.c_str());

    SessionStore store;
    store.SetStoragePath(file_path);
    int active_index = 7;
    const auto loaded = store.Load(&active_index);
    EXPECT_TRUE(loaded.empty());
    EXPECT_TRUE(active_index == 0);
}

TEST(Session, LoadCorruptFile) {
    const std::wstring file_path = MakeTempSessionPath(L"FileExplorer_Session_Corrupt_");
    if (file_path.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    WriteTextFile(file_path, L"not-an-ini-format\r\n@@@\r\n");

    SessionStore store;
    store.SetStoragePath(file_path);
    int active_index = 3;
    const auto loaded = store.Load(&active_index);
    EXPECT_TRUE(loaded.empty());
    EXPECT_TRUE(active_index == 0);

    DeleteFileW(file_path.c_str());
}

TEST(Session, ActiveTabPreserved) {
    const std::wstring file_path = MakeTempSessionPath(L"FileExplorer_Session_Active_");
    if (file_path.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    DeleteFileW(file_path.c_str());

    SessionStore writer;
    writer.SetStoragePath(file_path);
    const std::vector<SessionStore::SessionTab> tabs = {
        {L"C:\\A", false},
        {L"C:\\B", false},
        {L"C:\\C", false},
    };
    EXPECT_TRUE(writer.Save(tabs, 2));

    SessionStore reader;
    reader.SetStoragePath(file_path);
    int active_index = -1;
    (void)reader.Load(&active_index);
    EXPECT_TRUE(active_index == 2);

    DeleteFileW(file_path.c_str());
}

TEST(Session, PinnedStatePreserved) {
    const std::wstring file_path = MakeTempSessionPath(L"FileExplorer_Session_Pinned_");
    if (file_path.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    DeleteFileW(file_path.c_str());

    SessionStore writer;
    writer.SetStoragePath(file_path);
    const std::vector<SessionStore::SessionTab> tabs = {
        {L"C:\\Pinned", true},
        {L"C:\\Regular", false},
    };
    EXPECT_TRUE(writer.Save(tabs, 0));

    SessionStore reader;
    reader.SetStoragePath(file_path);
    int active_index = 0;
    const auto loaded = reader.Load(&active_index);
    EXPECT_TRUE(loaded.size() == 2U);
    EXPECT_TRUE(loaded[0].pinned);
    EXPECT_FALSE(loaded[1].pinned);

    DeleteFileW(file_path.c_str());
}

TEST(Session, MaxTabsRoundtrip) {
    const std::wstring file_path = MakeTempSessionPath(L"FileExplorer_Session_MaxTabs_");
    if (file_path.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    DeleteFileW(file_path.c_str());

    std::vector<SessionStore::SessionTab> tabs;
    tabs.reserve(20);
    for (int i = 0; i < 20; ++i) {
        tabs.push_back({L"C:\\Tab" + std::to_wstring(i), (i % 3) == 0});
    }

    SessionStore writer;
    writer.SetStoragePath(file_path);
    EXPECT_TRUE(writer.Save(tabs, 19));

    SessionStore reader;
    reader.SetStoragePath(file_path);
    int active_index = 0;
    const auto loaded = reader.Load(&active_index);
    EXPECT_TRUE(loaded.size() == 20U);
    EXPECT_TRUE(active_index == 19);
    EXPECT_TRUE(loaded[0].path == L"C:\\Tab0");
    EXPECT_TRUE(loaded[19].path == L"C:\\Tab19");

    DeleteFileW(file_path.c_str());
}

}  // namespace fileexplorer::tests
