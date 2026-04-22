#include <gtest/gtest.h>

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>

#include <algorithm>
#include <string>

#include "FavouritesStore.h"

namespace {

std::wstring MakeTempCsvPath(const wchar_t* prefix) {
    wchar_t temp_path[MAX_PATH] = {};
    if (GetTempPathW(static_cast<DWORD>(_countof(temp_path)), temp_path) == 0U) {
        return L"";
    }

    std::wstring path = temp_path;
    path.append(prefix);
    path.append(std::to_wstring(GetCurrentProcessId()));
    path.push_back(L'_');
    path.append(std::to_wstring(GetTickCount64()));
    path.append(L".csv");
    return path;
}

int FindIndexByPath(const std::vector<fileexplorer::FavouriteEntry>& entries, const std::wstring& path) {
    for (int i = 0; i < static_cast<int>(entries.size()); ++i) {
        if (CompareStringOrdinal(
                entries[static_cast<size_t>(i)].path.c_str(),
                -1,
                path.c_str(),
                -1,
                TRUE) == CSTR_EQUAL) {
            return i;
        }
    }
    return -1;
}

}  // namespace

namespace fileexplorer::tests {

TEST(Favstore, AddAndLoadFlyout) {
    const std::wstring flyout_file = MakeTempCsvPath(L"FileExplorer_Fav_Flyout_");
    const std::wstring regular_file = MakeTempCsvPath(L"FileExplorer_Fav_Regular_");
    if (flyout_file.empty() || regular_file.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    DeleteFileW(flyout_file.c_str());
    DeleteFileW(regular_file.c_str());

    FavouritesStore writer;
    writer.SetStoragePaths(flyout_file, regular_file);
    EXPECT_TRUE(writer.AddFlyout(L"C:\\Users\\Administrator\\Desktop", L"Desktop"));

    FavouritesStore reader;
    reader.SetStoragePaths(flyout_file, regular_file);
    EXPECT_TRUE(reader.Load());
    const auto& flyouts = reader.GetFlyouts();
    EXPECT_TRUE(flyouts.size() == 1U);
    EXPECT_TRUE(flyouts[0].path == L"C:\\Users\\Administrator\\Desktop");
    EXPECT_TRUE(flyouts[0].friendly_name == L"Desktop");

    DeleteFileW(flyout_file.c_str());
    DeleteFileW(regular_file.c_str());
}

TEST(Favstore, AddAndLoadRegular) {
    const std::wstring flyout_file = MakeTempCsvPath(L"FileExplorer_Fav2_Flyout_");
    const std::wstring regular_file = MakeTempCsvPath(L"FileExplorer_Fav2_Regular_");
    if (flyout_file.empty() || regular_file.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    DeleteFileW(flyout_file.c_str());
    DeleteFileW(regular_file.c_str());

    FavouritesStore writer;
    writer.SetStoragePaths(flyout_file, regular_file);
    EXPECT_TRUE(writer.AddRegular(L"D:\\Work\\Projects", L"Projects"));

    FavouritesStore reader;
    reader.SetStoragePaths(flyout_file, regular_file);
    EXPECT_TRUE(reader.Load());
    const auto& regulars = reader.GetRegulars();
    EXPECT_TRUE(regulars.size() == 1U);
    EXPECT_TRUE(regulars[0].path == L"D:\\Work\\Projects");
    EXPECT_TRUE(regulars[0].friendly_name == L"Projects");

    DeleteFileW(flyout_file.c_str());
    DeleteFileW(regular_file.c_str());
}

TEST(Favstore, RemoveFlyout) {
    const std::wstring flyout_file = MakeTempCsvPath(L"FileExplorer_Fav3_Flyout_");
    const std::wstring regular_file = MakeTempCsvPath(L"FileExplorer_Fav3_Regular_");
    if (flyout_file.empty() || regular_file.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    DeleteFileW(flyout_file.c_str());
    DeleteFileW(regular_file.c_str());

    FavouritesStore store;
    store.SetStoragePaths(flyout_file, regular_file);
    EXPECT_TRUE(store.AddFlyout(L"C:\\Users\\Administrator\\Desktop", L"Desktop"));
    EXPECT_TRUE(store.AddFlyout(L"C:\\Users\\Administrator\\Downloads", L"Downloads"));

    const int index = FindIndexByPath(store.GetFlyouts(), L"C:\\Users\\Administrator\\Desktop");
    EXPECT_TRUE(index >= 0);
    EXPECT_TRUE(store.Remove(FavouriteType::Flyout, static_cast<size_t>(index)));
    EXPECT_TRUE(store.GetFlyouts().size() == 1U);
    EXPECT_TRUE(store.GetFlyouts()[0].path == L"C:\\Users\\Administrator\\Downloads");

    DeleteFileW(flyout_file.c_str());
    DeleteFileW(regular_file.c_str());
}

TEST(Favstore, RenameRegular) {
    const std::wstring flyout_file = MakeTempCsvPath(L"FileExplorer_Fav4_Flyout_");
    const std::wstring regular_file = MakeTempCsvPath(L"FileExplorer_Fav4_Regular_");
    if (flyout_file.empty() || regular_file.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    DeleteFileW(flyout_file.c_str());
    DeleteFileW(regular_file.c_str());

    FavouritesStore store;
    store.SetStoragePaths(flyout_file, regular_file);
    EXPECT_TRUE(store.AddRegular(L"D:\\Docs", L"My Docs"));

    const int index = FindIndexByPath(store.GetRegulars(), L"D:\\Docs");
    EXPECT_TRUE(index >= 0);
    EXPECT_TRUE(store.Rename(FavouriteType::Regular, static_cast<size_t>(index), L"Work Docs"));

    const int renamed_index = FindIndexByPath(store.GetRegulars(), L"D:\\Docs");
    EXPECT_TRUE(renamed_index >= 0);
    EXPECT_TRUE(store.GetRegulars()[static_cast<size_t>(renamed_index)].friendly_name == L"Work Docs");

    DeleteFileW(flyout_file.c_str());
    DeleteFileW(regular_file.c_str());
}

TEST(Favstore, AlphabeticOrder) {
    const std::wstring flyout_file = MakeTempCsvPath(L"FileExplorer_Fav5_Flyout_");
    const std::wstring regular_file = MakeTempCsvPath(L"FileExplorer_Fav5_Regular_");
    if (flyout_file.empty() || regular_file.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    DeleteFileW(flyout_file.c_str());
    DeleteFileW(regular_file.c_str());

    FavouritesStore store;
    store.SetStoragePaths(flyout_file, regular_file);
    EXPECT_TRUE(store.AddRegular(L"C:\\Zeta", L"zeta"));
    EXPECT_TRUE(store.AddRegular(L"C:\\Alpha", L"Alpha"));
    EXPECT_TRUE(store.AddRegular(L"C:\\Beta", L"beta"));

    const auto& regulars = store.GetRegulars();
    EXPECT_TRUE(regulars.size() == 3U);
    EXPECT_TRUE(regulars[0].friendly_name == L"Alpha");
    EXPECT_TRUE(regulars[1].friendly_name == L"beta");
    EXPECT_TRUE(regulars[2].friendly_name == L"zeta");

    DeleteFileW(flyout_file.c_str());
    DeleteFileW(regular_file.c_str());
}

TEST(Favstore, DuplicatePathIgnored) {
    const std::wstring flyout_file = MakeTempCsvPath(L"FileExplorer_Fav6_Flyout_");
    const std::wstring regular_file = MakeTempCsvPath(L"FileExplorer_Fav6_Regular_");
    if (flyout_file.empty() || regular_file.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    DeleteFileW(flyout_file.c_str());
    DeleteFileW(regular_file.c_str());

    FavouritesStore store;
    store.SetStoragePaths(flyout_file, regular_file);
    EXPECT_TRUE(store.AddRegular(L"C:\\Users\\Administrator\\Desktop", L"Desktop"));
    EXPECT_TRUE(store.AddRegular(L"c:\\users\\administrator\\desktop\\", L"Desktop Updated"));

    const auto& regulars = store.GetRegulars();
    EXPECT_TRUE(regulars.size() == 1U);
    EXPECT_TRUE(regulars[0].friendly_name == L"Desktop Updated");

    DeleteFileW(flyout_file.c_str());
    DeleteFileW(regular_file.c_str());
}

TEST(Favstore, SpecialPathChars) {
    const std::wstring flyout_file = MakeTempCsvPath(L"FileExplorer_Fav7_Flyout_");
    const std::wstring regular_file = MakeTempCsvPath(L"FileExplorer_Fav7_Regular_");
    if (flyout_file.empty() || regular_file.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    DeleteFileW(flyout_file.c_str());
    DeleteFileW(regular_file.c_str());

    FavouritesStore store;
    store.SetStoragePaths(flyout_file, regular_file);
    EXPECT_TRUE(store.AddRegular(L"C:\\Users\\Árvíztűrő\\Δelta,Docs", L"Δ Docs, Team"));

    FavouritesStore reader;
    reader.SetStoragePaths(flyout_file, regular_file);
    EXPECT_TRUE(reader.Load());
    const auto& regulars = reader.GetRegulars();
    EXPECT_TRUE(regulars.size() == 1U);
    EXPECT_TRUE(regulars[0].path == L"C:\\Users\\Árvíztűrő\\Δelta,Docs");
    EXPECT_TRUE(regulars[0].friendly_name == L"Δ Docs, Team");

    DeleteFileW(flyout_file.c_str());
    DeleteFileW(regular_file.c_str());
}

TEST(Favstore, PersistAndReload) {
    const std::wstring flyout_file = MakeTempCsvPath(L"FileExplorer_Fav8_Flyout_");
    const std::wstring regular_file = MakeTempCsvPath(L"FileExplorer_Fav8_Regular_");
    if (flyout_file.empty() || regular_file.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    DeleteFileW(flyout_file.c_str());
    DeleteFileW(regular_file.c_str());

    FavouritesStore writer;
    writer.SetStoragePaths(flyout_file, regular_file);
    EXPECT_TRUE(writer.AddFlyout(L"C:\\Users\\Administrator\\Desktop", L"Desktop"));
    EXPECT_TRUE(writer.AddRegular(L"D:\\Work\\Reports", L"Reports"));

    FavouritesStore reader;
    reader.SetStoragePaths(flyout_file, regular_file);
    EXPECT_TRUE(reader.Load());
    EXPECT_TRUE(reader.had_existing_files_on_last_load());
    EXPECT_TRUE(reader.GetFlyouts().size() == 1U);
    EXPECT_TRUE(reader.GetRegulars().size() == 1U);
    EXPECT_TRUE(reader.GetFlyouts()[0].friendly_name == L"Desktop");
    EXPECT_TRUE(reader.GetRegulars()[0].friendly_name == L"Reports");

    DeleteFileW(flyout_file.c_str());
    DeleteFileW(regular_file.c_str());
}

}  // namespace fileexplorer::tests
