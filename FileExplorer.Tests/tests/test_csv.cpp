#include <gtest/gtest.h>

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>

#include <iterator>
#include <string>

#include "CsvParser.h"

namespace fileexplorer::tests {

TEST(CsvParser, ParseQuotedCommaFields) {
    const auto rows = CsvParser::Parse(L"\"path,with,comma\",\"Friendly Name\"");
    if (!(rows.size() == 1U && rows[0].size() == 2U)) {
        EXPECT_TRUE(false);
        return;
    }
    EXPECT_TRUE(rows[0][0] == L"path,with,comma");
    EXPECT_TRUE(rows[0][1] == L"Friendly Name");
}

TEST(CsvParser, ParseQuotedNewlineField) {
    const auto rows = CsvParser::Parse(L"\"C:\\\\Root\",\"line1\r\nline2\"");
    if (!(rows.size() == 1U && rows[0].size() == 2U)) {
        EXPECT_TRUE(false);
        return;
    }
    EXPECT_TRUE(rows[0][0] == L"C:\\\\Root");
    EXPECT_TRUE(rows[0][1] == L"line1\r\nline2");
}

TEST(CsvParser, ParseTrimsUnquotedWhitespace) {
    const auto rows = CsvParser::Parse(L"   C:\\Temp   ,   Friendly   \r\n");
    if (!(rows.size() == 1U && rows[0].size() == 2U)) {
        EXPECT_TRUE(false);
        return;
    }
    EXPECT_TRUE(rows[0][0] == L"C:\\Temp");
    EXPECT_TRUE(rows[0][1] == L"Friendly");
}

TEST(CsvParser, MissingFileReturnsEmptyRows) {
    wchar_t temp_path[MAX_PATH] = {};
    if (GetTempPathW(static_cast<DWORD>(std::size(temp_path)), temp_path) == 0U) {
        EXPECT_TRUE(false);
        return;
    }

    std::wstring missing_file = temp_path;
    missing_file.append(L"FileExplorer_Csv_Missing_");
    missing_file.append(std::to_wstring(GetTickCount64()));
    missing_file.append(L".csv");

    CsvParser::CsvRows rows;
    bool file_exists = true;
    EXPECT_TRUE(CsvParser::ReadFile(missing_file, &rows, &file_exists));
    EXPECT_FALSE(file_exists);
    EXPECT_TRUE(rows.empty());
}

TEST(CsvParser, WriteReadRoundTripWithSpecialCharacters) {
    wchar_t temp_path[MAX_PATH] = {};
    if (GetTempPathW(static_cast<DWORD>(std::size(temp_path)), temp_path) == 0U) {
        EXPECT_TRUE(false);
        return;
    }

    std::wstring file_path = temp_path;
    file_path.append(L"FileExplorer_Csv_Roundtrip_");
    file_path.append(std::to_wstring(GetTickCount64()));
    file_path.append(L".csv");

    CsvParser::CsvRows written = {
        {L"C:\\Projects,Archive", L"Alpha, One"},
        {L"C:\\Users\\Name\\Docs", L"Quoted \"Name\""},
        {L"C:\\Users\\Name\\Line", L"First line\r\nSecond line"},
    };

    if (!CsvParser::WriteFile(file_path, written)) {
        EXPECT_TRUE(false);
        return;
    }

    CsvParser::CsvRows read_back;
    bool file_exists = false;
    if (!CsvParser::ReadFile(file_path, &read_back, &file_exists)) {
        EXPECT_TRUE(false);
        return;
    }
    EXPECT_TRUE(file_exists);
    if (read_back.size() != written.size()) {
        EXPECT_TRUE(false);
        DeleteFileW(file_path.c_str());
        return;
    }
    for (size_t i = 0; i < written.size(); ++i) {
        if (read_back[i].size() != written[i].size()) {
            EXPECT_TRUE(false);
            DeleteFileW(file_path.c_str());
            return;
        }
        for (size_t j = 0; j < written[i].size(); ++j) {
            EXPECT_TRUE(read_back[i][j] == written[i][j]);
        }
    }

    DeleteFileW(file_path.c_str());
}

}  // namespace fileexplorer::tests
