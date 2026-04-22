#include <gtest/gtest.h>

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>

#include <string>

#include "CsvParser.h"

namespace {

std::wstring MakeTempCsvPath(const wchar_t* prefix) {
    wchar_t temp_path[MAX_PATH] = {};
    if (GetTempPathW(static_cast<DWORD>(_countof(temp_path)), temp_path) == 0U) {
        return L"";
    }

    std::wstring file_path = temp_path;
    file_path.append(prefix);
    file_path.append(std::to_wstring(GetCurrentProcessId()));
    file_path.push_back(L'_');
    file_path.append(std::to_wstring(GetTickCount64()));
    file_path.append(L".csv");
    return file_path;
}

}  // namespace

namespace fileexplorer::tests {

TEST(Csv, ParseSimpleRow) {
    const CsvParser::CsvRows rows = CsvParser::Parse(L"alpha,beta,gamma");
    EXPECT_TRUE(rows.size() == 1U);
    EXPECT_TRUE(rows[0].size() == 3U);
    EXPECT_TRUE(rows[0][0] == L"alpha");
    EXPECT_TRUE(rows[0][1] == L"beta");
    EXPECT_TRUE(rows[0][2] == L"gamma");
}

TEST(Csv, ParseQuotedFields) {
    const CsvParser::CsvRows rows = CsvParser::Parse(L"\"alpha\",\"beta\"");
    EXPECT_TRUE(rows.size() == 1U);
    EXPECT_TRUE(rows[0].size() == 2U);
    EXPECT_TRUE(rows[0][0] == L"alpha");
    EXPECT_TRUE(rows[0][1] == L"beta");
}

TEST(Csv, ParseCommaInQuotes) {
    const CsvParser::CsvRows rows = CsvParser::Parse(L"\"path,with,comma\",value");
    EXPECT_TRUE(rows.size() == 1U);
    EXPECT_TRUE(rows[0].size() == 2U);
    EXPECT_TRUE(rows[0][0] == L"path,with,comma");
    EXPECT_TRUE(rows[0][1] == L"value");
}

TEST(Csv, ParseNewlineInQuotes) {
    const CsvParser::CsvRows rows = CsvParser::Parse(L"\"line1\r\nline2\",ok");
    EXPECT_TRUE(rows.size() == 1U);
    EXPECT_TRUE(rows[0].size() == 2U);
    EXPECT_TRUE(rows[0][0] == L"line1\r\nline2");
    EXPECT_TRUE(rows[0][1] == L"ok");
}

TEST(Csv, ParseEmptyField) {
    const CsvParser::CsvRows rows = CsvParser::Parse(L"left,,right");
    EXPECT_TRUE(rows.size() == 1U);
    EXPECT_TRUE(rows[0].size() == 3U);
    EXPECT_TRUE(rows[0][0] == L"left");
    EXPECT_TRUE(rows[0][1].empty());
    EXPECT_TRUE(rows[0][2] == L"right");
}

TEST(Csv, ParseMultipleRows) {
    const CsvParser::CsvRows rows = CsvParser::Parse(L"a,b\r\nc,d\r\ne,f");
    EXPECT_TRUE(rows.size() == 3U);
    EXPECT_TRUE(rows[0].size() == 2U && rows[0][0] == L"a" && rows[0][1] == L"b");
    EXPECT_TRUE(rows[1].size() == 2U && rows[1][0] == L"c" && rows[1][1] == L"d");
    EXPECT_TRUE(rows[2].size() == 2U && rows[2][0] == L"e" && rows[2][1] == L"f");
}

TEST(Csv, WriteAndReadRoundtrip) {
    const std::wstring file_path = MakeTempCsvPath(L"FileExplorer_Csv_Roundtrip_");
    if (file_path.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    const CsvParser::CsvRows written = {
        {L"C:\\Root", L"Friendly"},
        {L"D:\\Work\\Item.txt", L"Item"},
    };
    if (!CsvParser::WriteFile(file_path, written)) {
        EXPECT_TRUE(false);
        return;
    }

    CsvParser::CsvRows read_back;
    bool file_exists = false;
    if (!CsvParser::ReadFile(file_path, &read_back, &file_exists)) {
        EXPECT_TRUE(false);
        DeleteFileW(file_path.c_str());
        return;
    }

    EXPECT_TRUE(file_exists);
    EXPECT_TRUE(read_back.size() == written.size());
    EXPECT_TRUE(read_back[0].size() == 2U && read_back[0][0] == L"C:\\Root" && read_back[0][1] == L"Friendly");
    EXPECT_TRUE(
        read_back[1].size() == 2U &&
        read_back[1][0] == L"D:\\Work\\Item.txt" &&
        read_back[1][1] == L"Item");

    DeleteFileW(file_path.c_str());
}

TEST(Csv, WriteSpecialChars) {
    const std::wstring file_path = MakeTempCsvPath(L"FileExplorer_Csv_Special_");
    if (file_path.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    const CsvParser::CsvRows written = {
        {L"C:\\Users\\Árvíztűrő\\Δelta", L"Hello, 世界"},
        {L"C:\\Temp\\\"Quoted\".txt", L"line1\r\nline2"},
    };
    if (!CsvParser::WriteFile(file_path, written)) {
        EXPECT_TRUE(false);
        return;
    }

    CsvParser::CsvRows read_back;
    bool file_exists = false;
    if (!CsvParser::ReadFile(file_path, &read_back, &file_exists)) {
        EXPECT_TRUE(false);
        DeleteFileW(file_path.c_str());
        return;
    }

    EXPECT_TRUE(file_exists);
    EXPECT_TRUE(read_back.size() == 2U);
    EXPECT_TRUE(read_back[0].size() == 2U && read_back[0][0] == written[0][0] && read_back[0][1] == written[0][1]);
    EXPECT_TRUE(read_back[1].size() == 2U && read_back[1][0] == written[1][0] && read_back[1][1] == written[1][1]);

    DeleteFileW(file_path.c_str());
}

TEST(Csv, ParseCRLF) {
    const CsvParser::CsvRows rows = CsvParser::Parse(L"a,b\r\nc,d\r\n");
    EXPECT_TRUE(rows.size() == 2U);
    EXPECT_TRUE(rows[0].size() == 2U && rows[0][0] == L"a" && rows[0][1] == L"b");
    EXPECT_TRUE(rows[1].size() == 2U && rows[1][0] == L"c" && rows[1][1] == L"d");
}

TEST(Csv, ParseLFOnly) {
    const CsvParser::CsvRows rows = CsvParser::Parse(L"a,b\nc,d\n");
    EXPECT_TRUE(rows.size() == 2U);
    EXPECT_TRUE(rows[0].size() == 2U && rows[0][0] == L"a" && rows[0][1] == L"b");
    EXPECT_TRUE(rows[1].size() == 2U && rows[1][0] == L"c" && rows[1][1] == L"d");
}

TEST(Csv, MissingFile) {
    const std::wstring file_path = MakeTempCsvPath(L"FileExplorer_Csv_Missing_");
    if (file_path.empty()) {
        EXPECT_TRUE(false);
        return;
    }

    DeleteFileW(file_path.c_str());

    CsvParser::CsvRows rows;
    bool file_exists = true;
    EXPECT_TRUE(CsvParser::ReadFile(file_path, &rows, &file_exists));
    EXPECT_FALSE(file_exists);
    EXPECT_TRUE(rows.empty());
}

TEST(Csv, CorruptLine) {
    const CsvParser::CsvRows rows = CsvParser::Parse(L"one,\"unterminated\r\ntwo,three");
    EXPECT_FALSE(rows.empty());
    EXPECT_TRUE(rows[0].size() >= 2U);
    EXPECT_TRUE(rows[0][0] == L"one");
}

}  // namespace fileexplorer::tests
