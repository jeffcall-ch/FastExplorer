#include <gtest/gtest.h>

#include <algorithm>
#include <string>

namespace {

std::wstring NormalizeSeparators(std::wstring path) {
    std::replace(path.begin(), path.end(), L'/', L'\\');
    return path;
}

std::wstring ExtractExtension(const std::wstring& path) {
    const std::wstring normalized = NormalizeSeparators(path);
    const std::wstring::size_type dot = normalized.find_last_of(L'.');
    const std::wstring::size_type slash = normalized.find_last_of(L'\\');
    if (dot == std::wstring::npos || (slash != std::wstring::npos && dot < slash)) {
        return L"";
    }
    return normalized.substr(dot);
}

std::wstring ExtractFolderName(const std::wstring& path) {
    std::wstring normalized = NormalizeSeparators(path);
    while (normalized.size() > 3 && !normalized.empty() && normalized.back() == L'\\') {
        normalized.pop_back();
    }
    const std::wstring::size_type slash = normalized.find_last_of(L'\\');
    if (slash == std::wstring::npos) {
        return normalized;
    }
    if (slash + 1 >= normalized.size()) {
        return normalized;
    }
    return normalized.substr(slash + 1);
}

bool IsRoot(const std::wstring& path) {
    const std::wstring normalized = NormalizeSeparators(path);
    return normalized.size() == 3 &&
        normalized[1] == L':' &&
        normalized[2] == L'\\';
}

std::wstring ParentPath(const std::wstring& path) {
    std::wstring normalized = NormalizeSeparators(path);
    while (normalized.size() > 3 && !normalized.empty() && normalized.back() == L'\\') {
        normalized.pop_back();
    }
    if (IsRoot(normalized)) {
        return L"";
    }

    const std::wstring::size_type slash = normalized.find_last_of(L'\\');
    if (slash == std::wstring::npos) {
        return L"";
    }
    if (slash == 2 && normalized[1] == L':') {
        return normalized.substr(0, 3);
    }
    if (slash == 0) {
        return L"";
    }
    return normalized.substr(0, slash);
}

}  // namespace

namespace fileexplorer::tests {

TEST(Path, ExtractExtension) {
    EXPECT_TRUE(ExtractExtension(L"C:\\Work\\report.final.pdf") == L".pdf");
    EXPECT_TRUE(ExtractExtension(L"C:\\Work\\README") == L"");
}

TEST(Path, ExtractFolderName) {
    EXPECT_TRUE(ExtractFolderName(L"C:\\Users\\Administrator\\Downloads") == L"Downloads");
    const std::wstring root_name = ExtractFolderName(L"C:\\");
    EXPECT_TRUE(root_name == L"C:\\" || root_name == L"C:");
}

TEST(Path, IsRoot) {
    EXPECT_TRUE(IsRoot(L"C:\\"));
    EXPECT_FALSE(IsRoot(L"C:\\Users"));
}

TEST(Path, ParentPath) {
    EXPECT_TRUE(ParentPath(L"C:\\Users\\Administrator\\Downloads") == L"C:\\Users\\Administrator");
    EXPECT_TRUE(ParentPath(L"C:\\Users") == L"C:\\");
}

TEST(Path, NormalizeSeparators) {
    EXPECT_TRUE(NormalizeSeparators(L"C:/Users/Administrator/Documents") == L"C:\\Users\\Administrator\\Documents");
}

TEST(Path, LongPath) {
    std::wstring long_folder = L"C:\\";
    for (int i = 0; i < 70; ++i) {
        long_folder.append(L"segment");
        long_folder.append(std::to_wstring(i));
        long_folder.push_back(L'\\');
    }
    long_folder.append(L"file.txt");

    const std::wstring normalized = NormalizeSeparators(long_folder);
    EXPECT_TRUE(normalized.size() > 260U);
    EXPECT_TRUE(ExtractExtension(normalized) == L".txt");
    EXPECT_FALSE(ExtractFolderName(ParentPath(normalized)).empty());
}

}  // namespace fileexplorer::tests
