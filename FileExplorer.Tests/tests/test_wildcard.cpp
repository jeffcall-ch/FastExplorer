#include <gtest/gtest.h>

#include <string>

#include "WildcardMatch.h"

namespace fileexplorer::tests {

TEST(Wildcard, ExactMatch) {
    EXPECT_TRUE(fileexplorer::WildcardMatch(L"report.pdf", L"report.pdf"));
}

TEST(Wildcard, StarMatchesEmpty) {
    EXPECT_TRUE(fileexplorer::WildcardMatch(L"report*", L"report"));
}

TEST(Wildcard, StarMatchesSequence) {
    EXPECT_TRUE(fileexplorer::WildcardMatch(L"rep*", L"report_final"));
}

TEST(Wildcard, QuestionMatchesOne) {
    EXPECT_TRUE(fileexplorer::WildcardMatch(L"file?.txt", L"file1.txt"));
}

TEST(Wildcard, QuestionNoMatchEmpty) {
    EXPECT_FALSE(fileexplorer::WildcardMatch(L"file?.txt", L"file.txt"));
}

TEST(Wildcard, StarAtStart) {
    EXPECT_TRUE(fileexplorer::WildcardMatch(L"*.txt", L"notes.txt"));
}

TEST(Wildcard, StarAtEnd) {
    EXPECT_TRUE(fileexplorer::WildcardMatch(L"log*", L"log_2026_04"));
}

TEST(Wildcard, MultipleStars) {
    EXPECT_TRUE(fileexplorer::WildcardMatch(L"*rep*final*", L"my_report_final_v2.doc"));
}

TEST(Wildcard, CaseInsensitive) {
    EXPECT_TRUE(fileexplorer::WildcardMatch(L"*.PDF", L"REPORT.pdf"));
}

TEST(Wildcard, ExtensionPattern) {
    EXPECT_TRUE(fileexplorer::WildcardMatch(L"*.pdf", L"report.pdf"));
}

TEST(Wildcard, DoubleExtension) {
    EXPECT_TRUE(fileexplorer::WildcardMatch(L"*.tar.gz", L"archive.tar.gz"));
}

TEST(Wildcard, NoMatchDifferentExt) {
    EXPECT_FALSE(fileexplorer::WildcardMatch(L"*.pdf", L"report.docx"));
}

TEST(Wildcard, QuestionInMiddle) {
    EXPECT_TRUE(fileexplorer::WildcardMatch(L"rep?rt.pdf", L"report.pdf"));
}

TEST(Wildcard, EmptyPattern) {
    EXPECT_FALSE(fileexplorer::WildcardMatch(L"", L"anything"));
}

TEST(Wildcard, EmptyString) {
    EXPECT_TRUE(fileexplorer::WildcardMatch(L"*", L""));
}

TEST(Wildcard, LongPath) {
    std::wstring long_name(300, L'a');
    long_name.append(L".txt");
    EXPECT_TRUE(fileexplorer::WildcardMatch(L"*.txt", long_name.c_str()));
}

}  // namespace fileexplorer::tests
