#pragma once

#include <functional>
#include <iostream>
#include <sstream>
#include <string>

namespace testing {

void InitGoogleTest(int* argc, char** argv);
bool RegisterTest(const char* suite_name, const char* test_name, std::function<void()> test_body);
int RunAllRegisteredTests();
void ReportFailure(const char* file, int line, const std::string& message);

inline int RUN_ALL_TESTS() {
    return RunAllRegisteredTests();
}

}  // namespace testing

#define RUN_ALL_TESTS() ::testing::RUN_ALL_TESTS()

#define TEST(test_suite_name, test_name)                                                                          \
    static void test_suite_name##_##test_name##_Body();                                                           \
    namespace {                                                                                                   \
    const bool test_suite_name##_##test_name##_Registered =                                                       \
        ::testing::RegisterTest(#test_suite_name, #test_name, []() { test_suite_name##_##test_name##_Body(); }); \
    }                                                                                                             \
    static void test_suite_name##_##test_name##_Body()

#define EXPECT_TRUE(condition)                                                                       \
    do {                                                                                             \
        if (!(condition)) {                                                                          \
            ::testing::ReportFailure(__FILE__, __LINE__, "EXPECT_TRUE failed: " #condition);        \
        }                                                                                            \
    } while (false)

#define EXPECT_FALSE(condition)                                                                      \
    do {                                                                                             \
        if (condition) {                                                                             \
            ::testing::ReportFailure(__FILE__, __LINE__, "EXPECT_FALSE failed: " #condition);       \
        }                                                                                            \
    } while (false)

#define EXPECT_EQ(expected, actual)                                                                     \
    do {                                                                                                \
        const auto expected_value = (expected);                                                         \
        const auto actual_value = (actual);                                                             \
        if (!(expected_value == actual_value)) {                                                        \
            std::ostringstream stream;                                                                  \
            stream << "EXPECT_EQ failed. Expected: " << expected_value << ", Actual: " << actual_value; \
            ::testing::ReportFailure(__FILE__, __LINE__, stream.str());                                \
        }                                                                                               \
    } while (false)

#define ASSERT_TRUE(condition) EXPECT_TRUE(condition)
#define ASSERT_FALSE(condition) EXPECT_FALSE(condition)
#define ASSERT_EQ(expected, actual) EXPECT_EQ(expected, actual)
