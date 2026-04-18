#include <gtest/gtest.h>

#include <exception>
#include <utility>
#include <vector>

namespace testing {
namespace {

struct RegisteredTest {
    std::string suite_name;
    std::string test_name;
    std::function<void()> test_body;
};

std::vector<RegisteredTest>& Registry() {
    static std::vector<RegisteredTest> tests;
    return tests;
}

int& FailureCount() {
    static int failures = 0;
    return failures;
}

}  // namespace

void InitGoogleTest(int* argc, char** argv) {
    (void)argc;
    (void)argv;
}

bool RegisterTest(const char* suite_name, const char* test_name, std::function<void()> test_body) {
    Registry().push_back(RegisteredTest{suite_name, test_name, std::move(test_body)});
    return true;
}

void ReportFailure(const char* file, int line, const std::string& message) {
    ++FailureCount();
    std::cout << file << "(" << line << "): Failure" << std::endl;
    std::cout << message << std::endl;
}

int RunAllRegisteredTests() {
    auto& tests = Registry();
    FailureCount() = 0;

    std::cout << "[==========] Running " << tests.size() << " tests from 0 test suites." << std::endl;
    for (const auto& test : tests) {
        const int failures_before = FailureCount();
        std::cout << "[ RUN      ] " << test.suite_name << "." << test.test_name << std::endl;
        try {
            test.test_body();
        } catch (const std::exception& ex) {
            ReportFailure(__FILE__, __LINE__, std::string("Unhandled std::exception: ") + ex.what());
        } catch (...) {
            ReportFailure(__FILE__, __LINE__, "Unhandled unknown exception.");
        }

        if (FailureCount() == failures_before) {
            std::cout << "[       OK ] " << test.suite_name << "." << test.test_name << std::endl;
        } else {
            std::cout << "[  FAILED  ] " << test.suite_name << "." << test.test_name << std::endl;
        }
    }

    std::cout << "[==========] " << tests.size() << " tests from 0 test suites ran." << std::endl;
    if (FailureCount() == 0) {
        std::cout << "[  PASSED  ] " << tests.size() << " tests." << std::endl;
        return 0;
    }

    std::cout << "[  FAILED  ] " << FailureCount() << " tests." << std::endl;
    return 1;
}

}  // namespace testing
