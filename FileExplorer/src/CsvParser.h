#pragma once

#include <string>
#include <string_view>
#include <vector>

namespace fileexplorer {

class CsvParser final {
public:
    using CsvRow = std::vector<std::wstring>;
    using CsvRows = std::vector<CsvRow>;

    static CsvRows Parse(std::wstring_view text);
    static std::wstring Serialize(const CsvRows& rows);

    static bool ReadFile(
        const std::wstring& path,
        CsvRows* rows_out,
        bool* file_exists_out = nullptr);
    static bool WriteFile(const std::wstring& path, const CsvRows& rows);
};

}  // namespace fileexplorer
