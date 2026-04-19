#include "CsvParser.h"

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>

#include <algorithm>

namespace {

void LogLastError(const wchar_t* context) {
    const DWORD error_code = GetLastError();
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (GetLastError=%lu).\r\n", context, error_code);
    OutputDebugStringW(buffer);
}

std::wstring TrimWhitespace(std::wstring value) {
    const auto not_space = [](wchar_t ch) { return !iswspace(ch); };
    const auto begin = std::find_if(value.begin(), value.end(), not_space);
    if (begin == value.end()) {
        return L"";
    }
    const auto reverse_begin = std::find_if(value.rbegin(), value.rend(), not_space).base();
    return std::wstring(begin, reverse_begin);
}

bool NeedsQuoting(const std::wstring& value) {
    if (value.empty()) {
        return false;
    }

    if (iswspace(value.front()) || iswspace(value.back())) {
        return true;
    }

    return value.find_first_of(L",\"\r\n") != std::wstring::npos;
}

std::wstring EscapeForCsv(const std::wstring& value) {
    std::wstring escaped;
    escaped.reserve(value.size() + 8);
    for (const wchar_t ch : value) {
        if (ch == L'"') {
            escaped.push_back(L'"');
        }
        escaped.push_back(ch);
    }
    return escaped;
}

bool WideToUtf8(const std::wstring& text, std::string* utf8_out) {
    if (utf8_out == nullptr) {
        return false;
    }

    utf8_out->clear();
    if (text.empty()) {
        return true;
    }

    const int required_bytes = WideCharToMultiByte(
        CP_UTF8,
        0,
        text.c_str(),
        static_cast<int>(text.size()),
        nullptr,
        0,
        nullptr,
        nullptr);
    if (required_bytes <= 0) {
        return false;
    }

    utf8_out->resize(static_cast<size_t>(required_bytes));
    const int converted = WideCharToMultiByte(
        CP_UTF8,
        0,
        text.c_str(),
        static_cast<int>(text.size()),
        utf8_out->data(),
        required_bytes,
        nullptr,
        nullptr);
    return converted == required_bytes;
}

bool Utf8ToWide(const char* bytes, int byte_count, std::wstring* wide_out) {
    if (wide_out == nullptr || bytes == nullptr || byte_count < 0) {
        return false;
    }

    wide_out->clear();
    if (byte_count == 0) {
        return true;
    }

    int required_chars = MultiByteToWideChar(
        CP_UTF8,
        MB_ERR_INVALID_CHARS,
        bytes,
        byte_count,
        nullptr,
        0);
    if (required_chars <= 0) {
        required_chars = MultiByteToWideChar(
            CP_UTF8,
            0,
            bytes,
            byte_count,
            nullptr,
            0);
        if (required_chars <= 0) {
            return false;
        }
    }

    wide_out->resize(static_cast<size_t>(required_chars));
    const int converted = MultiByteToWideChar(
        CP_UTF8,
        0,
        bytes,
        byte_count,
        wide_out->data(),
        required_chars);
    return converted == required_chars;
}

}  // namespace

namespace fileexplorer {

CsvParser::CsvRows CsvParser::Parse(std::wstring_view text) {
    CsvRows rows;
    CsvRow current_row;
    std::wstring current_field;

    bool in_quotes = false;
    bool field_was_quoted = false;
    bool after_quote = false;
    bool saw_any_non_newline = false;

    auto finalize_field = [&]() {
        std::wstring value = current_field;
        if (!field_was_quoted) {
            value = TrimWhitespace(std::move(value));
        }
        current_row.push_back(std::move(value));
        current_field.clear();
        field_was_quoted = false;
        after_quote = false;
    };

    auto finalize_row = [&]() {
        bool row_has_content = false;
        for (const std::wstring& field : current_row) {
            if (!field.empty()) {
                row_has_content = true;
                break;
            }
        }
        if (row_has_content || saw_any_non_newline) {
            rows.push_back(std::move(current_row));
        }
        current_row.clear();
        saw_any_non_newline = false;
    };

    const size_t length = text.size();
    for (size_t i = 0; i < length; ++i) {
        const wchar_t ch = text[i];

        if (in_quotes) {
            if (ch == L'"') {
                const bool escaped_quote = (i + 1 < length && text[i + 1] == L'"');
                if (escaped_quote) {
                    current_field.push_back(L'"');
                    ++i;
                } else {
                    in_quotes = false;
                    after_quote = true;
                }
            } else {
                current_field.push_back(ch);
            }
            continue;
        }

        if (after_quote) {
            if (ch == L',' || ch == L'\r' || ch == L'\n') {
                finalize_field();
                if (ch == L'\r') {
                    if (i + 1 < length && text[i + 1] == L'\n') {
                        ++i;
                    }
                    finalize_row();
                    continue;
                }
                if (ch == L'\n') {
                    finalize_row();
                    continue;
                }
                continue;
            }

            if (iswspace(ch)) {
                continue;
            }

            // Tolerate malformed CSV that contains data after closing quote.
            current_field.push_back(ch);
            saw_any_non_newline = true;
            after_quote = false;
            continue;
        }

        if (ch == L'"' && current_field.empty()) {
            in_quotes = true;
            field_was_quoted = true;
            saw_any_non_newline = true;
            continue;
        }

        if (ch == L',') {
            finalize_field();
            saw_any_non_newline = true;
            continue;
        }

        if (ch == L'\r' || ch == L'\n') {
            if (ch == L'\r' && i + 1 < length && text[i + 1] == L'\n') {
                ++i;
            }
            finalize_field();
            finalize_row();
            continue;
        }

        current_field.push_back(ch);
        if (!iswspace(ch)) {
            saw_any_non_newline = true;
        }
    }

    if (in_quotes || after_quote || !current_field.empty() || !current_row.empty()) {
        finalize_field();
    }
    if (!current_row.empty()) {
        finalize_row();
    }

    return rows;
}

std::wstring CsvParser::Serialize(const CsvRows& rows) {
    std::wstring output;

    for (size_t row_index = 0; row_index < rows.size(); ++row_index) {
        const CsvRow& row = rows[row_index];
        for (size_t column = 0; column < row.size(); ++column) {
            const std::wstring& field = row[column];
            if (NeedsQuoting(field)) {
                output.push_back(L'"');
                output.append(EscapeForCsv(field));
                output.push_back(L'"');
            } else {
                output.append(field);
            }

            if (column + 1 < row.size()) {
                output.push_back(L',');
            }
        }

        if (row_index + 1 < rows.size()) {
            output.append(L"\r\n");
        }
    }

    return output;
}

bool CsvParser::ReadFile(const std::wstring& path, CsvRows* rows_out, bool* file_exists_out) {
    if (rows_out == nullptr) {
        return false;
    }
    rows_out->clear();

    const DWORD attributes = GetFileAttributesW(path.c_str());
    const bool file_exists =
        (attributes != INVALID_FILE_ATTRIBUTES) && ((attributes & FILE_ATTRIBUTE_DIRECTORY) == 0);
    if (file_exists_out != nullptr) {
        *file_exists_out = file_exists;
    }

    if (!file_exists) {
        return true;
    }

    HANDLE file_handle = CreateFileW(
        path.c_str(),
        GENERIC_READ,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        nullptr,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        nullptr);
    if (file_handle == INVALID_HANDLE_VALUE) {
        LogLastError(L"CreateFileW(CsvParser Read)");
        return false;
    }

    LARGE_INTEGER size = {};
    if (!GetFileSizeEx(file_handle, &size)) {
        LogLastError(L"GetFileSizeEx(CsvParser)");
        CloseHandle(file_handle);
        return false;
    }

    if (size.QuadPart < 0 || size.QuadPart > static_cast<LONGLONG>(0x7FFFFFFF)) {
        CloseHandle(file_handle);
        return false;
    }

    const DWORD byte_count = static_cast<DWORD>(size.QuadPart);
    std::vector<char> bytes(byte_count);
    DWORD read_bytes = 0;
    bool read_ok = true;
    if (byte_count > 0) {
        read_ok =
            ::ReadFile(file_handle, bytes.data(), byte_count, &read_bytes, nullptr) != FALSE &&
            read_bytes == byte_count;
    }
    if (!CloseHandle(file_handle)) {
        LogLastError(L"CloseHandle(CsvParser Read)");
    }

    if (!read_ok) {
        LogLastError(L"ReadFile(CsvParser)");
        return false;
    }

    std::wstring content;
    if (byte_count >= 2 &&
        static_cast<unsigned char>(bytes[0]) == 0xFF &&
        static_cast<unsigned char>(bytes[1]) == 0xFE) {
        const size_t utf16_bytes = byte_count - 2;
        const size_t wchar_count = utf16_bytes / sizeof(wchar_t);
        content.assign(
            reinterpret_cast<const wchar_t*>(bytes.data() + 2),
            reinterpret_cast<const wchar_t*>(bytes.data() + 2) + wchar_count);
    } else {
        int offset = 0;
        if (byte_count >= 3 &&
            static_cast<unsigned char>(bytes[0]) == 0xEF &&
            static_cast<unsigned char>(bytes[1]) == 0xBB &&
            static_cast<unsigned char>(bytes[2]) == 0xBF) {
            offset = 3;
        }
        if (!Utf8ToWide(bytes.data() + offset, static_cast<int>(byte_count - offset), &content)) {
            return false;
        }
    }

    *rows_out = Parse(content);
    return true;
}

bool CsvParser::WriteFile(const std::wstring& path, const CsvRows& rows) {
    const std::wstring serialized = Serialize(rows);

    std::string utf8;
    if (!WideToUtf8(serialized, &utf8)) {
        return false;
    }

    HANDLE file_handle = CreateFileW(
        path.c_str(),
        GENERIC_WRITE,
        FILE_SHARE_READ,
        nullptr,
        CREATE_ALWAYS,
        FILE_ATTRIBUTE_NORMAL,
        nullptr);
    if (file_handle == INVALID_HANDLE_VALUE) {
        LogLastError(L"CreateFileW(CsvParser Write)");
        return false;
    }

    const unsigned char bom[] = {0xEF, 0xBB, 0xBF};
    DWORD written = 0;
    bool ok = ::WriteFile(file_handle, bom, sizeof(bom), &written, nullptr) != FALSE && written == sizeof(bom);
    if (ok && !utf8.empty()) {
        written = 0;
        ok = ::WriteFile(
                 file_handle,
                 utf8.data(),
                 static_cast<DWORD>(utf8.size()),
                 &written,
                 nullptr) != FALSE &&
            written == static_cast<DWORD>(utf8.size());
    }

    if (!CloseHandle(file_handle)) {
        LogLastError(L"CloseHandle(CsvParser Write)");
    }
    if (!ok) {
        LogLastError(L"WriteFile(CsvParser)");
    }
    return ok;
}

}  // namespace fileexplorer
