#include "WildcardMatch.h"

#include <cwctype>

namespace fileexplorer {

bool WildcardMatch(const wchar_t* pattern, const wchar_t* text) {
    if (pattern == nullptr || text == nullptr) {
        return false;
    }

    const wchar_t* pattern_cursor = pattern;
    const wchar_t* text_cursor = text;
    const wchar_t* last_star_pattern = nullptr;
    const wchar_t* last_star_text = nullptr;

    while (*text_cursor != L'\0') {
        if (*pattern_cursor == L'*') {
            while (*pattern_cursor == L'*') {
                ++pattern_cursor;
            }

            if (*pattern_cursor == L'\0') {
                return true;
            }

            last_star_pattern = pattern_cursor;
            last_star_text = text_cursor;
            continue;
        }

        const wchar_t pattern_ch = static_cast<wchar_t>(towlower(*pattern_cursor));
        const wchar_t text_ch = static_cast<wchar_t>(towlower(*text_cursor));
        if (*pattern_cursor == L'?' || (*pattern_cursor != L'\0' && pattern_ch == text_ch)) {
            ++pattern_cursor;
            ++text_cursor;
            continue;
        }

        if (last_star_pattern != nullptr && last_star_text != nullptr) {
            ++last_star_text;
            text_cursor = last_star_text;
            pattern_cursor = last_star_pattern;
            continue;
        }

        return false;
    }

    while (*pattern_cursor == L'*') {
        ++pattern_cursor;
    }

    return *pattern_cursor == L'\0';
}

}  // namespace fileexplorer
