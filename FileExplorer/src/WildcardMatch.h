#pragma once

namespace fileexplorer {

// Case-insensitive wildcard match where:
// - '*' matches zero or more characters
// - '?' matches exactly one character
bool WildcardMatch(const wchar_t* pattern, const wchar_t* text);

}  // namespace fileexplorer
