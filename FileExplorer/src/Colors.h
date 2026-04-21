#pragma once

#include <Windows.h>

namespace fileexplorer::colors {

// Win32 COLORREF uses 0x00BBGGRR byte order.
constexpr COLORREF kBackground = 0x001C1C1C;
constexpr COLORREF kForeground = 0x00E8E8E8;

// Phase 10 shared shell surfaces (kept visually identical to current UI).
constexpr COLORREF kBgBase = 0x00201A16;          // RGB(0x16, 0x1A, 0x20)
constexpr COLORREF kBgNavBar = 0x00261F1A;        // RGB(0x1A, 0x1F, 0x26)
constexpr COLORREF kBgFileList = 0x00221C18;      // RGB(0x18, 0x1C, 0x22)
constexpr COLORREF kBgSidebar = 0x002A221C;       // RGB(0x1C, 0x22, 0x2A)
constexpr COLORREF kBgStatus = 0x00201A16;        // RGB(0x16, 0x1A, 0x20)
constexpr COLORREF kTextPrimary = 0x00EEE7E2;     // RGB(0xE2, 0xE7, 0xEE)
constexpr COLORREF kSeparatorSubtle = 0x003D332C; // RGB(0x2C, 0x33, 0x3D)

// File type accent colors (dark mode).
constexpr COLORREF kFilePdf = 0x00756CE0;      // RGB(0xE0, 0x6C, 0x75)
constexpr COLORREF kFileExcel = 0x0079C398;    // RGB(0x98, 0xC3, 0x79)
constexpr COLORREF kFileWord = 0x00EFAF61;     // RGB(0x61, 0xAF, 0xEF)
constexpr COLORREF kFileArchive = 0x007BC0E5;  // RGB(0xE5, 0xC0, 0x7B)
constexpr COLORREF kFileText = 0x00C2B656;     // RGB(0x56, 0xB6, 0xC2)

}  // namespace fileexplorer::colors
