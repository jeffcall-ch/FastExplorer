#pragma once

#include <Windows.h>

namespace fileexplorer {

class App final {
public:
    int Run(HINSTANCE instance, int show_command, const wchar_t* command_line);
};

}  // namespace fileexplorer
