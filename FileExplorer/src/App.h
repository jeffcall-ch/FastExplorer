#pragma once

#include <Windows.h>

namespace fileexplorer {

class App final {
public:
    int Run(HINSTANCE instance, int show_command);
};

}  // namespace fileexplorer
