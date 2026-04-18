#include "Dpi.h"

#include <Windows.h>

namespace fileexplorer {

int ScaleForDpi(int value, unsigned int dpi) {
    return MulDiv(value, static_cast<int>(dpi), 96);
}

}  // namespace fileexplorer
