#include "Theme.h"

#include "Dpi.h"

namespace {

constexpr fileexplorer::LayoutMetrics kDefaultMetrics = {
    36,
    38,
    24,
    320,
    4,
    32,
};

}  // namespace

namespace fileexplorer {

LayoutMetrics DefaultLayoutMetrics() {
    return kDefaultMetrics;
}

LayoutMetrics ScaleLayoutMetrics(unsigned int dpi) {
    const LayoutMetrics base = DefaultLayoutMetrics();
    return {
        ScaleForDpi(base.tabStripHeight, dpi),
        ScaleForDpi(base.navBarHeight, dpi),
        ScaleForDpi(base.statusBarHeight, dpi),
        ScaleForDpi(base.sidebarWidth, dpi),
        ScaleForDpi(base.sidebarResizeGripWidth, dpi),
        ScaleForDpi(base.searchBarHeight, dpi),
    };
}

}  // namespace fileexplorer
