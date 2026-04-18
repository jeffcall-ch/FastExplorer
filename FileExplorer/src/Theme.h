#pragma once

namespace fileexplorer {

struct LayoutMetrics {
    int tabStripHeight;
    int navBarHeight;
    int statusBarHeight;
    int sidebarWidth;
    int sidebarResizeGripWidth;
    int searchBarHeight;
};

LayoutMetrics DefaultLayoutMetrics();
LayoutMetrics ScaleLayoutMetrics(unsigned int dpi);

}  // namespace fileexplorer
