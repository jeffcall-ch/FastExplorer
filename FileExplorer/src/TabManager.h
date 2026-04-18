#pragma once

#include <string>
#include <vector>

namespace fileexplorer {

struct TabState {
    std::wstring path;
    std::wstring displayName;
    bool pinned = false;
    std::vector<std::wstring> history;
    int historyIndex = -1;
};

class TabManager final {
public:
    TabManager();

    const std::vector<TabState>& tabs() const noexcept;
    int active_index() const noexcept;
    const TabState* active_tab() const noexcept;

    int AddTab(const std::wstring& path, bool activate, bool pinned);
    bool Activate(int index);
    bool ActivateNext();
    bool ActivatePrevious();
    bool JumpToIndex(int index);

    bool CanCloseTab(int index) const;
    bool CanCloseActiveTab() const;
    bool CloseTab(int index);
    bool CloseActiveTab();
    bool CloseOtherTabs(int index);
    bool CloseTabsToRight(int index);

    bool TogglePin(int index);
    bool DuplicateTab(int index);

    bool NavigateTo(const std::wstring& path);
    bool NavigateBack();
    bool NavigateForward();
    bool NavigateUp();
    bool RefreshActive();

    bool CanNavigateBack() const;
    bool CanNavigateForward() const;
    bool CanNavigateUp() const;

    static std::wstring DefaultTabPath();
    static std::wstring BuildDisplayName(const std::wstring& path);

private:
    bool IsValidIndex(int index) const noexcept;

    std::vector<TabState> tabs_;
    int active_index_{0};
};

}  // namespace fileexplorer
