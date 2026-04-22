#pragma once

#include <cstdint>
#include <string>
#include <vector>

namespace fileexplorer {

struct TabState {
    uint64_t id = 0;
    std::wstring path;
    std::wstring displayName;
    bool pinned = false;
    std::vector<std::wstring> history;
    int historyIndex = -1;
};

class TabManager final {
public:
    struct SessionTab {
        std::wstring path;
        bool pinned = false;
    };

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
    bool MoveTab(int from_index, int to_index);

    bool NavigateTo(const std::wstring& path);
    bool NavigateBack();
    bool NavigateForward();
    bool NavigateUp();
    bool RefreshActive();

    bool CanNavigateBack() const;
    bool CanNavigateForward() const;
    bool CanNavigateUp() const;

    std::vector<SessionTab> CaptureSession() const;
    bool RestoreSession(const std::vector<SessionTab>& tabs, int active_index);

    static std::wstring DefaultTabPath();
    static std::wstring BuildDisplayName(const std::wstring& path);

private:
    bool IsValidIndex(int index) const noexcept;

    std::vector<TabState> tabs_;
    int active_index_{0};
    uint64_t next_tab_id_{1};
};

}  // namespace fileexplorer
