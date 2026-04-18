#include "TabManager.h"

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>

#include <algorithm>
#include <iterator>

namespace {

std::wstring TrimTrailingSeparators(std::wstring path) {
    if (path.size() <= 3) {
        return path;
    }

    while (!path.empty()) {
        const wchar_t tail = path.back();
        if (tail != L'\\' && tail != L'/') {
            break;
        }
        if (path.size() == 3 && path[1] == L':') {
            break;
        }
        path.pop_back();
    }

    return path;
}

std::wstring NormalizePath(std::wstring path) {
    if (path.empty()) {
        return path;
    }

    std::replace(path.begin(), path.end(), L'/', L'\\');
    path = TrimTrailingSeparators(std::move(path));

    if (path.size() == 2 && path[1] == L':') {
        path.push_back(L'\\');
    }

    return path;
}

bool IsDriveRoot(const std::wstring& path) {
    return path.size() == 3 && path[1] == L':' && path[2] == L'\\';
}

std::wstring ParentPath(const std::wstring& path) {
    const std::wstring normalized = NormalizePath(path);
    if (normalized.empty() || IsDriveRoot(normalized)) {
        return L"";
    }

    const std::wstring::size_type last_separator = normalized.find_last_of(L'\\');
    if (last_separator == std::wstring::npos) {
        return L"";
    }

    if (last_separator == 2 && normalized[1] == L':') {
        return normalized.substr(0, 3);
    }

    if (last_separator == 0) {
        return L"";
    }

    return normalized.substr(0, last_separator);
}

}  // namespace

namespace fileexplorer {

TabManager::TabManager() {
    const int initial_index = AddTab(DefaultTabPath(), true, false);
    if (initial_index < 0) {
        tabs_.push_back(TabState{L"C:\\", L"C:", false, {L"C:\\"}, 0});
        active_index_ = 0;
    }
}

const std::vector<TabState>& TabManager::tabs() const noexcept {
    return tabs_;
}

int TabManager::active_index() const noexcept {
    return active_index_;
}

const TabState* TabManager::active_tab() const noexcept {
    if (!IsValidIndex(active_index_)) {
        return nullptr;
    }
    return &tabs_[active_index_];
}

int TabManager::AddTab(const std::wstring& path, bool activate, bool pinned) {
    std::wstring normalized_path = NormalizePath(path);
    if (normalized_path.empty()) {
        normalized_path = NormalizePath(DefaultTabPath());
    }

    TabState state = {};
    state.path = normalized_path;
    state.displayName = BuildDisplayName(normalized_path);
    state.pinned = pinned;
    state.history.push_back(normalized_path);
    state.historyIndex = 0;

    const int insert_index = active_index_ + 1;
    if (tabs_.empty() || !IsValidIndex(active_index_)) {
        tabs_.push_back(state);
        active_index_ = 0;
        return 0;
    }

    tabs_.insert(tabs_.begin() + insert_index, std::move(state));
    if (activate) {
        active_index_ = insert_index;
    } else if (insert_index <= active_index_) {
        ++active_index_;
    }

    return insert_index;
}

bool TabManager::Activate(int index) {
    if (!IsValidIndex(index)) {
        return false;
    }
    active_index_ = index;
    return true;
}

bool TabManager::ActivateNext() {
    if (tabs_.size() <= 1) {
        return false;
    }
    active_index_ = (active_index_ + 1) % static_cast<int>(tabs_.size());
    return true;
}

bool TabManager::ActivatePrevious() {
    if (tabs_.size() <= 1) {
        return false;
    }
    active_index_ = (active_index_ - 1 + static_cast<int>(tabs_.size())) % static_cast<int>(tabs_.size());
    return true;
}

bool TabManager::JumpToIndex(int index) {
    return Activate(index);
}

bool TabManager::CanCloseTab(int index) const {
    if (!IsValidIndex(index)) {
        return false;
    }
    if (tabs_.size() <= 1) {
        return false;
    }
    return !tabs_[index].pinned;
}

bool TabManager::CanCloseActiveTab() const {
    return CanCloseTab(active_index_);
}

bool TabManager::CloseTab(int index) {
    if (!CanCloseTab(index)) {
        return false;
    }

    tabs_.erase(tabs_.begin() + index);
    if (index < active_index_) {
        --active_index_;
    } else if (index == active_index_) {
        active_index_ = std::min(index, static_cast<int>(tabs_.size()) - 1);
    }

    if (active_index_ < 0) {
        active_index_ = 0;
    }

    return true;
}

bool TabManager::CloseActiveTab() {
    return CloseTab(active_index_);
}

bool TabManager::CloseOtherTabs(int index) {
    if (!IsValidIndex(index)) {
        return false;
    }

    const std::wstring active_path = tabs_[index].path;
    const std::wstring active_display_name = tabs_[index].displayName;

    std::vector<TabState> kept_tabs;
    kept_tabs.reserve(tabs_.size());
    for (int i = 0; i < static_cast<int>(tabs_.size()); ++i) {
        if (i == index || tabs_[i].pinned) {
            kept_tabs.push_back(tabs_[i]);
        }
    }

    if (kept_tabs.size() == tabs_.size()) {
        active_index_ = std::clamp(active_index_, 0, static_cast<int>(tabs_.size()) - 1);
        return false;
    }

    tabs_ = std::move(kept_tabs);
    for (int i = 0; i < static_cast<int>(tabs_.size()); ++i) {
        if (tabs_[i].path == active_path && tabs_[i].displayName == active_display_name) {
            active_index_ = i;
            return true;
        }
    }

    active_index_ = std::clamp(active_index_, 0, static_cast<int>(tabs_.size()) - 1);
    return true;
}

bool TabManager::CloseTabsToRight(int index) {
    if (!IsValidIndex(index)) {
        return false;
    }

    bool changed = false;
    for (int i = static_cast<int>(tabs_.size()) - 1; i > index; --i) {
        if (CanCloseTab(i)) {
            tabs_.erase(tabs_.begin() + i);
            changed = true;
        }
    }

    if (tabs_.empty()) {
        tabs_.push_back(TabState{DefaultTabPath(), BuildDisplayName(DefaultTabPath()), false, {DefaultTabPath()}, 0});
    }

    active_index_ = std::clamp(active_index_, 0, static_cast<int>(tabs_.size()) - 1);
    return changed;
}

bool TabManager::TogglePin(int index) {
    if (!IsValidIndex(index)) {
        return false;
    }
    tabs_[index].pinned = !tabs_[index].pinned;
    return true;
}

bool TabManager::DuplicateTab(int index) {
    if (!IsValidIndex(index)) {
        return false;
    }

    TabState duplicated = tabs_[index];
    duplicated.pinned = false;

    const int insert_index = index + 1;
    tabs_.insert(tabs_.begin() + insert_index, std::move(duplicated));
    active_index_ = insert_index;
    return true;
}

bool TabManager::NavigateTo(const std::wstring& path) {
    if (!IsValidIndex(active_index_)) {
        return false;
    }

    std::wstring normalized_path = NormalizePath(path);
    if (normalized_path.empty()) {
        return false;
    }

    TabState& tab = tabs_[active_index_];
    if (tab.historyIndex < 0 || tab.historyIndex >= static_cast<int>(tab.history.size())) {
        tab.history.clear();
        tab.history.push_back(normalized_path);
        tab.historyIndex = 0;
    } else if (tab.history[tab.historyIndex] != normalized_path) {
        const int next_index = tab.historyIndex + 1;
        if (next_index < static_cast<int>(tab.history.size())) {
            tab.history.erase(tab.history.begin() + next_index, tab.history.end());
        }
        tab.history.push_back(normalized_path);
        tab.historyIndex = static_cast<int>(tab.history.size()) - 1;
    }

    tab.path = normalized_path;
    tab.displayName = BuildDisplayName(normalized_path);
    return true;
}

bool TabManager::NavigateBack() {
    if (!CanNavigateBack()) {
        return false;
    }

    TabState& tab = tabs_[active_index_];
    --tab.historyIndex;
    tab.path = tab.history[tab.historyIndex];
    tab.displayName = BuildDisplayName(tab.path);
    return true;
}

bool TabManager::NavigateForward() {
    if (!CanNavigateForward()) {
        return false;
    }

    TabState& tab = tabs_[active_index_];
    ++tab.historyIndex;
    tab.path = tab.history[tab.historyIndex];
    tab.displayName = BuildDisplayName(tab.path);
    return true;
}

bool TabManager::NavigateUp() {
    if (!CanNavigateUp()) {
        return false;
    }

    const std::wstring parent = ParentPath(tabs_[active_index_].path);
    if (parent.empty()) {
        return false;
    }
    return NavigateTo(parent);
}

bool TabManager::RefreshActive() {
    if (!IsValidIndex(active_index_)) {
        return false;
    }
    tabs_[active_index_].displayName = BuildDisplayName(tabs_[active_index_].path);
    return true;
}

bool TabManager::CanNavigateBack() const {
    if (!IsValidIndex(active_index_)) {
        return false;
    }

    const TabState& tab = tabs_[active_index_];
    return tab.historyIndex > 0 && tab.historyIndex < static_cast<int>(tab.history.size());
}

bool TabManager::CanNavigateForward() const {
    if (!IsValidIndex(active_index_)) {
        return false;
    }

    const TabState& tab = tabs_[active_index_];
    return tab.historyIndex >= 0 && (tab.historyIndex + 1) < static_cast<int>(tab.history.size());
}

bool TabManager::CanNavigateUp() const {
    if (!IsValidIndex(active_index_)) {
        return false;
    }
    return !ParentPath(tabs_[active_index_].path).empty();
}

std::wstring TabManager::DefaultTabPath() {
    wchar_t buffer[MAX_PATH] = {};
    const DWORD chars = GetEnvironmentVariableW(L"USERPROFILE", buffer, static_cast<DWORD>(std::size(buffer)));
    if (chars > 0 && chars < std::size(buffer)) {
        return std::wstring(buffer);
    }
    return L"C:\\";
}

std::wstring TabManager::BuildDisplayName(const std::wstring& path) {
    const std::wstring normalized = TrimTrailingSeparators(path);
    if (normalized.empty()) {
        return L"New Tab";
    }

    if (normalized.size() >= 2 && normalized[1] == L':') {
        if (normalized.size() <= 3) {
            return normalized.substr(0, 2);
        }
    }

    const std::wstring::size_type last_separator = normalized.find_last_of(L"\\/");
    if (last_separator == std::wstring::npos) {
        return normalized;
    }

    if (last_separator + 1 >= normalized.size()) {
        return normalized;
    }

    return normalized.substr(last_separator + 1);
}

bool TabManager::IsValidIndex(int index) const noexcept {
    return index >= 0 && index < static_cast<int>(tabs_.size());
}

}  // namespace fileexplorer
