#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>

#include <atomic>
#include <array>
#include <memory>
#include <string>
#include <thread>
#include <type_traits>
#include <unordered_map>
#include <unordered_set>

#include "FileListView.h"
#include "FolderWorker.h"
#include "NavBar.h"
#include "Settings.h"
#include "Sidebar.h"
#include "TabManager.h"
#include "TabStrip.h"
#include "Theme.h"

namespace fileexplorer {

class MainWindow final {
public:
    MainWindow();
    ~MainWindow();

    MainWindow(const MainWindow&) = delete;
    MainWindow& operator=(const MainWindow&) = delete;

    bool Create(HINSTANCE instance, int show_command);
    int MessageLoop();

private:
    struct BrushDeleter {
        void operator()(HBRUSH brush) const noexcept;
    };
    struct FontDeleter {
        void operator()(HFONT font) const noexcept;
    };

    using UniqueBrush = std::unique_ptr<std::remove_pointer_t<HBRUSH>, BrushDeleter>;
    using UniqueFont = std::unique_ptr<std::remove_pointer_t<HFONT>, FontDeleter>;

    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM w_param, LPARAM l_param);
    static LRESULT CALLBACK SidebarSplitterProc(HWND hwnd, UINT message, WPARAM w_param, LPARAM l_param);
    LRESULT HandleMessage(UINT message, WPARAM w_param, LPARAM l_param);
    LRESULT HandleSidebarSplitterMessage(HWND hwnd, UINT message, WPARAM w_param, LPARAM l_param);

    bool RegisterWindowClass();
    bool CreateMainWindow(int show_command);
    bool InitializeChildZones();
    void CreateZoneBrushes();
    void DestroyZoneBrushes();
    void ApplyWindowChrome();
    void RecalculateLayout();
    void EnsureStatusBarFont();
    void LayoutChildZones();
    void LoadLayoutSettings();
    void SaveLayoutSettings() const;
    bool ApplySidebarWidthFromPointerX(int pointer_x);
    int EffectiveSidebarWidthPx(int client_width) const;
    int ClampSidebarWidthPx(int requested_sidebar_width_px, int client_width) const;
    void OnDpiChanged(UINT new_dpi, RECT* suggested_rect);
    void PaintFallbackBackground(HDC hdc);
    void PaintFileListInsetGutter(HDC hdc);
    void UpdateWindowTitle();
    void HandleTabStateChanged();
    void StartFolderLoad(
        const std::wstring& path,
        bool incremental_refresh,
        const std::wstring& post_load_select_name = L"");
    void StartSearch(const std::wstring& pattern);
    void CancelSearchWorker();
    void ExitSearchModeUi(bool clear_sidebar_text);
    void StoreSearchSnapshotForTab(uint64_t tab_id);
    void RemoveSearchSnapshotForTab(uint64_t tab_id);
    void PruneSearchSnapshots();
    void PreparePostLoadSelectionForTransition(const std::wstring& from_path, const std::wstring& to_path);
    void StartDebouncedRefreshForActiveFolder(uint64_t watch_generation);
    void RestartFolderWatcher(const std::wstring& path);
    void StopFolderWatcher();
    bool HandleTabKeyboardShortcut(WPARAM key_code, bool ctrl_down, bool shift_down);

    HINSTANCE instance_{nullptr};
    HWND hwnd_{nullptr};
    HWND tab_strip_hwnd_{nullptr};
    HWND nav_bar_hwnd_{nullptr};
    HWND file_list_hwnd_{nullptr};
    HWND sidebar_hwnd_{nullptr};
    HWND sidebar_splitter_hwnd_{nullptr};
    HWND status_bar_hwnd_{nullptr};

    UINT dpi_{96U};
    LayoutMetrics metrics_{};
    int sidebar_width_logical_{320};
    Settings::FileListColumnWidthsLogical file_list_column_widths_logical_{{260, 60, 140, 90, 320}};
    bool sidebar_resize_active_{false};

    bool use_solid_fallback_background_{true};

    TabManager tab_manager_{};
    TabStrip tab_strip_{};
    NavBar nav_bar_{};
    FileListView file_list_view_{};
    Sidebar sidebar_{};

    std::shared_ptr<std::atomic<uint64_t>> folder_generation_source_{};
    std::shared_ptr<std::atomic<uint64_t>> search_generation_source_{};
    std::shared_ptr<std::atomic<bool>> search_cancel_token_{};
    bool search_active_{false};
    std::wstring search_root_path_{};
    std::wstring search_pattern_{};
    ULONGLONG search_started_tick_ms_{0};
    uint64_t last_active_tab_id_{0};
    std::unordered_map<uint64_t, FileListView::SearchSnapshot> search_snapshots_{};
    bool pending_incremental_refresh_{false};

    std::thread folder_watch_thread_{};
    std::shared_ptr<std::atomic<bool>> folder_watch_stop_{};
    std::atomic<uint64_t> folder_watch_generation_{0};
    std::wstring watched_folder_path_{};
    bool pending_debounced_refresh_{false};
    uint64_t pending_debounced_refresh_generation_{0};
    std::wstring pending_post_load_selection_name_{};

    UniqueBrush nav_brush_{nullptr};
    UniqueBrush file_list_brush_{nullptr};
    UniqueBrush sidebar_brush_{nullptr};
    UniqueBrush status_brush_{nullptr};
    UniqueFont status_font_{nullptr};
    int status_font_pixel_height_{0};

    Settings settings_{};
};

}  // namespace fileexplorer
