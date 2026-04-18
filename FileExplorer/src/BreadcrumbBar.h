#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>
#include <CommCtrl.h>

#include <atomic>
#include <memory>
#include <string>
#include <type_traits>
#include <vector>

namespace fileexplorer {

constexpr UINT WM_FE_BREADCRUMB_NAVIGATE = WM_APP + 102;

class BreadcrumbBar final {
public:
    BreadcrumbBar();
    ~BreadcrumbBar();

    BreadcrumbBar(const BreadcrumbBar&) = delete;
    BreadcrumbBar& operator=(const BreadcrumbBar&) = delete;

    bool Create(HWND parent, HINSTANCE instance, int control_id);
    HWND hwnd() const noexcept;

    void SetDpi(UINT dpi);
    void SetPath(const std::wstring& path);
    void ActivateEditMode();
    void CancelEditMode();
    bool IsEditMode() const noexcept;
    void Refresh();

private:
    enum class HitPart {
        None,
        Segment,
        Separator,
    };

    struct HitTestResult {
        HitPart part = HitPart::None;
        int index = -1;
    };

    struct SegmentLayout {
        std::wstring label;
        std::wstring path;
        RECT segment_rect{};
        RECT separator_rect{};
    };

    struct SiblingMenuPayload {
        uint64_t generation = 0;
        std::wstring base_path;
        POINT anchor_screen{};
        std::vector<std::wstring> sibling_names;
    };

    struct FontDeleter {
        void operator()(HFONT font) const noexcept;
    };

    struct BrushDeleter {
        void operator()(HBRUSH brush) const noexcept;
    };

    using UniqueFont = std::unique_ptr<std::remove_pointer_t<HFONT>, FontDeleter>;
    using UniqueBrush = std::unique_ptr<std::remove_pointer_t<HBRUSH>, BrushDeleter>;

    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM w_param, LPARAM l_param);
    static LRESULT CALLBACK EditSubclassProc(
        HWND hwnd,
        UINT message,
        WPARAM w_param,
        LPARAM l_param,
        UINT_PTR subclass_id,
        DWORD_PTR ref_data);
    LRESULT HandleMessage(UINT message, WPARAM w_param, LPARAM l_param);

    bool RegisterWindowClass();
    bool CreateEditControl();
    void LayoutEditControl();
    void DrawControl(HDC hdc);
    void RebuildSegments(HDC hdc);
    void EnsureSegmentsLayout();
    HitTestResult HitTest(POINT point) const;

    void HandleMouseMove(POINT point);
    void HandleMouseLeave();
    void HandleLeftClick(POINT point);

    void PostNavigatePath(const std::wstring& path) const;
    void StartSiblingFlyout(int segment_index, POINT anchor_screen);
    void ShowSiblingFlyout(const SiblingMenuPayload& payload);

    void CommitEditMode();
    void ExitEditMode(bool return_focus_to_control);

    std::vector<SegmentLayout> BuildSegmentsFromPath(const std::wstring& path) const;
    int MeasureTextWidth(HDC hdc, const std::wstring& text) const;

    HWND parent_hwnd_{nullptr};
    HWND hwnd_{nullptr};
    HWND edit_hwnd_{nullptr};
    HINSTANCE instance_{nullptr};
    int control_id_{0};
    UINT dpi_{96U};

    std::wstring current_path_{L"C:\\"};
    std::vector<SegmentLayout> segments_;

    int hovered_segment_index_{-1};
    int hovered_separator_index_{-1};
    bool edit_mode_{false};

    std::atomic<uint64_t> siblings_generation_{0};

    UniqueFont font_{nullptr};
    int font_pixel_height_{0};
    UniqueBrush edit_background_brush_{nullptr};
};

}  // namespace fileexplorer
