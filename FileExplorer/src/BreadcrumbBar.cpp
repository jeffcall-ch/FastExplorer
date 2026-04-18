#include "BreadcrumbBar.h"

#include <Windowsx.h>
#include <Shlwapi.h>
#include <uxtheme.h>

#include <algorithm>
#include <cwchar>
#include <thread>

#include "WorkerMessages.h"

namespace {

constexpr wchar_t kBreadcrumbClassName[] = L"FE_BreadcrumbBar";

constexpr COLORREF kBackgroundColor = RGB(0x20, 0x26, 0x2F);
constexpr COLORREF kBorderColor = RGB(0x37, 0x40, 0x4B);
constexpr COLORREF kTextColor = RGB(0xE2, 0xE8, 0xF0);
constexpr COLORREF kHoverFillColor = RGB(0x2C, 0x34, 0x3F);
constexpr COLORREF kSeparatorColor = RGB(0x9B, 0xA4, 0xB2);
constexpr COLORREF kEditBackgroundColor = RGB(0x20, 0x26, 0x2F);
constexpr COLORREF kEditTextColor = RGB(0xE2, 0xE8, 0xF0);

constexpr UINT_PTR kEditSubclassId = 1;
constexpr UINT kSiblingMenuIdBase = 6000;

void LogLastError(const wchar_t* context) {
    const DWORD error_code = GetLastError();
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (GetLastError=%lu).\r\n", context, error_code);
    OutputDebugStringW(buffer);
}

void LogHResult(const wchar_t* context, HRESULT hr) {
    wchar_t buffer[256] = {};
    swprintf_s(buffer, L"[FileExplorer] %s failed (HRESULT=0x%08lX).\r\n", context, static_cast<unsigned long>(hr));
    OutputDebugStringW(buffer);
}

bool RectContainsPoint(const RECT& rect, POINT point) {
    return point.x >= rect.left && point.x < rect.right && point.y >= rect.top && point.y < rect.bottom;
}

bool IsDirectoryRoot(const std::wstring& path) {
    return path.size() == 3 && path[1] == L':' && path[2] == L'\\';
}

std::wstring NormalizePath(std::wstring path) {
    if (path.empty()) {
        return path;
    }

    std::replace(path.begin(), path.end(), L'/', L'\\');

    if (path.size() > 3) {
        while (!path.empty() && (path.back() == L'\\' || path.back() == L'/')) {
            if (IsDirectoryRoot(path)) {
                break;
            }
            path.pop_back();
        }
    }

    if (path.size() == 2 && path[1] == L':') {
        path.push_back(L'\\');
    }

    return path;
}

std::wstring JoinPath(std::wstring left, const std::wstring& right) {
    if (left.empty()) {
        return right;
    }

    if (left.back() != L'\\') {
        left.push_back(L'\\');
    }
    left.append(right);
    return left;
}

struct FindHandle {
    HANDLE handle = INVALID_HANDLE_VALUE;

    ~FindHandle() {
        if (handle != INVALID_HANDLE_VALUE) {
            FindClose(handle);
        }
    }
};

void FillRoundedRect(HDC hdc, const RECT& rect, COLORREF fill, COLORREF border, int radius) {
    HBRUSH brush = CreateSolidBrush(fill);
    HPEN pen = CreatePen(PS_SOLID, 1, border);
    if (brush == nullptr || pen == nullptr) {
        if (brush != nullptr) {
            DeleteObject(brush);
        }
        if (pen != nullptr) {
            DeleteObject(pen);
        }
        return;
    }

    HGDIOBJ old_brush = SelectObject(hdc, brush);
    HGDIOBJ old_pen = SelectObject(hdc, pen);
    RoundRect(hdc, rect.left, rect.top, rect.right, rect.bottom, radius, radius);
    SelectObject(hdc, old_pen);
    SelectObject(hdc, old_brush);
    DeleteObject(pen);
    DeleteObject(brush);
}

}  // namespace

namespace fileexplorer {

void BreadcrumbBar::FontDeleter::operator()(HFONT font) const noexcept {
    if (font != nullptr) {
        DeleteObject(font);
    }
}

void BreadcrumbBar::BrushDeleter::operator()(HBRUSH brush) const noexcept {
    if (brush != nullptr) {
        DeleteObject(brush);
    }
}

BreadcrumbBar::BreadcrumbBar() = default;

BreadcrumbBar::~BreadcrumbBar() {
    if (edit_hwnd_ != nullptr) {
        RemoveWindowSubclass(edit_hwnd_, &BreadcrumbBar::EditSubclassProc, kEditSubclassId);
        edit_hwnd_ = nullptr;
    }
}

bool BreadcrumbBar::Create(HWND parent, HINSTANCE instance, int control_id) {
    parent_hwnd_ = parent;
    instance_ = instance;
    control_id_ = control_id;

    if (!RegisterWindowClass()) {
        return false;
    }

    hwnd_ = CreateWindowExW(
        0,
        kBreadcrumbClassName,
        L"",
        WS_CHILD | WS_VISIBLE,
        0,
        0,
        0,
        0,
        parent,
        reinterpret_cast<HMENU>(static_cast<INT_PTR>(control_id)),
        instance,
        this);
    if (hwnd_ == nullptr) {
        LogLastError(L"CreateWindowExW(BreadcrumbBar)");
        return false;
    }

    return true;
}

HWND BreadcrumbBar::hwnd() const noexcept {
    return hwnd_;
}

void BreadcrumbBar::SetDpi(UINT dpi) {
    dpi_ = (dpi == 0U) ? 96U : dpi;
    font_.reset();
    font_pixel_height_ = 0;
    LayoutEditControl();
    Refresh();
}

void BreadcrumbBar::SetPath(const std::wstring& path) {
    std::wstring normalized = NormalizePath(path);
    if (normalized.empty()) {
        normalized = L"C:\\";
    }

    current_path_ = normalized;
    if (edit_mode_ && edit_hwnd_ != nullptr) {
        if (!SetWindowTextW(edit_hwnd_, current_path_.c_str())) {
            LogLastError(L"SetWindowTextW(BreadcrumbEdit)");
        }
    }

    Refresh();
}

void BreadcrumbBar::ActivateEditMode() {
    if (edit_hwnd_ == nullptr) {
        return;
    }

    edit_mode_ = true;
    if (!SetWindowTextW(edit_hwnd_, current_path_.c_str())) {
        LogLastError(L"SetWindowTextW(ActivateEditMode)");
    }
    ShowWindow(edit_hwnd_, SW_SHOW);
    SetFocus(edit_hwnd_);
    SendMessageW(edit_hwnd_, EM_SETSEL, 0, -1);
    Refresh();
}

void BreadcrumbBar::CancelEditMode() {
    ExitEditMode(true);
}

bool BreadcrumbBar::IsEditMode() const noexcept {
    return edit_mode_;
}

void BreadcrumbBar::Refresh() {
    if (hwnd_ != nullptr) {
        InvalidateRect(hwnd_, nullptr, TRUE);
    }
}

bool BreadcrumbBar::RegisterWindowClass() {
    WNDCLASSEXW window_class = {};
    window_class.cbSize = sizeof(window_class);
    window_class.style = CS_HREDRAW | CS_VREDRAW;
    window_class.lpfnWndProc = &BreadcrumbBar::WindowProc;
    window_class.hInstance = instance_;
    window_class.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    window_class.hbrBackground = nullptr;
    window_class.lpszClassName = kBreadcrumbClassName;

    const ATOM class_id = RegisterClassExW(&window_class);
    if (class_id == 0 && GetLastError() != ERROR_CLASS_ALREADY_EXISTS) {
        LogLastError(L"RegisterClassExW(FE_BreadcrumbBar)");
        return false;
    }
    return true;
}

bool BreadcrumbBar::CreateEditControl() {
    edit_hwnd_ = CreateWindowExW(
        0,
        WC_EDITW,
        L"",
        WS_CHILD | ES_AUTOHSCROLL,
        0,
        0,
        0,
        0,
        hwnd_,
        nullptr,
        instance_,
        nullptr);
    if (edit_hwnd_ == nullptr) {
        LogLastError(L"CreateWindowExW(BreadcrumbEdit)");
        return false;
    }

    if (!SetWindowSubclass(edit_hwnd_, &BreadcrumbBar::EditSubclassProc, kEditSubclassId, reinterpret_cast<DWORD_PTR>(this))) {
        LogLastError(L"SetWindowSubclass(BreadcrumbEdit)");
        return false;
    }

    const HRESULT set_theme_result = SetWindowTheme(edit_hwnd_, L"", L"");
    if (FAILED(set_theme_result)) {
        LogHResult(L"SetWindowTheme(BreadcrumbEdit)", set_theme_result);
    }

    const int margin = MulDiv(10, static_cast<int>(dpi_), 96);
    SendMessageW(
        edit_hwnd_,
        EM_SETMARGINS,
        EC_LEFTMARGIN | EC_RIGHTMARGIN,
        MAKELPARAM(margin, margin));

    edit_background_brush_.reset(CreateSolidBrush(kEditBackgroundColor));
    if (!edit_background_brush_) {
        LogLastError(L"CreateSolidBrush(BreadcrumbEditBackground)");
    }

    const HRESULT auto_complete_result =
        SHAutoComplete(edit_hwnd_, SHACF_FILESYS_ONLY | SHACF_AUTOSUGGEST_FORCE_ON | SHACF_AUTOAPPEND_FORCE_ON);
    if (FAILED(auto_complete_result)) {
        LogHResult(L"SHAutoComplete", auto_complete_result);
    }

    ShowWindow(edit_hwnd_, SW_HIDE);
    return true;
}

void BreadcrumbBar::LayoutEditControl() {
    if (edit_hwnd_ == nullptr || hwnd_ == nullptr) {
        return;
    }

    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        LogLastError(L"GetClientRect(BreadcrumbBar)");
        return;
    }

    const int margin = MulDiv(2, static_cast<int>(dpi_), 96);
    const int width = (client_rect.right - client_rect.left) - (margin * 2);
    const int height = (client_rect.bottom - client_rect.top) - (margin * 2);
    if (!MoveWindow(edit_hwnd_, margin, margin, (std::max)(0, width), (std::max)(0, height), TRUE)) {
        LogLastError(L"MoveWindow(BreadcrumbEdit)");
    }
}

void BreadcrumbBar::DrawControl(HDC hdc) {
    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        return;
    }

    FillRoundedRect(
        hdc,
        client_rect,
        kBackgroundColor,
        kBorderColor,
        MulDiv(6, static_cast<int>(dpi_), 96));

    const int requested_pixel_height = -MulDiv(9, static_cast<int>(dpi_), 72);
    if (font_ == nullptr || font_pixel_height_ != requested_pixel_height) {
        font_.reset(CreateFontW(
            requested_pixel_height,
            0,
            0,
            0,
            FW_NORMAL,
            FALSE,
            FALSE,
            FALSE,
            DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS,
            CLIP_DEFAULT_PRECIS,
            CLEARTYPE_QUALITY,
            DEFAULT_PITCH | FF_DONTCARE,
            L"Segoe UI Variable Text"));
        if (font_ == nullptr) {
            LogLastError(L"CreateFontW(Breadcrumb)");
        }
        font_pixel_height_ = requested_pixel_height;
    }

    HFONT old_font = nullptr;
    if (font_ != nullptr) {
        old_font = static_cast<HFONT>(SelectObject(hdc, font_.get()));
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, kTextColor);

    if (!edit_mode_) {
        RebuildSegments(hdc);
        TEXTMETRICW metrics = {};
        if (!GetTextMetricsW(hdc, &metrics)) {
            LogLastError(L"GetTextMetricsW(Breadcrumb)");
        }

        for (int i = 0; i < static_cast<int>(segments_.size()); ++i) {
            const SegmentLayout& segment = segments_[i];

            if (i == hovered_segment_index_) {
                FillRoundedRect(
                    hdc,
                    segment.segment_rect,
                    kHoverFillColor,
                    kHoverFillColor,
                    MulDiv(6, static_cast<int>(dpi_), 96));
            }

            const int segment_x = segment.segment_rect.left + MulDiv(8, static_cast<int>(dpi_), 96);
            const int segment_y = segment.segment_rect.top +
                ((segment.segment_rect.bottom - segment.segment_rect.top - metrics.tmHeight) / 2);
            TextOutW(
                hdc,
                segment_x,
                segment_y,
                segment.label.c_str(),
                static_cast<int>(segment.label.size()));

            if (i + 1 < static_cast<int>(segments_.size())) {
                if (i == hovered_separator_index_) {
                    FillRoundedRect(
                        hdc,
                        segment.separator_rect,
                        kHoverFillColor,
                        kHoverFillColor,
                        MulDiv(6, static_cast<int>(dpi_), 96));
                }

                SetTextColor(hdc, kSeparatorColor);
                const int sep_x = segment.separator_rect.left + MulDiv(7, static_cast<int>(dpi_), 96);
                const int sep_y = segment.separator_rect.top +
                    ((segment.separator_rect.bottom - segment.separator_rect.top - metrics.tmHeight) / 2);
                TextOutW(hdc, sep_x, sep_y, L"\u203A", 1);
                SetTextColor(hdc, kTextColor);
            }
        }
    }

    if (old_font != nullptr) {
        SelectObject(hdc, old_font);
    }
}

void BreadcrumbBar::EnsureSegmentsLayout() {
    if (hwnd_ == nullptr || edit_mode_) {
        return;
    }

    HDC hdc = GetDC(hwnd_);
    if (hdc == nullptr) {
        return;
    }

    const int requested_pixel_height = -MulDiv(9, static_cast<int>(dpi_), 72);
    if (font_ == nullptr || font_pixel_height_ != requested_pixel_height) {
        font_.reset(CreateFontW(
            requested_pixel_height,
            0,
            0,
            0,
            FW_NORMAL,
            FALSE,
            FALSE,
            FALSE,
            DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS,
            CLIP_DEFAULT_PRECIS,
            CLEARTYPE_QUALITY,
            DEFAULT_PITCH | FF_DONTCARE,
            L"Segoe UI Variable Text"));
        if (font_ == nullptr) {
            LogLastError(L"CreateFontW(BreadcrumbEnsureSegments)");
        }
        font_pixel_height_ = requested_pixel_height;
    }

    HFONT old_font = nullptr;
    if (font_ != nullptr) {
        old_font = static_cast<HFONT>(SelectObject(hdc, font_.get()));
    }

    RebuildSegments(hdc);

    if (old_font != nullptr) {
        SelectObject(hdc, old_font);
    }
    ReleaseDC(hwnd_, hdc);
}

void BreadcrumbBar::RebuildSegments(HDC hdc) {
    segments_ = BuildSegmentsFromPath(current_path_);
    if (segments_.empty()) {
        return;
    }

    RECT client_rect = {};
    if (!GetClientRect(hwnd_, &client_rect)) {
        LogLastError(L"GetClientRect(RebuildSegments)");
        segments_.clear();
        return;
    }

    const int margin = MulDiv(6, static_cast<int>(dpi_), 96);
    const int segment_padding = MulDiv(8, static_cast<int>(dpi_), 96);
    const int separator_padding = MulDiv(6, static_cast<int>(dpi_), 96);
    const int top = margin;
    const int bottom = (std::max)(top + 1, static_cast<int>(client_rect.bottom) - margin);

    const int separator_width = MeasureTextWidth(hdc, L"\u203A") + (separator_padding * 2);

    int x = margin;
    for (int i = 0; i < static_cast<int>(segments_.size()); ++i) {
        SegmentLayout& segment = segments_[i];
        const int text_width = MeasureTextWidth(hdc, segment.label);
        const int segment_width = text_width + (segment_padding * 2);

        segment.segment_rect = {
            x,
            top,
            x + segment_width,
            bottom,
        };
        x += segment_width;

        if (i + 1 < static_cast<int>(segments_.size())) {
            segment.separator_rect = {
                x,
                top,
                x + separator_width,
                bottom,
            };
            x += separator_width;
        } else {
            segment.separator_rect = {0, 0, 0, 0};
        }
    }
}

BreadcrumbBar::HitTestResult BreadcrumbBar::HitTest(POINT point) const {
    for (int i = 0; i < static_cast<int>(segments_.size()); ++i) {
        const SegmentLayout& segment = segments_[i];
        if (RectContainsPoint(segment.segment_rect, point)) {
            return {HitPart::Segment, i};
        }
        if (i + 1 < static_cast<int>(segments_.size()) && RectContainsPoint(segment.separator_rect, point)) {
            return {HitPart::Separator, i};
        }
    }
    return {};
}

void BreadcrumbBar::HandleMouseMove(POINT point) {
    EnsureSegmentsLayout();

    TRACKMOUSEEVENT track = {};
    track.cbSize = sizeof(track);
    track.dwFlags = TME_LEAVE;
    track.hwndTrack = hwnd_;
    TrackMouseEvent(&track);

    const HitTestResult hit = HitTest(point);
    const int new_hover_segment = (hit.part == HitPart::Segment) ? hit.index : -1;
    const int new_hover_separator = (hit.part == HitPart::Separator) ? hit.index : -1;

    if (new_hover_segment != hovered_segment_index_ || new_hover_separator != hovered_separator_index_) {
        hovered_segment_index_ = new_hover_segment;
        hovered_separator_index_ = new_hover_separator;
        Refresh();
    }
}

void BreadcrumbBar::HandleMouseLeave() {
    if (hovered_segment_index_ != -1 || hovered_separator_index_ != -1) {
        hovered_segment_index_ = -1;
        hovered_separator_index_ = -1;
        Refresh();
    }
}

void BreadcrumbBar::HandleLeftClick(POINT point) {
    if (edit_mode_) {
        return;
    }

    EnsureSegmentsLayout();

    const HitTestResult hit = HitTest(point);
    if (hit.part == HitPart::Segment && hit.index >= 0 && hit.index < static_cast<int>(segments_.size())) {
        PostNavigatePath(segments_[hit.index].path);
        return;
    }

    if (hit.part == HitPart::Separator && hit.index >= 0 && hit.index < static_cast<int>(segments_.size())) {
        POINT anchor_screen = point;
        ClientToScreen(hwnd_, &anchor_screen);
        StartSiblingFlyout(hit.index, anchor_screen);
        return;
    }

    ActivateEditMode();
}

void BreadcrumbBar::PostNavigatePath(const std::wstring& path) const {
    HWND target_hwnd = hwnd_ != nullptr ? GetAncestor(hwnd_, GA_ROOT) : nullptr;
    if (target_hwnd == nullptr) {
        target_hwnd = parent_hwnd_;
    }
    if (target_hwnd == nullptr) {
        return;
    }

    auto path_payload = std::make_unique<std::wstring>(NormalizePath(path));
    if (path_payload->empty()) {
        return;
    }

    if (!PostMessageW(
            target_hwnd,
            WM_FE_BREADCRUMB_NAVIGATE,
            0,
            reinterpret_cast<LPARAM>(path_payload.get()))) {
        LogLastError(L"PostMessageW(WM_FE_BREADCRUMB_NAVIGATE)");
        return;
    }

    path_payload.release();
}

void BreadcrumbBar::StartSiblingFlyout(int segment_index, POINT anchor_screen) {
    if (segment_index < 0 || segment_index >= static_cast<int>(segments_.size())) {
        return;
    }

    const std::wstring base_path = segments_[segment_index].path;
    if (base_path.empty()) {
        return;
    }

    const uint64_t generation = ++siblings_generation_;
    HWND target_hwnd = hwnd_;

    std::thread([target_hwnd, generation, base_path, anchor_screen]() {
        auto payload = std::make_unique<BreadcrumbBar::SiblingMenuPayload>();
        payload->generation = generation;
        payload->base_path = base_path;
        payload->anchor_screen = anchor_screen;

        std::wstring wildcard = NormalizePath(base_path);
        if (wildcard.empty()) {
            return;
        }
        if (wildcard.back() != L'\\') {
            wildcard.push_back(L'\\');
        }
        wildcard.append(L"*");

        WIN32_FIND_DATAW find_data = {};
        FindHandle find_handle{};
        find_handle.handle = FindFirstFileW(wildcard.c_str(), &find_data);
        if (find_handle.handle != INVALID_HANDLE_VALUE) {
            do {
                const bool is_directory = (find_data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                if (!is_directory) {
                    continue;
                }

                const wchar_t* name = find_data.cFileName;
                if (name[0] == L'.' && (name[1] == L'\0' || (name[1] == L'.' && name[2] == L'\0'))) {
                    continue;
                }

                payload->sibling_names.emplace_back(name);
            } while (FindNextFileW(find_handle.handle, &find_data));
        }

        std::sort(
            payload->sibling_names.begin(),
            payload->sibling_names.end(),
            [](const std::wstring& left, const std::wstring& right) {
                return CompareStringOrdinal(
                           left.c_str(),
                           static_cast<int>(left.size()),
                           right.c_str(),
                           static_cast<int>(right.size()),
                           TRUE) == CSTR_LESS_THAN;
            });

        if (!PostMessageW(target_hwnd, WM_FE_SIBLINGS_READY, 0, reinterpret_cast<LPARAM>(payload.get()))) {
            LogLastError(L"PostMessageW(BreadcrumbSiblingsReady)");
            return;
        }
        payload.release();
    }).detach();
}

void BreadcrumbBar::ShowSiblingFlyout(const SiblingMenuPayload& payload) {
    if (payload.generation != siblings_generation_.load()) {
        return;
    }

    HMENU menu = CreatePopupMenu();
    if (menu == nullptr) {
        LogLastError(L"CreatePopupMenu(BreadcrumbSibling)");
        return;
    }

    if (payload.sibling_names.empty()) {
        AppendMenuW(menu, MF_STRING | MF_GRAYED, kSiblingMenuIdBase, L"(No sibling folders)");
    } else {
        for (int i = 0; i < static_cast<int>(payload.sibling_names.size()); ++i) {
            const UINT menu_id = kSiblingMenuIdBase + static_cast<UINT>(i);
            AppendMenuW(menu, MF_STRING, menu_id, payload.sibling_names[i].c_str());
        }
    }

    const UINT selected = static_cast<UINT>(TrackPopupMenuEx(
        menu,
        TPM_RIGHTBUTTON | TPM_RETURNCMD,
        payload.anchor_screen.x,
        payload.anchor_screen.y,
        hwnd_,
        nullptr));

    DestroyMenu(menu);

    if (selected < kSiblingMenuIdBase || selected >= kSiblingMenuIdBase + payload.sibling_names.size()) {
        return;
    }

    const size_t sibling_index = static_cast<size_t>(selected - kSiblingMenuIdBase);
    const std::wstring target = JoinPath(payload.base_path, payload.sibling_names[sibling_index]);
    PostNavigatePath(target);
}

void BreadcrumbBar::CommitEditMode() {
    if (!edit_mode_ || edit_hwnd_ == nullptr) {
        return;
    }

    const int text_length = GetWindowTextLengthW(edit_hwnd_);
    std::wstring entered_path(static_cast<size_t>(text_length), L'\0');
    if (text_length > 0) {
        GetWindowTextW(edit_hwnd_, entered_path.data(), text_length + 1);
    }

    ExitEditMode(true);

    entered_path = NormalizePath(entered_path);
    if (!entered_path.empty()) {
        PostNavigatePath(entered_path);
    }
}

void BreadcrumbBar::ExitEditMode(bool return_focus_to_control) {
    if (!edit_mode_) {
        return;
    }

    edit_mode_ = false;
    if (edit_hwnd_ != nullptr) {
        ShowWindow(edit_hwnd_, SW_HIDE);
    }
    if (return_focus_to_control && hwnd_ != nullptr) {
        SetFocus(hwnd_);
    }
    Refresh();
}

std::vector<BreadcrumbBar::SegmentLayout> BreadcrumbBar::BuildSegmentsFromPath(const std::wstring& path) const {
    const std::wstring normalized = NormalizePath(path);
    if (normalized.empty()) {
        return {};
    }

    std::vector<SegmentLayout> built_segments;
    built_segments.reserve(16);

    if (normalized.size() >= 2 && normalized[1] == L':') {
        std::wstring cumulative = normalized.substr(0, 2);
        if (cumulative.size() == 2) {
            cumulative.push_back(L'\\');
        }
        built_segments.push_back(SegmentLayout{normalized.substr(0, 2), cumulative, {}, {}});

        size_t scan = (normalized.size() > 2 && normalized[2] == L'\\') ? 3 : 2;
        while (scan < normalized.size()) {
            size_t next_sep = normalized.find(L'\\', scan);
            const size_t token_len =
                (next_sep == std::wstring::npos) ? (normalized.size() - scan) : (next_sep - scan);
            if (token_len > 0) {
                const std::wstring token = normalized.substr(scan, token_len);
                cumulative = JoinPath(cumulative, token);
                built_segments.push_back(SegmentLayout{token, cumulative, {}, {}});
            }

            if (next_sep == std::wstring::npos) {
                break;
            }
            scan = next_sep + 1;
        }
        return built_segments;
    }

    size_t scan = 0;
    std::wstring cumulative;
    while (scan < normalized.size()) {
        size_t next_sep = normalized.find(L'\\', scan);
        const size_t token_len =
            (next_sep == std::wstring::npos) ? (normalized.size() - scan) : (next_sep - scan);
        if (token_len > 0) {
            const std::wstring token = normalized.substr(scan, token_len);
            cumulative = cumulative.empty() ? token : JoinPath(cumulative, token);
            built_segments.push_back(SegmentLayout{token, cumulative, {}, {}});
        }

        if (next_sep == std::wstring::npos) {
            break;
        }
        scan = next_sep + 1;
    }

    return built_segments;
}

int BreadcrumbBar::MeasureTextWidth(HDC hdc, const std::wstring& text) const {
    if (text.empty()) {
        return 0;
    }

    SIZE size = {};
    if (!GetTextExtentPoint32W(hdc, text.c_str(), static_cast<int>(text.size()), &size)) {
        LogLastError(L"GetTextExtentPoint32W(Breadcrumb)");
        return 0;
    }
    return size.cx;
}

LRESULT CALLBACK BreadcrumbBar::WindowProc(HWND hwnd, UINT message, WPARAM w_param, LPARAM l_param) {
    BreadcrumbBar* self = nullptr;
    if (message == WM_NCCREATE) {
        auto* create_struct = reinterpret_cast<CREATESTRUCTW*>(l_param);
        self = reinterpret_cast<BreadcrumbBar*>(create_struct->lpCreateParams);
        if (self == nullptr) {
            return FALSE;
        }
        self->hwnd_ = hwnd;
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
    } else {
        self = reinterpret_cast<BreadcrumbBar*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    }

    if (self != nullptr) {
        return self->HandleMessage(message, w_param, l_param);
    }
    return DefWindowProcW(hwnd, message, w_param, l_param);
}

LRESULT CALLBACK BreadcrumbBar::EditSubclassProc(
    HWND hwnd,
    UINT message,
    WPARAM w_param,
    LPARAM l_param,
    UINT_PTR subclass_id,
    DWORD_PTR ref_data) {
    (void)subclass_id;

    auto* self = reinterpret_cast<BreadcrumbBar*>(ref_data);
    if (self == nullptr) {
        return DefSubclassProc(hwnd, message, w_param, l_param);
    }

    switch (message) {
    case WM_KEYDOWN:
        if (w_param == VK_RETURN) {
            self->CommitEditMode();
            return 0;
        }
        if (w_param == VK_ESCAPE) {
            self->CancelEditMode();
            return 0;
        }
        break;

    case WM_KILLFOCUS:
        self->CancelEditMode();
        break;

    default:
        break;
    }

    return DefSubclassProc(hwnd, message, w_param, l_param);
}

LRESULT BreadcrumbBar::HandleMessage(UINT message, WPARAM w_param, LPARAM l_param) {
    switch (message) {
    case WM_CREATE:
        if (!CreateEditControl()) {
            return -1;
        }
        LayoutEditControl();
        return 0;

    case WM_SIZE:
        LayoutEditControl();
        Refresh();
        return 0;

    case WM_ERASEBKGND:
        return 1;

    case WM_PAINT: {
        PAINTSTRUCT paint = {};
        HDC hdc = BeginPaint(hwnd_, &paint);
        DrawControl(hdc);
        EndPaint(hwnd_, &paint);
        return 0;
    }

    case WM_MOUSEMOVE: {
        POINT point = {GET_X_LPARAM(l_param), GET_Y_LPARAM(l_param)};
        HandleMouseMove(point);
        return 0;
    }

    case WM_MOUSELEAVE:
        HandleMouseLeave();
        return 0;

    case WM_LBUTTONDOWN: {
        POINT point = {GET_X_LPARAM(l_param), GET_Y_LPARAM(l_param)};
        HandleLeftClick(point);
        return 0;
    }

    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORSTATIC:
        if (reinterpret_cast<HWND>(l_param) == edit_hwnd_) {
            const HDC edit_hdc = reinterpret_cast<HDC>(w_param);
            SetTextColor(edit_hdc, kEditTextColor);
            SetBkColor(edit_hdc, kEditBackgroundColor);
            SetBkMode(edit_hdc, OPAQUE);
            if (edit_background_brush_) {
                return reinterpret_cast<LRESULT>(edit_background_brush_.get());
            }
            return reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_WINDOW));
        }
        break;

    case WM_FE_SIBLINGS_READY: {
        auto payload = std::unique_ptr<SiblingMenuPayload>(reinterpret_cast<SiblingMenuPayload*>(l_param));
        if (payload) {
            ShowSiblingFlyout(*payload);
        }
        return 0;
    }

    case WM_NCDESTROY:
        ++siblings_generation_;
        SetWindowLongPtrW(hwnd_, GWLP_USERDATA, 0);
        hwnd_ = nullptr;
        break;

    default:
        break;
    }

    return DefWindowProcW(hwnd_, message, w_param, l_param);
}

}  // namespace fileexplorer
