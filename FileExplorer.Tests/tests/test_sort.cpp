#include <gtest/gtest.h>

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>

#include <algorithm>
#include <string>
#include <vector>

namespace {

enum class LocalSortColumn {
    Name,
    Extension,
    DateModified,
    Size,
};

enum class LocalSortDirection {
    Ascending,
    Descending,
};

enum class DateBucket {
    Today,
    Yesterday,
    ThisWeek,
    ThisMonth,
    Older,
};

struct SortEntry {
    std::wstring name;
    std::wstring extension;
    FILETIME modified_time{};
    ULONGLONG size_bytes = 0;
    bool is_folder = false;
};

int CompareTextInsensitive(const std::wstring& left, const std::wstring& right) {
    const int result = CompareStringOrdinal(
        left.c_str(),
        static_cast<int>(left.size()),
        right.c_str(),
        static_cast<int>(right.size()),
        TRUE);
    if (result == CSTR_LESS_THAN) {
        return -1;
    }
    if (result == CSTR_GREATER_THAN) {
        return 1;
    }
    return 0;
}

ULONGLONG FileTimeToUint64(const FILETIME& value) {
    ULARGE_INTEGER ui = {};
    ui.LowPart = value.dwLowDateTime;
    ui.HighPart = value.dwHighDateTime;
    return ui.QuadPart;
}

FILETIME Uint64ToFileTime(ULONGLONG value) {
    ULARGE_INTEGER ui = {};
    ui.QuadPart = value;
    FILETIME file_time = {};
    file_time.dwLowDateTime = ui.LowPart;
    file_time.dwHighDateTime = ui.HighPart;
    return file_time;
}

bool SameDate(const SYSTEMTIME& left, const SYSTEMTIME& right) {
    return left.wYear == right.wYear && left.wMonth == right.wMonth && left.wDay == right.wDay;
}

DateBucket BucketForFileTime(const FILETIME& file_time) {
    FILETIME file_local = {};
    if (!FileTimeToLocalFileTime(&file_time, &file_local)) {
        return DateBucket::Older;
    }

    FILETIME now_utc = {};
    GetSystemTimeAsFileTime(&now_utc);
    FILETIME now_local = {};
    if (!FileTimeToLocalFileTime(&now_utc, &now_local)) {
        return DateBucket::Older;
    }

    SYSTEMTIME file_st = {};
    SYSTEMTIME now_st = {};
    if (!FileTimeToSystemTime(&file_local, &file_st) || !FileTimeToSystemTime(&now_local, &now_st)) {
        return DateBucket::Older;
    }

    if (SameDate(file_st, now_st)) {
        return DateBucket::Today;
    }

    constexpr ULONGLONG kTicksPerDay = 24ULL * 60ULL * 60ULL * 10000000ULL;
    const ULONGLONG now_ticks = FileTimeToUint64(now_local);
    const ULONGLONG file_ticks = FileTimeToUint64(file_local);

    if (now_ticks >= kTicksPerDay) {
        const FILETIME yesterday_ft = Uint64ToFileTime(now_ticks - kTicksPerDay);
        SYSTEMTIME yesterday_st = {};
        if (FileTimeToSystemTime(&yesterday_ft, &yesterday_st) && SameDate(file_st, yesterday_st)) {
            return DateBucket::Yesterday;
        }
    }

    if (now_ticks >= file_ticks && (now_ticks - file_ticks) <= (7ULL * kTicksPerDay)) {
        return DateBucket::ThisWeek;
    }

    if (file_st.wYear == now_st.wYear && file_st.wMonth == now_st.wMonth) {
        return DateBucket::ThisMonth;
    }

    return DateBucket::Older;
}

FILETIME FileTimeDaysAgo(int days_ago) {
    FILETIME now_utc = {};
    GetSystemTimeAsFileTime(&now_utc);
    const ULONGLONG now_ticks = FileTimeToUint64(now_utc);
    constexpr ULONGLONG kTicksPerDay = 24ULL * 60ULL * 60ULL * 10000000ULL;
    const ULONGLONG offset = static_cast<ULONGLONG>((std::max)(0, days_ago)) * kTicksPerDay;
    return Uint64ToFileTime(now_ticks - offset);
}

void SortEntries(
    std::vector<SortEntry>* entries,
    LocalSortColumn column,
    LocalSortDirection direction) {
    if (entries == nullptr) {
        return;
    }

    std::sort(
        entries->begin(),
        entries->end(),
        [column, direction](const SortEntry& left, const SortEntry& right) {
            const bool enforce_folders_first = column != LocalSortColumn::DateModified;
            if (enforce_folders_first && left.is_folder != right.is_folder) {
                return left.is_folder && !right.is_folder;
            }

            int compare = 0;
            switch (column) {
            case LocalSortColumn::Name:
                compare = CompareTextInsensitive(left.name, right.name);
                break;
            case LocalSortColumn::Extension:
                compare = CompareTextInsensitive(left.extension, right.extension);
                if (compare == 0) {
                    compare = CompareTextInsensitive(left.name, right.name);
                }
                break;
            case LocalSortColumn::DateModified: {
                const ULONGLONG left_time = FileTimeToUint64(left.modified_time);
                const ULONGLONG right_time = FileTimeToUint64(right.modified_time);
                compare = (left_time < right_time) ? -1 : (left_time > right_time ? 1 : 0);
                break;
            }
            case LocalSortColumn::Size:
                compare = (left.size_bytes < right.size_bytes) ? -1 : (left.size_bytes > right.size_bytes ? 1 : 0);
                break;
            }

            if (compare == 0) {
                compare = CompareTextInsensitive(left.name, right.name);
            }

            if (direction == LocalSortDirection::Descending) {
                compare = -compare;
            }
            return compare < 0;
        });
}

}  // namespace

namespace fileexplorer::tests {

TEST(Sort, SortByNameAscending) {
    std::vector<SortEntry> entries = {
        {L"zeta.txt", L".txt", {}, 20, false},
        {L"Alpha.txt", L".txt", {}, 10, false},
        {L"beta.txt", L".txt", {}, 15, false},
    };

    SortEntries(&entries, LocalSortColumn::Name, LocalSortDirection::Ascending);
    EXPECT_TRUE(entries[0].name == L"Alpha.txt");
    EXPECT_TRUE(entries[1].name == L"beta.txt");
    EXPECT_TRUE(entries[2].name == L"zeta.txt");
}

TEST(Sort, SortByNameDescending) {
    std::vector<SortEntry> entries = {
        {L"alpha.txt", L".txt", {}, 20, false},
        {L"Gamma.txt", L".txt", {}, 10, false},
        {L"beta.txt", L".txt", {}, 15, false},
    };

    SortEntries(&entries, LocalSortColumn::Name, LocalSortDirection::Descending);
    EXPECT_TRUE(entries[0].name == L"Gamma.txt");
    EXPECT_TRUE(entries[1].name == L"beta.txt");
    EXPECT_TRUE(entries[2].name == L"alpha.txt");
}

TEST(Sort, SortByExtension) {
    std::vector<SortEntry> entries = {
        {L"alpha.docx", L".docx", {}, 20, false},
        {L"beta.txt", L".txt", {}, 10, false},
        {L"gamma.pdf", L".pdf", {}, 15, false},
    };

    SortEntries(&entries, LocalSortColumn::Extension, LocalSortDirection::Ascending);
    EXPECT_TRUE(entries[0].extension == L".docx");
    EXPECT_TRUE(entries[1].extension == L".pdf");
    EXPECT_TRUE(entries[2].extension == L".txt");
}

TEST(Sort, SortByDate) {
    std::vector<SortEntry> entries = {
        {L"new.txt", L".txt", FileTimeDaysAgo(0), 1, false},
        {L"old.txt", L".txt", FileTimeDaysAgo(7), 1, false},
        {L"mid.txt", L".txt", FileTimeDaysAgo(2), 1, false},
    };

    SortEntries(&entries, LocalSortColumn::DateModified, LocalSortDirection::Descending);
    EXPECT_TRUE(entries[0].name == L"new.txt");
    EXPECT_TRUE(entries[1].name == L"mid.txt");
    EXPECT_TRUE(entries[2].name == L"old.txt");
}

TEST(Sort, SortBySize) {
    std::vector<SortEntry> entries = {
        {L"large.bin", L".bin", {}, 900, false},
        {L"small.bin", L".bin", {}, 100, false},
        {L"medium.bin", L".bin", {}, 500, false},
    };

    SortEntries(&entries, LocalSortColumn::Size, LocalSortDirection::Ascending);
    EXPECT_TRUE(entries[0].name == L"small.bin");
    EXPECT_TRUE(entries[1].name == L"medium.bin");
    EXPECT_TRUE(entries[2].name == L"large.bin");
}

TEST(Sort, FolderAlwaysFirst) {
    std::vector<SortEntry> entries = {
        {L"z_file.txt", L".txt", {}, 1, false},
        {L"A_Folder", L"", {}, 0, true},
        {L"a_file.txt", L".txt", {}, 2, false},
    };

    SortEntries(&entries, LocalSortColumn::Name, LocalSortDirection::Descending);
    EXPECT_TRUE(entries[0].is_folder);
    EXPECT_TRUE(entries[0].name == L"A_Folder");
}

TEST(Sort, DateModifiedTreatsFoldersLikeFiles) {
    std::vector<SortEntry> entries = {
        {L"folder_old", L"", FileTimeDaysAgo(10), 0, true},
        {L"file_new.txt", L".txt", FileTimeDaysAgo(0), 100, false},
        {L"folder_mid", L"", FileTimeDaysAgo(3), 0, true},
        {L"file_older.txt", L".txt", FileTimeDaysAgo(7), 50, false},
    };

    SortEntries(&entries, LocalSortColumn::DateModified, LocalSortDirection::Descending);
    EXPECT_TRUE(entries[0].name == L"file_new.txt");
    EXPECT_TRUE(entries[1].name == L"folder_mid");
    EXPECT_TRUE(entries[2].name == L"file_older.txt");
    EXPECT_TRUE(entries[3].name == L"folder_old");
}

TEST(Sort, DateGroupBuckets) {
    const FILETIME today = FileTimeDaysAgo(0);
    const FILETIME yesterday = FileTimeDaysAgo(1);
    const FILETIME this_week = FileTimeDaysAgo(3);
    const FILETIME this_month = FileTimeDaysAgo(12);
    const FILETIME older = FileTimeDaysAgo(50);

    EXPECT_TRUE(BucketForFileTime(today) == DateBucket::Today);
    EXPECT_TRUE(BucketForFileTime(yesterday) == DateBucket::Yesterday);
    EXPECT_TRUE(BucketForFileTime(this_week) == DateBucket::ThisWeek);
    EXPECT_TRUE(BucketForFileTime(this_month) == DateBucket::ThisMonth);
    EXPECT_TRUE(BucketForFileTime(older) == DateBucket::Older);
}

TEST(Sort, MixedCaseNames) {
    std::vector<SortEntry> entries = {
        {L"delta.txt", L".txt", {}, 1, false},
        {L"Bravo.txt", L".txt", {}, 1, false},
        {L"alpha.txt", L".txt", {}, 1, false},
    };

    SortEntries(&entries, LocalSortColumn::Name, LocalSortDirection::Ascending);
    EXPECT_TRUE(entries[0].name == L"alpha.txt");
    EXPECT_TRUE(entries[1].name == L"Bravo.txt");
    EXPECT_TRUE(entries[2].name == L"delta.txt");
}

}  // namespace fileexplorer::tests
