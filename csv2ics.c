/*============================================================================
 *  CSV2ICS - CSV to ICS Calendar Converter
 *  A Petzold-style Win32 C application
 *  Single-file, no external dependencies
 *============================================================================*/

#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif
#define WIN32_LEAN_AND_MEAN
#define _CRT_SECURE_NO_WARNINGS

#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <shlobj.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>
#include <time.h>
#include <stdbool.h>

#include <wincrypt.h>

#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")

#pragma comment(linker, "\"/manifestdependency:type='win32' \
name='Microsoft.Windows.Common-Controls' version='6.0.0.0' \
processorArchitecture='*' publicKeyToken='6595b64144ccf1df' \
language='*'\"")

/*============================================================================
 *  CONSTANTS AND CONTROL IDS
 *============================================================================*/

#define APP_NAME        L"CSV to ICS Converter"
#define APP_CLASS       L"CSV2ICSWindowClass"

#define CLIENT_WIDTH    820
#define CLIENT_HEIGHT   720
#define NAV_Y           (CLIENT_HEIGHT - 50)

/* Control IDs */
#define IDC_BTN_OPEN        1001
#define IDC_LBL_PATH        1002
#define IDC_CHK_HEADER      1003
#define IDC_RADIO_MDY       1004
#define IDC_RADIO_DMY       1005
#define IDC_RADIO_ISO       1006
#define IDC_LISTVIEW        1006
#define IDC_BTN_NEXT        1010
#define IDC_BTN_BACK        1011

/* Page 2 - Mapping combos */
#define IDC_COMBO_DTSTART   1100
#define IDC_COMBO_DTEND     1101
#define IDC_COMBO_SUMMARY   1102
#define IDC_COMBO_DESC      1103
#define IDC_COMBO_LOCATION  1104
#define IDC_COMBO_URL       1105
#define IDC_COMBO_CATEGORIES 1106
#define IDC_COMBO_STATUS    1107
#define IDC_COMBO_TRANSP    1108
#define IDC_COMBO_PRIORITY  1109
#define IDC_COMBO_STARTTIME 1110
#define IDC_COMBO_ENDTIME   1111
#define IDC_COMBO_ALLDAY    1112
#define IDC_COMBO_REMINDER  1113
#define IDC_COMBO_RECURRENCE 1114
#define IDC_PREVIEW         1120

/* Page 3 - Export */
#define IDC_RADIO_SINGLE    1200
#define IDC_RADIO_SEPARATE  1201
#define IDC_LBL_COUNT       1202
#define IDC_BTN_EXPORT      1203
#define IDC_LBL_STATUS      1204
#define IDC_COMBO_REMIND1   1210
#define IDC_COMBO_REMIND2   1211
#define IDC_CHK_TRAVEL      1212
#define IDC_BTN_STARTOVER   1213
#define IDM_ABOUT           9001

/* Pages */
enum { PAGE_FILE = 0, PAGE_MAP = 1, PAGE_EXPORT = 2 };

/* ICS field indices */
enum {
    ICS_DTSTART = 0, ICS_DTEND, ICS_SUMMARY, ICS_DESCRIPTION,
    ICS_LOCATION, ICS_URL, ICS_CATEGORIES, ICS_STATUS,
    ICS_TRANSP, ICS_PRIORITY,
    ICS_START_TIME, ICS_END_TIME, ICS_ALL_DAY, ICS_REMINDER, ICS_RECURRENCE,
    ICS_FIELD_COUNT
};

static const wchar_t* ICS_FIELD_NAMES[] = {
    L"Start Date (required)", L"End Date", L"SUMMARY", L"DESCRIPTION",
    L"LOCATION", L"URL", L"CATEGORIES", L"STATUS",
    L"TRANSP", L"PRIORITY",
    L"Start Time", L"End Time", L"All Day", L"Reminder", L"Recurrence"
};

static const int COMBO_IDS[] = {
    IDC_COMBO_DTSTART, IDC_COMBO_DTEND, IDC_COMBO_SUMMARY, IDC_COMBO_DESC,
    IDC_COMBO_LOCATION, IDC_COMBO_URL, IDC_COMBO_CATEGORIES, IDC_COMBO_STATUS,
    IDC_COMBO_TRANSP, IDC_COMBO_PRIORITY,
    IDC_COMBO_STARTTIME, IDC_COMBO_ENDTIME, IDC_COMBO_ALLDAY, IDC_COMBO_REMINDER,
    IDC_COMBO_RECURRENCE
};

/* Reminder duration options (in minutes), -1 = none */
static const int REMINDER_VALUES[] = { -1, 0, 5, 10, 15, 30, 60, 120, 1440, 2880 };
static const wchar_t* REMINDER_LABELS[] = {
    L"None", L"At time of event", L"5 minutes before", L"10 minutes before",
    L"15 minutes before", L"30 minutes before", L"1 hour before",
    L"2 hours before", L"1 day before", L"2 days before"
};
#define REMINDER_OPTION_COUNT 10

/*============================================================================
 *  DATA STRUCTURES
 *============================================================================*/

#define MAX_COLUMNS     64
#define MAX_FIELD_LEN   4096
#define MAX_ROWS        100000

typedef struct {
    int year, month, day;
    int hour, minute, second;
    bool has_time;
    bool valid;
} ParsedDateTime;

typedef struct {
    wchar_t* fields[MAX_COLUMNS];
    int field_count;
} CsvRow;

typedef struct {
    CsvRow   header;
    CsvRow*  rows;
    int      row_count;
    int      col_count;
    bool     has_header;
} CsvData;

typedef struct {
    HWND hwndMain;
    HFONT hFont;

    /* Page 1 controls */
    HWND hwndBtnOpen, hwndLblPath, hwndChkHeader;
    HWND hwndRadioMDY, hwndRadioDMY, hwndRadioISO;
    HWND hwndListView;
    HWND hwndLblPage1Title;
    HWND hwndLblDateFmt;

    /* Page 2 controls */
    HWND hwndLblMapTitle;
    HWND hwndComboLabel[ICS_FIELD_COUNT];
    HWND hwndCombo[ICS_FIELD_COUNT];
    HWND hwndLblPreview;
    HWND hwndPreview;

    /* Page 3 controls */
    HWND hwndLblExportTitle;
    HWND hwndRadioSingle, hwndRadioSeparate;
    HWND hwndLblCount, hwndBtnExport, hwndLblStatus;
    HWND hwndLblRemindSection;
    HWND hwndLblRemind1, hwndComboRemind1;
    HWND hwndLblRemind2, hwndComboRemind2;
    HWND hwndChkTravel;
    HWND hwndBtnStartOver;

    /* Navigation */
    HWND hwndBtnBack, hwndBtnNext;
    int  currentPage;

    /* Data */
    CsvData  csv;
    int      fieldMapping[ICS_FIELD_COUNT];
    wchar_t  csvFilePath[MAX_PATH];
    int      dateFormatPref; /* 0=MDY, 1=DMY */
    int      defaultReminder1; /* index into REMINDER_VALUES */
    int      defaultReminder2;
    bool     appleTravel;
} AppState;

/*============================================================================
 *  FORWARD DECLARATIONS
 *============================================================================*/

static LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);

/* CSV */
static bool     CsvLoad(AppState* state, const wchar_t* filePath);
static int      CsvParseRow(const wchar_t* data, int* pos, int dataLen, wchar_t** fields, int maxFields);
static bool     CsvDetectHeader(CsvData* csv);
static void     CsvFree(CsvData* csv);

/* Date/Time */
static ParsedDateTime ParseDateTime(const wchar_t* input, int pref);

/* ICS */
static void     IcsEscapeText(const wchar_t* input, wchar_t* output, int outLen);
static void     IcsWriteFolded(FILE* f, const char* line);
static void     NextDay(ParsedDateTime* dt);
static int      ParseReminderMinutes(const wchar_t* input);
static const char* ParseRecurrence(const wchar_t* input);
static bool     IcsWriteEvent(FILE* f, CsvRow* row, AppState* state, int eventIndex);
static bool     IcsExportSingle(const wchar_t* path, AppState* state);
static bool     IcsExportSeparate(const wchar_t* folder, AppState* state);

/* UI */
static void     CreatePage1Controls(AppState* state, HWND parent);
static void     CreatePage2Controls(AppState* state, HWND parent);
static void     CreatePage3Controls(AppState* state, HWND parent);
static void     ShowPage(AppState* state, int page);
static void     PopulateListView(AppState* state);
static void     PopulateMappingCombos(AppState* state);
static void     UpdateIcsPreview(AppState* state);
static void     DoOpenFile(AppState* state);
static void     DoExport(AppState* state);
static void     AutoMapFields(AppState* state);

/*============================================================================
 *  CSV PARSER
 *============================================================================*/

static wchar_t* WstrDup(const wchar_t* s) {
    size_t len = wcslen(s) + 1;
    wchar_t* dup = (wchar_t*)malloc(len * sizeof(wchar_t));
    if (dup) wcscpy(dup, s);
    return dup;
}

static bool CsvLoad(AppState* state, const wchar_t* filePath) {
    CsvFree(&state->csv);

    HANDLE hFile = CreateFile(filePath, GENERIC_READ, FILE_SHARE_READ, NULL,
                              OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile == INVALID_HANDLE_VALUE) return false;

    DWORD fileSize = GetFileSize(hFile, NULL);
    if (fileSize == INVALID_FILE_SIZE || fileSize == 0) {
        CloseHandle(hFile);
        return false;
    }

    char* rawData = (char*)malloc(fileSize + 1);
    if (!rawData) { CloseHandle(hFile); return false; }

    DWORD bytesRead;
    ReadFile(hFile, rawData, fileSize, &bytesRead, NULL);
    CloseHandle(hFile);
    rawData[bytesRead] = '\0';

    /* Detect BOM and convert to wide char */
    wchar_t* wideData = NULL;
    int wideLen = 0;
    int offset = 0;

    if (bytesRead >= 3 && (unsigned char)rawData[0] == 0xEF &&
        (unsigned char)rawData[1] == 0xBB && (unsigned char)rawData[2] == 0xBF) {
        offset = 3; /* UTF-8 BOM */
    }
    else if (bytesRead >= 2 && (unsigned char)rawData[0] == 0xFF &&
             (unsigned char)rawData[1] == 0xFE) {
        /* UTF-16 LE BOM */
        wideLen = (int)((bytesRead - 2) / sizeof(wchar_t));
        wideData = (wchar_t*)malloc((wideLen + 1) * sizeof(wchar_t));
        if (wideData) {
            memcpy(wideData, rawData + 2, wideLen * sizeof(wchar_t));
            wideData[wideLen] = L'\0';
        }
        free(rawData);
        if (!wideData) return false;
        goto parse;
    }

    wideLen = MultiByteToWideChar(CP_UTF8, 0, rawData + offset, bytesRead - offset, NULL, 0);
    wideData = (wchar_t*)malloc((wideLen + 1) * sizeof(wchar_t));
    if (!wideData) { free(rawData); return false; }
    MultiByteToWideChar(CP_UTF8, 0, rawData + offset, bytesRead - offset, wideData, wideLen);
    wideData[wideLen] = L'\0';
    free(rawData);

parse:;
    /* Parse CSV rows */
    state->csv.rows = (CsvRow*)calloc(MAX_ROWS, sizeof(CsvRow));
    if (!state->csv.rows) { free(wideData); return false; }

    int pos = 0;
    int rowCount = 0;
    wchar_t* fields[MAX_COLUMNS];

    /* Parse first row (potential header) */
    int firstRowCols = CsvParseRow(wideData, &pos, wideLen, fields, MAX_COLUMNS);
    if (firstRowCols <= 0) { free(wideData); return false; }

    state->csv.col_count = firstRowCols;
    state->csv.header.field_count = firstRowCols;
    for (int i = 0; i < firstRowCols; i++) {
        state->csv.header.fields[i] = fields[i];
    }

    /* Parse remaining rows */
    while (pos < wideLen && rowCount < MAX_ROWS) {
        int cols = CsvParseRow(wideData, &pos, wideLen, fields, MAX_COLUMNS);
        if (cols <= 0) break;

        CsvRow* row = &state->csv.rows[rowCount];
        row->field_count = cols;
        for (int i = 0; i < cols && i < MAX_COLUMNS; i++) {
            row->fields[i] = fields[i];
        }
        rowCount++;
    }

    state->csv.row_count = rowCount;
    state->csv.has_header = CsvDetectHeader(&state->csv);

    free(wideData);
    return true;
}

static int CsvParseRow(const wchar_t* data, int* pos, int dataLen, wchar_t** fields, int maxFields) {
    int fieldCount = 0;
    wchar_t buf[MAX_FIELD_LEN];
    int bufPos = 0;

    if (*pos >= dataLen) return 0;

    /* Skip blank lines */
    while (*pos < dataLen && (data[*pos] == L'\r' || data[*pos] == L'\n')) {
        (*pos)++;
    }
    if (*pos >= dataLen) return 0;

    enum { FIELD_START, IN_UNQUOTED, IN_QUOTED, QUOTE_END } parseState = FIELD_START;

    while (*pos <= dataLen && fieldCount < maxFields) {
        wchar_t ch = (*pos < dataLen) ? data[*pos] : L'\0';

        switch (parseState) {
        case FIELD_START:
            bufPos = 0;
            if (ch == L'"') {
                parseState = IN_QUOTED;
                (*pos)++;
            } else if (ch == L',' ) {
                buf[0] = L'\0';
                fields[fieldCount++] = WstrDup(buf);
                (*pos)++;
            } else if (ch == L'\r' || ch == L'\n' || ch == L'\0') {
                buf[0] = L'\0';
                if (fieldCount > 0 || bufPos > 0)
                    fields[fieldCount++] = WstrDup(buf);
                if (ch == L'\r' && *pos + 1 < dataLen && data[*pos + 1] == L'\n')
                    (*pos)++;
                (*pos)++;
                return fieldCount;
            } else {
                buf[bufPos++] = ch;
                parseState = IN_UNQUOTED;
                (*pos)++;
            }
            break;

        case IN_UNQUOTED:
            if (ch == L',' || ch == L'\r' || ch == L'\n' || ch == L'\0') {
                buf[bufPos] = L'\0';
                fields[fieldCount++] = WstrDup(buf);
                if (ch == L',') {
                    parseState = FIELD_START;
                    (*pos)++;
                } else {
                    if (ch == L'\r' && *pos + 1 < dataLen && data[*pos + 1] == L'\n')
                        (*pos)++;
                    (*pos)++;
                    return fieldCount;
                }
            } else {
                if (bufPos < MAX_FIELD_LEN - 1)
                    buf[bufPos++] = ch;
                (*pos)++;
            }
            break;

        case IN_QUOTED:
            if (ch == L'"') {
                if (*pos + 1 < dataLen && data[*pos + 1] == L'"') {
                    if (bufPos < MAX_FIELD_LEN - 1)
                        buf[bufPos++] = L'"';
                    (*pos) += 2;
                } else {
                    parseState = QUOTE_END;
                    (*pos)++;
                }
            } else if (ch == L'\0') {
                buf[bufPos] = L'\0';
                fields[fieldCount++] = WstrDup(buf);
                return fieldCount;
            } else {
                if (bufPos < MAX_FIELD_LEN - 1)
                    buf[bufPos++] = ch;
                (*pos)++;
            }
            break;

        case QUOTE_END:
            buf[bufPos] = L'\0';
            fields[fieldCount++] = WstrDup(buf);
            if (ch == L',') {
                parseState = FIELD_START;
                (*pos)++;
            } else {
                if (ch == L'\r' && *pos + 1 < dataLen && data[*pos + 1] == L'\n')
                    (*pos)++;
                (*pos)++;
                return fieldCount;
            }
            break;
        }
    }

    /* End of data with field in progress */
    if (parseState == IN_UNQUOTED || parseState == IN_QUOTED) {
        buf[bufPos] = L'\0';
        fields[fieldCount++] = WstrDup(buf);
    }

    return fieldCount;
}

static bool CsvDetectHeader(CsvData* csv) {
    if (csv->col_count == 0) return false;

    /* Heuristic: if any header field parses as a valid date, probably not a header */
    for (int i = 0; i < csv->header.field_count; i++) {
        ParsedDateTime dt = ParseDateTime(csv->header.fields[i], 0);
        if (dt.valid) return false;
    }

    /* Check if header fields look like labels (contain letters, not pure numbers) */
    int labelCount = 0;
    for (int i = 0; i < csv->header.field_count; i++) {
        const wchar_t* f = csv->header.fields[i];
        bool hasAlpha = false;
        for (int j = 0; f[j]; j++) {
            if (iswalpha(f[j])) { hasAlpha = true; break; }
        }
        if (hasAlpha) labelCount++;
    }

    return labelCount > csv->header.field_count / 2;
}

static void CsvFree(CsvData* csv) {
    for (int i = 0; i < csv->header.field_count; i++) {
        free(csv->header.fields[i]);
        csv->header.fields[i] = NULL;
    }
    csv->header.field_count = 0;

    if (csv->rows) {
        for (int r = 0; r < csv->row_count; r++) {
            for (int c = 0; c < csv->rows[r].field_count; c++) {
                free(csv->rows[r].fields[c]);
            }
        }
        free(csv->rows);
        csv->rows = NULL;
    }
    csv->row_count = 0;
    csv->col_count = 0;
}

/*============================================================================
 *  DATE/TIME PARSER
 *============================================================================*/

static bool IsDigits(const wchar_t* s, int count) {
    for (int i = 0; i < count; i++) {
        if (!iswdigit(s[i])) return false;
    }
    return true;
}

static int WtoI(const wchar_t* s, int count) {
    int val = 0;
    for (int i = 0; i < count; i++) {
        val = val * 10 + (s[i] - L'0');
    }
    return val;
}

static bool ValidateDate(int y, int m, int d) {
    if (y < 1900 || y > 2100 || m < 1 || m > 12 || d < 1 || d > 31) return false;
    int daysInMonth[] = {0,31,28,31,30,31,30,31,31,30,31,30,31};
    if ((y % 4 == 0 && y % 100 != 0) || y % 400 == 0) daysInMonth[2] = 29;
    return d <= daysInMonth[m];
}

static bool ValidateTime(int h, int m, int s) {
    return h >= 0 && h <= 23 && m >= 0 && m <= 59 && s >= 0 && s <= 59;
}

static bool TryParseTime(const wchar_t* s, int* h, int* m, int* sec) {
    *h = 0; *m = 0; *sec = 0;

    /* Skip leading whitespace */
    while (*s == L' ') s++;
    if (!*s) return false;

    /* HH:MM or HH:MM:SS, optional AM/PM */
    int hour = 0, minute = 0, second = 0;
    int consumed = 0;

    if (swscanf(s, L"%d:%d:%d%n", &hour, &minute, &second, &consumed) >= 3) {
        s += consumed;
    } else if (swscanf(s, L"%d:%d%n", &hour, &minute, &consumed) >= 2) {
        second = 0;
        s += consumed;
    } else {
        return false;
    }

    /* Check for AM/PM */
    while (*s == L' ') s++;
    if (_wcsnicmp(s, L"PM", 2) == 0 || _wcsnicmp(s, L"pm", 2) == 0) {
        if (hour != 12) hour += 12;
    } else if (_wcsnicmp(s, L"AM", 2) == 0 || _wcsnicmp(s, L"am", 2) == 0) {
        if (hour == 12) hour = 0;
    }

    if (!ValidateTime(hour, minute, second)) return false;

    *h = hour; *m = minute; *sec = second;
    return true;
}

static ParsedDateTime ParseDateTime(const wchar_t* input, int pref) {
    ParsedDateTime result = {0};
    if (!input || !*input) return result;

    /* Trim whitespace */
    while (*input == L' ') input++;
    wchar_t trimmed[512];
    wcsncpy(trimmed, input, 511);
    trimmed[511] = L'\0';
    int len = (int)wcslen(trimmed);
    while (len > 0 && (trimmed[len-1] == L' ' || trimmed[len-1] == L'\r' || trimmed[len-1] == L'\n')) {
        trimmed[--len] = L'\0';
    }

    if (len == 0) return result;

    /* Try ISO 8601: YYYY-MM-DDTHH:MM:SS or YYYY-MM-DD HH:MM:SS */
    if (len >= 10 && IsDigits(trimmed, 4) && trimmed[4] == L'-' &&
        IsDigits(trimmed + 5, 2) && trimmed[7] == L'-' && IsDigits(trimmed + 8, 2)) {
        result.year = WtoI(trimmed, 4);
        result.month = WtoI(trimmed + 5, 2);
        result.day = WtoI(trimmed + 8, 2);
        if (ValidateDate(result.year, result.month, result.day)) {
            result.valid = true;
            if (len > 10 && (trimmed[10] == L'T' || trimmed[10] == L' ')) {
                int h, m, s;
                if (TryParseTime(trimmed + 11, &h, &m, &s)) {
                    result.hour = h; result.minute = m; result.second = s;
                    result.has_time = true;
                }
            }
            return result;
        }
    }

    /* Try compact: YYYYMMDD or YYYYMMDDTHHMMSS */
    if (len >= 8 && IsDigits(trimmed, 8)) {
        int y = WtoI(trimmed, 4);
        int m = WtoI(trimmed + 4, 2);
        int d = WtoI(trimmed + 6, 2);
        if (ValidateDate(y, m, d)) {
            result.year = y; result.month = m; result.day = d;
            result.valid = true;
            if (len >= 15 && trimmed[8] == L'T' && IsDigits(trimmed + 9, 6)) {
                result.hour = WtoI(trimmed + 9, 2);
                result.minute = WtoI(trimmed + 11, 2);
                result.second = WtoI(trimmed + 13, 2);
                if (ValidateTime(result.hour, result.minute, result.second))
                    result.has_time = true;
            }
            return result;
        }
    }

    /* Try MM/DD/YYYY or DD/MM/YYYY with optional time */
    {
        int a, b, c;
        int consumed = 0;
        wchar_t sep;
        if (swscanf(trimmed, L"%d%lc%d%lc%d%n", &a, &sep, &b, &sep, &c, &consumed) == 5 &&
            (sep == L'/' || sep == L'-' || sep == L'.')) {
            int year, month, day;

            if (c > 100) {
                /* a/b/YYYY */
                year = c;
                if (pref == 1) { day = a; month = b; }
                else { month = a; day = b; }
            } else if (a > 100) {
                /* YYYY/m/d */
                year = a; month = b; day = c;
            } else {
                /* Two-digit year? a/b/cc */
                year = c < 100 ? (c + 2000) : c;
                if (pref == 1) { day = a; month = b; }
                else { month = a; day = b; }
            }

            if (ValidateDate(year, month, day)) {
                result.year = year; result.month = month; result.day = day;
                result.valid = true;
                if (consumed < len) {
                    int h, m, s;
                    if (TryParseTime(trimmed + consumed, &h, &m, &s)) {
                        result.hour = h; result.minute = m; result.second = s;
                        result.has_time = true;
                    }
                }
                return result;
            }
        }
    }

    /* Try long date: Month DD, YYYY or DD Month YYYY */
    {
        static const wchar_t* months[] = {
            L"january", L"february", L"march", L"april", L"may", L"june",
            L"july", L"august", L"september", L"october", L"november", L"december"
        };
        static const wchar_t* monthsShort[] = {
            L"jan", L"feb", L"mar", L"apr", L"may", L"jun",
            L"jul", L"aug", L"sep", L"oct", L"nov", L"dec"
        };

        wchar_t lower[512];
        for (int i = 0; i < len && i < 511; i++) lower[i] = towlower(trimmed[i]);
        lower[len < 511 ? len : 511] = L'\0';

        for (int mi = 0; mi < 12; mi++) {
            const wchar_t* mname = months[mi];
            const wchar_t* mshort = monthsShort[mi];
            wchar_t* found = wcsstr(lower, mname);
            int mlen = (int)wcslen(mname);
            if (!found) {
                found = wcsstr(lower, mshort);
                mlen = (int)wcslen(mshort);
            }
            if (found) {
                int monthNum = mi + 1;
                /* Extract numbers around the month name */
                int nums[3] = {0};
                int numCount = 0;
                const wchar_t* p = lower;
                while (*p && numCount < 3) {
                    if (iswdigit(*p)) {
                        nums[numCount++] = (int)wcstol(p, NULL, 10);
                        while (iswdigit(*p)) p++;
                    } else {
                        p++;
                    }
                }
                if (numCount >= 2) {
                    int day, year;
                    if (nums[0] > 31) { year = nums[0]; day = nums[1]; }
                    else if (nums[1] > 31) { day = nums[0]; year = nums[1]; }
                    else if (numCount >= 2) { day = nums[0]; year = nums[1]; }
                    else break;

                    if (year < 100) year += 2000;
                    if (ValidateDate(year, monthNum, day)) {
                        result.year = year; result.month = monthNum; result.day = day;
                        result.valid = true;
                        return result;
                    }
                }
                break;
            }
        }
    }

    return result;
}

/*============================================================================
 *  ICS GENERATOR
 *============================================================================*/

static void IcsEscapeText(const wchar_t* input, wchar_t* output, int outLen) {
    int j = 0;
    for (int i = 0; input[i] && j < outLen - 2; i++) {
        switch (input[i]) {
        case L'\\': if (j < outLen - 2) { output[j++] = L'\\'; output[j++] = L'\\'; } break;
        case L';':  if (j < outLen - 2) { output[j++] = L'\\'; output[j++] = L';';  } break;
        case L',':  if (j < outLen - 2) { output[j++] = L'\\'; output[j++] = L',';  } break;
        case L'\n': if (j < outLen - 2) { output[j++] = L'\\'; output[j++] = L'n';  } break;
        case L'\r': break; /* skip CR */
        default:    output[j++] = input[i]; break;
        }
    }
    output[j] = L'\0';
}

static void IcsWriteFolded(FILE* f, const char* line) {
    int len = (int)strlen(line);
    if (len <= 75) {
        fprintf(f, "%s\r\n", line);
        return;
    }

    /* Write first 75 octets */
    fwrite(line, 1, 75, f);
    fprintf(f, "\r\n");
    int pos = 75;

    while (pos < len) {
        int chunk = len - pos;
        if (chunk > 74) chunk = 74; /* 75 - 1 for the leading space */
        fprintf(f, " ");
        fwrite(line + pos, 1, chunk, f);
        fprintf(f, "\r\n");
        pos += chunk;
    }
}

static char* WideToUtf8(const wchar_t* wide) {
    if (!wide) return NULL;
    int len = WideCharToMultiByte(CP_UTF8, 0, wide, -1, NULL, 0, NULL, NULL);
    char* utf8 = (char*)malloc(len);
    if (utf8) WideCharToMultiByte(CP_UTF8, 0, wide, -1, utf8, len, NULL, NULL);
    return utf8;
}

static void IcsWriteProperty(FILE* f, const char* name, const wchar_t* value) {
    if (!value || !*value) return;

    wchar_t escaped[MAX_FIELD_LEN * 2];
    IcsEscapeText(value, escaped, MAX_FIELD_LEN * 2);

    char* utf8val = WideToUtf8(escaped);
    if (!utf8val) return;

    char line[MAX_FIELD_LEN * 4];
    snprintf(line, sizeof(line), "%s:%s", name, utf8val);
    IcsWriteFolded(f, line);
    free(utf8val);
}

static void IcsFormatDateTime(const ParsedDateTime* dt, char* buf, int bufLen) {
    if (dt->has_time) {
        snprintf(buf, bufLen, "%04d%02d%02dT%02d%02d%02d",
                 dt->year, dt->month, dt->day,
                 dt->hour, dt->minute, dt->second);
    } else {
        snprintf(buf, bufLen, "%04d%02d%02d", dt->year, dt->month, dt->day);
    }
}

static void NextDay(ParsedDateTime* dt) {
    int daysInMonth[] = {0,31,28,31,30,31,30,31,31,30,31,30,31};
    if ((dt->year % 4 == 0 && dt->year % 100 != 0) || dt->year % 400 == 0)
        daysInMonth[2] = 29;
    dt->day++;
    if (dt->day > daysInMonth[dt->month]) {
        dt->day = 1;
        dt->month++;
        if (dt->month > 12) {
            dt->month = 1;
            dt->year++;
        }
    }
}

/* Parse a reminder string like "15 minutes", "1 hour", "30 min" etc. Returns minutes, or -1 if empty/invalid */
static int ParseReminderMinutes(const wchar_t* input) {
    if (!input || !*input) return -1;

    /* Try to extract a number */
    int val = 0;
    const wchar_t* p = input;
    while (*p == L' ') p++;
    if (!iswdigit(*p)) {
        /* Try keyword matching */
        wchar_t lower[128];
        int len = (int)wcslen(input);
        for (int i = 0; i < len && i < 127; i++) lower[i] = towlower(input[i]);
        lower[len < 127 ? len : 127] = L'\0';
        if (wcsstr(lower, L"none") || wcsstr(lower, L"no") || wcsstr(lower, L"off")) return -1;
        return -1;
    }
    while (iswdigit(*p)) { val = val * 10 + (*p - L'0'); p++; }
    while (*p == L' ') p++;

    /* Check units */
    wchar_t lower[64];
    int ulen = (int)wcslen(p);
    for (int i = 0; i < ulen && i < 63; i++) lower[i] = towlower(p[i]);
    lower[ulen < 63 ? ulen : 63] = L'\0';

    if (wcsstr(lower, L"hour") || wcsstr(lower, L"hr")) return val * 60;
    if (wcsstr(lower, L"day")) return val * 1440;
    if (wcsstr(lower, L"week")) return val * 10080;
    /* Default: assume minutes */
    return val;
}

/* Parse a recurrence string. Returns RRULE value or NULL if not recognized */
static const char* ParseRecurrence(const wchar_t* input) {
    if (!input || !*input) return NULL;

    wchar_t lower[128];
    int len = (int)wcslen(input);
    for (int i = 0; i < len && i < 127; i++) lower[i] = towlower(input[i]);
    lower[len < 127 ? len : 127] = L'\0';

    if (wcsstr(lower, L"daily") || wcsstr(lower, L"every day")) return "FREQ=DAILY";
    if (wcsstr(lower, L"weekly") || wcsstr(lower, L"every week")) return "FREQ=WEEKLY";
    if (wcsstr(lower, L"monthly") || wcsstr(lower, L"every month")) return "FREQ=MONTHLY";
    if (wcsstr(lower, L"yearly") || wcsstr(lower, L"annual") || wcsstr(lower, L"every year")) return "FREQ=YEARLY";
    if (wcsstr(lower, L"weekday") || wcsstr(lower, L"work")) return "FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR";
    if (wcsstr(lower, L"fortnigh") || wcsstr(lower, L"bi-week") || wcsstr(lower, L"biweek")) return "FREQ=WEEKLY;INTERVAL=2";

    return NULL;
}

static bool IsBoolTrue(const wchar_t* input) {
    if (!input || !*input) return false;
    wchar_t lower[32];
    int len = (int)wcslen(input);
    for (int i = 0; i < len && i < 31; i++) lower[i] = towlower(input[i]);
    lower[len < 31 ? len : 31] = L'\0';
    return wcscmp(lower, L"yes") == 0 || wcscmp(lower, L"true") == 0 ||
           wcscmp(lower, L"1") == 0 || wcscmp(lower, L"y") == 0;
}

static void IcsWriteValarm(FILE* f, int minutes) {
    if (minutes < 0) return;
    fprintf(f, "BEGIN:VALARM\r\n");
    fprintf(f, "ACTION:DISPLAY\r\n");
    fprintf(f, "DESCRIPTION:Reminder\r\n");
    if (minutes == 0) {
        fprintf(f, "TRIGGER:PT0S\r\n");
    } else if (minutes >= 1440 && minutes % 1440 == 0) {
        fprintf(f, "TRIGGER:-P%dD\r\n", minutes / 1440);
    } else if (minutes >= 60 && minutes % 60 == 0) {
        fprintf(f, "TRIGGER:-PT%dH\r\n", minutes / 60);
    } else {
        fprintf(f, "TRIGGER:-PT%dM\r\n", minutes);
    }
    fprintf(f, "END:VALARM\r\n");
}

static void IcsGenerateUID(char* buf, int bufLen) {
    /* Generate RFC 4122 version 4 UUID */
    unsigned char bytes[16];
    HCRYPTPROV hProv;
    if (CryptAcquireContext(&hProv, NULL, NULL, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)) {
        CryptGenRandom(hProv, 16, bytes);
        CryptReleaseContext(hProv, 0);
    } else {
        /* Fallback: use time + counter */
        LARGE_INTEGER pc;
        QueryPerformanceCounter(&pc);
        srand((unsigned)(pc.QuadPart ^ GetTickCount()));
        for (int i = 0; i < 16; i++) bytes[i] = (unsigned char)rand();
    }
    /* Set version 4 and variant bits */
    bytes[6] = (bytes[6] & 0x0F) | 0x40;
    bytes[8] = (bytes[8] & 0x3F) | 0x80;
    snprintf(buf, bufLen,
        "%02x%02x%02x%02x-%02x%02x-%02x%02x-%02x%02x-%02x%02x%02x%02x%02x%02x",
        bytes[0], bytes[1], bytes[2], bytes[3],
        bytes[4], bytes[5], bytes[6], bytes[7],
        bytes[8], bytes[9], bytes[10], bytes[11],
        bytes[12], bytes[13], bytes[14], bytes[15]);
}

static const wchar_t* GetField(CsvRow* row, int* mapping, int field) {
    if (mapping[field] >= 0 && mapping[field] < row->field_count)
        return row->fields[mapping[field]];
    return NULL;
}

static bool IcsWriteEvent(FILE* f, CsvRow* row, AppState* state, int eventIndex) {
    (void)eventIndex;
    int* mapping = state->fieldMapping;
    int datePref = state->dateFormatPref;

    /* DTSTART is required */
    const wchar_t* dtStartStr = GetField(row, mapping, ICS_DTSTART);
    if (!dtStartStr) return false;
    ParsedDateTime dtStart = ParseDateTime(dtStartStr, datePref);
    if (!dtStart.valid) return false;

    /* Check All Day field */
    const wchar_t* allDayStr = GetField(row, mapping, ICS_ALL_DAY);
    bool forceAllDay = allDayStr && IsBoolTrue(allDayStr);

    /* Combine separate time column with date if present */
    if (!forceAllDay) {
        const wchar_t* startTimeStr = GetField(row, mapping, ICS_START_TIME);
        if (startTimeStr && *startTimeStr) {
            int h, m, s;
            if (TryParseTime(startTimeStr, &h, &m, &s)) {
                dtStart.hour = h; dtStart.minute = m; dtStart.second = s;
                dtStart.has_time = true;
            }
        }
    }

    if (forceAllDay) {
        dtStart.has_time = false;
    }

    fprintf(f, "BEGIN:VEVENT\r\n");

    /* UID */
    char uid[128];
    IcsGenerateUID(uid, sizeof(uid));
    char uidLine[256];
    snprintf(uidLine, sizeof(uidLine), "UID:%s", uid);
    IcsWriteFolded(f, uidLine);

    /* DTSTAMP */
    SYSTEMTIME st;
    GetSystemTime(&st);
    fprintf(f, "DTSTAMP:%04d%02d%02dT%02d%02d%02dZ\r\n",
            st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);

    /* DTSTART */
    char dtBuf[64];
    IcsFormatDateTime(&dtStart, dtBuf, sizeof(dtBuf));
    if (dtStart.has_time) {
        fprintf(f, "DTSTART:%s\r\n", dtBuf);
    } else {
        fprintf(f, "DTSTART;VALUE=DATE:%s\r\n", dtBuf);
    }

    /* DTEND */
    {
        ParsedDateTime dtEnd = {0};
        bool hasEnd = false;

        const wchar_t* dtEndStr = GetField(row, mapping, ICS_DTEND);
        if (dtEndStr && *dtEndStr) {
            dtEnd = ParseDateTime(dtEndStr, datePref);
            hasEnd = dtEnd.valid;
        }

        /* Combine separate end time column */
        if (hasEnd && !forceAllDay) {
            const wchar_t* endTimeStr = GetField(row, mapping, ICS_END_TIME);
            if (endTimeStr && *endTimeStr) {
                int h, m, s;
                if (TryParseTime(endTimeStr, &h, &m, &s)) {
                    dtEnd.hour = h; dtEnd.minute = m; dtEnd.second = s;
                    dtEnd.has_time = true;
                }
            }
        }

        if (!hasEnd) {
            dtEnd = dtStart;
            hasEnd = true;
        }

        if (forceAllDay) {
            dtEnd.has_time = false;
        }

        /* For all-day events (VALUE=DATE), DTEND is exclusive per RFC 5545.
           If DTEND <= DTSTART, set DTEND to DTSTART + 1 day. */
        if (!dtEnd.has_time) {
            if (dtEnd.year < dtStart.year ||
                (dtEnd.year == dtStart.year && dtEnd.month < dtStart.month) ||
                (dtEnd.year == dtStart.year && dtEnd.month == dtStart.month && dtEnd.day <= dtStart.day)) {
                dtEnd = dtStart;
                dtEnd.has_time = false;
                NextDay(&dtEnd);
            }
        }

        IcsFormatDateTime(&dtEnd, dtBuf, sizeof(dtBuf));
        if (dtEnd.has_time) {
            fprintf(f, "DTEND:%s\r\n", dtBuf);
        } else {
            fprintf(f, "DTEND;VALUE=DATE:%s\r\n", dtBuf);
        }
    }

    /* Text properties */
    struct { int icsField; const char* propName; } textProps[] = {
        { ICS_SUMMARY,     "SUMMARY" },
        { ICS_DESCRIPTION, "DESCRIPTION" },
        { ICS_LOCATION,    "LOCATION" },
        { ICS_URL,         "URL" },
        { ICS_CATEGORIES,  "CATEGORIES" },
        { ICS_STATUS,      "STATUS" },
        { ICS_TRANSP,      "TRANSP" },
        { ICS_PRIORITY,    "PRIORITY" },
    };

    for (int i = 0; i < (int)(sizeof(textProps) / sizeof(textProps[0])); i++) {
        const wchar_t* val = GetField(row, mapping, textProps[i].icsField);
        if (val && *val) {
            IcsWriteProperty(f, textProps[i].propName, val);
        }
    }

    /* Recurrence (RRULE) */
    const wchar_t* recurStr = GetField(row, mapping, ICS_RECURRENCE);
    if (recurStr) {
        const char* rrule = ParseRecurrence(recurStr);
        if (rrule) {
            fprintf(f, "RRULE:%s\r\n", rrule);
        }
    }

    /* Apple Travel Advisory - if location is present and user enabled it */
    if (state->appleTravel) {
        const wchar_t* loc = GetField(row, mapping, ICS_LOCATION);
        if (loc && *loc) {
            fprintf(f, "X-APPLE-TRAVEL-ADVISORY-BEHAVIOR:AUTOMATIC\r\n");
        }
    }

    /* VALARM - from CSV reminder column first, then default reminders */
    bool wroteAlarm1 = false;
    const wchar_t* reminderStr = GetField(row, mapping, ICS_REMINDER);
    if (reminderStr) {
        int mins = ParseReminderMinutes(reminderStr);
        if (mins >= 0) {
            IcsWriteValarm(f, mins);
            wroteAlarm1 = true;
        }
    }

    /* Default reminders (from page 3 settings) */
    if (!wroteAlarm1 && state->defaultReminder1 > 0) {
        IcsWriteValarm(f, REMINDER_VALUES[state->defaultReminder1]);
    }
    if (state->defaultReminder2 > 0) {
        IcsWriteValarm(f, REMINDER_VALUES[state->defaultReminder2]);
    }

    fprintf(f, "END:VEVENT\r\n");
    return true;
}

static void IcsWriteHeader(FILE* f) {
    fprintf(f, "BEGIN:VCALENDAR\r\n");
    fprintf(f, "VERSION:2.0\r\n");
    fprintf(f, "PRODID:-//CSV2ICS//CSV2ICS v1.0//EN\r\n");
    fprintf(f, "CALSCALE:GREGORIAN\r\n");
    fprintf(f, "METHOD:PUBLISH\r\n");
    /* Calendar ID for Apple Calendar grouping */
    char calId[128];
    IcsGenerateUID(calId, sizeof(calId));
    fprintf(f, "X-WR-RELCALID:%s\r\n", calId);
    fprintf(f, "X-WR-TIMEZONE:Europe/London\r\n");
    /* VTIMEZONE for UK: GMT Standard Time / BST */
    fprintf(f, "BEGIN:VTIMEZONE\r\n");
    fprintf(f, "TZID:Europe/London\r\n");
    fprintf(f, "BEGIN:STANDARD\r\n");
    fprintf(f, "DTSTART:19701025T020000\r\n");
    fprintf(f, "RRULE:FREQ=YEARLY;BYDAY=-1SU;BYMONTH=10\r\n");
    fprintf(f, "TZNAME:GMT\r\n");
    fprintf(f, "TZOFFSETFROM:+0100\r\n");
    fprintf(f, "TZOFFSETTO:+0000\r\n");
    fprintf(f, "END:STANDARD\r\n");
    fprintf(f, "BEGIN:DAYLIGHT\r\n");
    fprintf(f, "DTSTART:19700329T010000\r\n");
    fprintf(f, "RRULE:FREQ=YEARLY;BYDAY=-1SU;BYMONTH=3\r\n");
    fprintf(f, "TZNAME:BST\r\n");
    fprintf(f, "TZOFFSETFROM:+0000\r\n");
    fprintf(f, "TZOFFSETTO:+0100\r\n");
    fprintf(f, "END:DAYLIGHT\r\n");
    fprintf(f, "END:VTIMEZONE\r\n");
}

static void IcsWriteFooter(FILE* f) {
    fprintf(f, "END:VCALENDAR\r\n");
}

static bool IcsExportSingle(const wchar_t* path, AppState* state) {
    FILE* f = _wfopen(path, L"wb");
    if (!f) return false;

    IcsWriteHeader(f);

    int exported = 0;
    for (int i = 0; i < state->csv.row_count; i++) {
        if (IcsWriteEvent(f, &state->csv.rows[i], state, i)) {
            exported++;
        }
    }

    IcsWriteFooter(f);
    fclose(f);
    return exported > 0;
}

static bool IcsExportSeparate(const wchar_t* folder, AppState* state) {
    int exported = 0;

    for (int i = 0; i < state->csv.row_count; i++) {
        wchar_t filename[MAX_PATH];

        /* Use summary field for filename if available */
        wchar_t safeName[128] = L"event";
        if (state->fieldMapping[ICS_SUMMARY] >= 0 &&
            state->fieldMapping[ICS_SUMMARY] < state->csv.rows[i].field_count) {
            const wchar_t* summary = state->csv.rows[i].fields[state->fieldMapping[ICS_SUMMARY]];
            if (summary && *summary) {
                int j = 0;
                for (int k = 0; summary[k] && j < 100; k++) {
                    wchar_t ch = summary[k];
                    if (iswalnum(ch) || ch == L' ' || ch == L'-' || ch == L'_') {
                        safeName[j++] = ch;
                    }
                }
                safeName[j] = L'\0';
                if (j == 0) wcscpy(safeName, L"event");
            }
        }

        swprintf(filename, MAX_PATH, L"%s\\%s_%04d.ics", folder, safeName, i + 1);

        FILE* f = _wfopen(filename, L"wb");
        if (!f) continue;

        IcsWriteHeader(f);
        if (IcsWriteEvent(f, &state->csv.rows[i], state, i)) {
            exported++;
        }
        IcsWriteFooter(f);
        fclose(f);
    }

    return exported > 0;
}

/*============================================================================
 *  UI HELPERS
 *============================================================================*/

static HWND CreateLabel(HWND parent, const wchar_t* text, int x, int y, int w, int h, int id, HFONT font) {
    HWND hwnd = CreateWindowEx(0, L"STATIC", text,
        WS_CHILD | SS_LEFT,
        x, y, w, h, parent, (HMENU)(INT_PTR)id, GetModuleHandle(NULL), NULL);
    SendMessage(hwnd, WM_SETFONT, (WPARAM)font, TRUE);
    return hwnd;
}

static HWND CreateBtn(HWND parent, const wchar_t* text, int x, int y, int w, int h, int id, HFONT font) {
    HWND hwnd = CreateWindowEx(0, L"BUTTON", text,
        WS_CHILD | WS_TABSTOP | BS_PUSHBUTTON,
        x, y, w, h, parent, (HMENU)(INT_PTR)id, GetModuleHandle(NULL), NULL);
    SendMessage(hwnd, WM_SETFONT, (WPARAM)font, TRUE);
    return hwnd;
}

static HWND CreateCheck(HWND parent, const wchar_t* text, int x, int y, int w, int h, int id, HFONT font) {
    HWND hwnd = CreateWindowEx(0, L"BUTTON", text,
        WS_CHILD | WS_TABSTOP | BS_AUTOCHECKBOX,
        x, y, w, h, parent, (HMENU)(INT_PTR)id, GetModuleHandle(NULL), NULL);
    SendMessage(hwnd, WM_SETFONT, (WPARAM)font, TRUE);
    return hwnd;
}

static HWND CreateRadio(HWND parent, const wchar_t* text, int x, int y, int w, int h, int id, HFONT font, bool group) {
    DWORD style = WS_CHILD | WS_TABSTOP | BS_AUTORADIOBUTTON;
    if (group) style |= WS_GROUP;
    HWND hwnd = CreateWindowEx(0, L"BUTTON", text, style,
        x, y, w, h, parent, (HMENU)(INT_PTR)id, GetModuleHandle(NULL), NULL);
    SendMessage(hwnd, WM_SETFONT, (WPARAM)font, TRUE);
    return hwnd;
}

static HWND CreateCombo(HWND parent, int x, int y, int w, int h, int id, HFONT font) {
    HWND hwnd = CreateWindowEx(0, L"COMBOBOX", L"",
        WS_CHILD | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL,
        x, y, w, h, parent, (HMENU)(INT_PTR)id, GetModuleHandle(NULL), NULL);
    SendMessage(hwnd, WM_SETFONT, (WPARAM)font, TRUE);
    return hwnd;
}

static void CreatePage1Controls(AppState* state, HWND parent) {
    HFONT font = state->hFont;
    int y = 20;

    HFONT titleFont = CreateFont(-18, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, 0, 0, CLEARTYPE_QUALITY, 0, L"Segoe UI");

    state->hwndLblPage1Title = CreateLabel(parent, L"Step 1: Select CSV File", 20, y, 400, 25, 0, titleFont);
    y += 40;

    state->hwndBtnOpen = CreateBtn(parent, L"Open CSV File...", 20, y, 150, 30, IDC_BTN_OPEN, font);
    state->hwndLblPath = CreateLabel(parent, L"No file selected", 180, y + 5, 580, 20, IDC_LBL_PATH, font);
    y += 45;

    state->hwndChkHeader = CreateCheck(parent, L"First row contains headers", 20, y, 250, 25, IDC_CHK_HEADER, font);
    y += 35;

    state->hwndLblDateFmt = CreateLabel(parent, L"Date format:", 20, y + 2, 100, 20, 0, font);
    state->hwndRadioMDY = CreateRadio(parent, L"MM/DD/YYYY", 130, y, 120, 25, IDC_RADIO_MDY, font, true);
    state->hwndRadioDMY = CreateRadio(parent, L"DD/MM/YYYY", 260, y, 120, 25, IDC_RADIO_DMY, font, false);
    state->hwndRadioISO = CreateRadio(parent, L"YYYY-MM-DD", 390, y, 120, 25, IDC_RADIO_ISO, font, false);
    SendMessage(state->hwndRadioDMY, BM_SETCHECK, BST_CHECKED, 0);
    y += 35;

    CreateLabel(parent, L"Preview:", 20, y, 100, 20, 0, font);
    y += 22;

    int lvHeight = NAV_Y - y - 15;
    state->hwndListView = CreateWindowEx(WS_EX_CLIENTEDGE, WC_LISTVIEW, L"",
        WS_CHILD | LVS_REPORT | LVS_SINGLESEL | LVS_NOSORTHEADER,
        20, y, 780, lvHeight, parent, (HMENU)(INT_PTR)IDC_LISTVIEW, GetModuleHandle(NULL), NULL);
    SendMessage(state->hwndListView, WM_SETFONT, (WPARAM)font, TRUE);
    ListView_SetExtendedListViewStyle(state->hwndListView, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
}

static void CreatePage2Controls(AppState* state, HWND parent) {
    HFONT font = state->hFont;
    int y = 20;

    HFONT titleFont = CreateFont(-18, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, 0, 0, CLEARTYPE_QUALITY, 0, L"Segoe UI");

    state->hwndLblMapTitle = CreateLabel(parent, L"Step 2: Map CSV Columns to ICS Fields", 20, y, 500, 25, 0, titleFont);
    y += 35;

    /* Use two-column layout: left side for core fields, right side for extra fields */
    int labelX = 20, comboX = 170, rowH = 28;
    int labelX2 = 440, comboX2 = 570;

    /* Left column: core fields (0-9) */
    for (int i = 0; i <= ICS_PRIORITY; i++) {
        state->hwndComboLabel[i] = CreateLabel(parent, ICS_FIELD_NAMES[i], labelX, y + 3, 145, 20, 0, font);
        state->hwndCombo[i] = CreateCombo(parent, comboX, y, 240, 200, COMBO_IDS[i], font);
        y += rowH;
    }

    /* Right column: new fields (Start Time, End Time, All Day, Reminder, Recurrence) */
    int ry = 55; /* align with first combo row */
    for (int i = ICS_START_TIME; i < ICS_FIELD_COUNT; i++) {
        state->hwndComboLabel[i] = CreateLabel(parent, ICS_FIELD_NAMES[i], labelX2, ry + 3, 125, 20, 0, font);
        state->hwndCombo[i] = CreateCombo(parent, comboX2, ry, 230, 200, COMBO_IDS[i], font);
        ry += rowH;
    }

    y += 8;
    state->hwndLblPreview = CreateLabel(parent, L"ICS Preview (first event):", 20, y, 250, 20, 0, font);
    y += 22;

    int prevHeight = NAV_Y - y - 15;
    state->hwndPreview = CreateWindowEx(WS_EX_CLIENTEDGE, L"EDIT", L"",
        WS_CHILD | WS_VSCROLL | ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL,
        20, y, 780, prevHeight, parent, (HMENU)(INT_PTR)IDC_PREVIEW, GetModuleHandle(NULL), NULL);

    HFONT monoFont = CreateFont(-13, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, 0, 0, CLEARTYPE_QUALITY, FIXED_PITCH, L"Consolas");
    SendMessage(state->hwndPreview, WM_SETFONT, (WPARAM)monoFont, TRUE);
}

static void CreatePage3Controls(AppState* state, HWND parent) {
    HFONT font = state->hFont;
    int y = 20;

    HFONT titleFont = CreateFont(-18, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, 0, 0, CLEARTYPE_QUALITY, 0, L"Segoe UI");

    state->hwndLblExportTitle = CreateLabel(parent, L"Step 3: Export ICS", 20, y, 400, 25, 0, titleFont);
    y += 45;

    state->hwndRadioSingle = CreateRadio(parent, L"Single ICS file (all events)", 20, y, 300, 25, IDC_RADIO_SINGLE, font, true);
    y += 28;
    state->hwndRadioSeparate = CreateRadio(parent, L"Separate ICS file per event", 20, y, 300, 25, IDC_RADIO_SEPARATE, font, false);
    SendMessage(state->hwndRadioSingle, BM_SETCHECK, BST_CHECKED, 0);
    y += 40;

    /* Default reminders */
    HFONT sectionFont = CreateFont(-15, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, 0, 0, CLEARTYPE_QUALITY, 0, L"Segoe UI");
    state->hwndLblRemindSection = CreateLabel(parent, L"Default Reminders:", 20, y, 200, 20, 0, sectionFont);
    y += 28;

    state->hwndLblRemind1 = CreateLabel(parent, L"Reminder 1:", 20, y + 3, 100, 20, 0, font);
    state->hwndComboRemind1 = CreateCombo(parent, 130, y, 220, 200, IDC_COMBO_REMIND1, font);
    for (int i = 0; i < REMINDER_OPTION_COUNT; i++)
        SendMessage(state->hwndComboRemind1, CB_ADDSTRING, 0, (LPARAM)REMINDER_LABELS[i]);
    SendMessage(state->hwndComboRemind1, CB_SETCURSEL, 4, 0); /* Default: 15 minutes */
    state->defaultReminder1 = 4;
    y += 30;

    state->hwndLblRemind2 = CreateLabel(parent, L"Reminder 2:", 20, y + 3, 100, 20, 0, font);
    state->hwndComboRemind2 = CreateCombo(parent, 130, y, 220, 200, IDC_COMBO_REMIND2, font);
    for (int i = 0; i < REMINDER_OPTION_COUNT; i++)
        SendMessage(state->hwndComboRemind2, CB_ADDSTRING, 0, (LPARAM)REMINDER_LABELS[i]);
    SendMessage(state->hwndComboRemind2, CB_SETCURSEL, 0, 0); /* Default: None */
    state->defaultReminder2 = 0;
    y += 35;

    /* Apple Travel */
    state->hwndChkTravel = CreateCheck(parent, L"Apple Calendar: \"Time to Leave\" alerts (requires Location)",
                                       20, y, 500, 25, IDC_CHK_TRAVEL, font);
    y += 40;

    state->hwndLblCount = CreateLabel(parent, L"Events to export: 0", 20, y, 300, 25, IDC_LBL_COUNT, font);
    y += 35;

    state->hwndBtnExport = CreateBtn(parent, L"Export...", 20, y, 150, 35, IDC_BTN_EXPORT, font);
    state->hwndBtnStartOver = CreateBtn(parent, L"Start Over", 190, y, 130, 35, IDC_BTN_STARTOVER, font);
    y += 50;

    state->hwndLblStatus = CreateLabel(parent, L"Ready", 20, y, 780, 60, IDC_LBL_STATUS, font);
}

static void ShowControls(HWND* controls, int count, int show) {
    for (int i = 0; i < count; i++) {
        if (controls[i]) ShowWindow(controls[i], show ? SW_SHOW : SW_HIDE);
    }
}

static void ShowPage(AppState* state, int page) {
    /* Page 1 controls */
    HWND page1[] = {
        state->hwndLblPage1Title, state->hwndBtnOpen, state->hwndLblPath,
        state->hwndChkHeader, state->hwndLblDateFmt,
        state->hwndRadioMDY, state->hwndRadioDMY, state->hwndRadioISO, state->hwndListView
    };

    /* Page 2 controls */
    HWND page2[ICS_FIELD_COUNT * 2 + 4];
    int p2count = 0;
    page2[p2count++] = state->hwndLblMapTitle;
    for (int i = 0; i < ICS_FIELD_COUNT; i++) {
        page2[p2count++] = state->hwndComboLabel[i];
        page2[p2count++] = state->hwndCombo[i];
    }
    page2[p2count++] = state->hwndLblPreview;
    page2[p2count++] = state->hwndPreview;

    /* Page 3 controls */
    HWND page3[] = {
        state->hwndLblExportTitle, state->hwndRadioSingle, state->hwndRadioSeparate,
        state->hwndLblRemindSection,
        state->hwndLblRemind1, state->hwndComboRemind1,
        state->hwndLblRemind2, state->hwndComboRemind2,
        state->hwndChkTravel,
        state->hwndLblCount, state->hwndBtnExport, state->hwndBtnStartOver, state->hwndLblStatus
    };

    ShowControls(page1, sizeof(page1) / sizeof(page1[0]), page == PAGE_FILE);
    ShowControls(page2, p2count, page == PAGE_MAP);
    ShowControls(page3, sizeof(page3) / sizeof(page3[0]), page == PAGE_EXPORT);

    /* Navigation buttons */
    ShowWindow(state->hwndBtnBack, page > PAGE_FILE ? SW_SHOW : SW_HIDE);
    if (page == PAGE_EXPORT) {
        SetWindowText(state->hwndBtnNext, L"Close");
        ShowWindow(state->hwndBtnNext, SW_SHOW);
    } else {
        SetWindowText(state->hwndBtnNext, L"Next >");
        ShowWindow(state->hwndBtnNext, SW_SHOW);
    }

    state->currentPage = page;

    if (page == PAGE_MAP) {
        PopulateMappingCombos(state);
        UpdateIcsPreview(state);
    }
    if (page == PAGE_EXPORT) {
        wchar_t countText[64];
        swprintf(countText, 64, L"Events to export: %d", state->csv.row_count);
        SetWindowText(state->hwndLblCount, countText);
        SetWindowText(state->hwndLblStatus, L"Ready");
    }
}

static void PopulateListView(AppState* state) {
    ListView_DeleteAllItems(state->hwndListView);

    /* Remove all columns */
    while (ListView_DeleteColumn(state->hwndListView, 0)) {}

    if (state->csv.col_count == 0) return;

    /* Add columns */
    int colWidth = 760 / state->csv.col_count;
    if (colWidth < 80) colWidth = 80;
    if (colWidth > 200) colWidth = 200;

    for (int i = 0; i < state->csv.col_count; i++) {
        LVCOLUMN lvc = {0};
        lvc.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT;
        lvc.fmt = LVCFMT_LEFT;
        lvc.cx = colWidth;
        if (state->csv.has_header) {
            lvc.pszText = state->csv.header.fields[i];
        } else {
            wchar_t colName[32];
            swprintf(colName, 32, L"Column %d", i + 1);
            lvc.pszText = colName;
        }
        ListView_InsertColumn(state->hwndListView, i, &lvc);
    }

    /* Add rows (max preview rows) */
    int maxPreview = state->csv.row_count < 100 ? state->csv.row_count : 100;

    /* If no header detected, first row of data is in header struct and should show */
    if (!state->csv.has_header) {
        /* Insert the "header" row as first data row */
        LVITEM lvi = {0};
        lvi.mask = LVIF_TEXT;
        lvi.iItem = 0;
        lvi.pszText = state->csv.header.fields[0];
        ListView_InsertItem(state->hwndListView, &lvi);
        for (int c = 1; c < state->csv.header.field_count; c++) {
            ListView_SetItemText(state->hwndListView, 0, c, state->csv.header.fields[c]);
        }
    }

    for (int r = 0; r < maxPreview; r++) {
        int displayRow = state->csv.has_header ? r : r + 1;
        LVITEM lvi = {0};
        lvi.mask = LVIF_TEXT;
        lvi.iItem = displayRow;
        lvi.pszText = (r < state->csv.rows[r].field_count) ? state->csv.rows[r].fields[0] : L"";
        ListView_InsertItem(state->hwndListView, &lvi);

        for (int c = 1; c < state->csv.rows[r].field_count; c++) {
            ListView_SetItemText(state->hwndListView, displayRow, c, state->csv.rows[r].fields[c]);
        }
    }
}

static void PopulateMappingCombos(AppState* state) {
    for (int i = 0; i < ICS_FIELD_COUNT; i++) {
        SendMessage(state->hwndCombo[i], CB_RESETCONTENT, 0, 0);

        /* Add "(none)" for optional fields */
        if (i != ICS_DTSTART) {
            SendMessage(state->hwndCombo[i], CB_ADDSTRING, 0, (LPARAM)L"(none)");
        }

        /* Add column names */
        for (int c = 0; c < state->csv.col_count; c++) {
            wchar_t label[256];
            if (state->csv.has_header && state->csv.header.fields[c]) {
                swprintf(label, 256, L"%s", state->csv.header.fields[c]);
            } else {
                swprintf(label, 256, L"Column %d", c + 1);
            }
            SendMessage(state->hwndCombo[i], CB_ADDSTRING, 0, (LPARAM)label);
        }

        /* Set selection based on mapping */
        int mapIdx = state->fieldMapping[i];
        if (mapIdx < 0) {
            SendMessage(state->hwndCombo[i], CB_SETCURSEL, 0, 0);
        } else {
            int comboIdx = (i == ICS_DTSTART) ? mapIdx : mapIdx + 1;
            SendMessage(state->hwndCombo[i], CB_SETCURSEL, comboIdx, 0);
        }
    }
}

static void AutoMapFields(AppState* state) {
    /* Initialize all mappings to -1 (unmapped) */
    for (int i = 0; i < ICS_FIELD_COUNT; i++) {
        state->fieldMapping[i] = -1;
    }

    if (!state->csv.has_header) {
        /* Without headers, try to detect date columns */
        if (state->csv.row_count > 0) {
            for (int c = 0; c < state->csv.rows[0].field_count; c++) {
                ParsedDateTime dt = ParseDateTime(state->csv.rows[0].fields[c], state->dateFormatPref);
                if (dt.valid && state->fieldMapping[ICS_DTSTART] < 0) {
                    state->fieldMapping[ICS_DTSTART] = c;
                    break;
                }
            }
        }
        return;
    }

    /* Try to auto-map based on header names */
    struct { int field; const wchar_t* keywords[8]; } autoMap[] = {
        { ICS_DTSTART,     { L"start", L"dtstart", L"begin", L"date", L"start date", L"start_date", NULL } },
        { ICS_DTEND,       { L"end", L"dtend", L"end date", L"end_date", L"finish", NULL } },
        { ICS_SUMMARY,     { L"summary", L"title", L"subject", L"name", L"event", NULL } },
        { ICS_DESCRIPTION, { L"description", L"desc", L"details", L"notes", L"body", NULL } },
        { ICS_LOCATION,    { L"location", L"place", L"venue", L"where", L"address", NULL } },
        { ICS_URL,         { L"url", L"link", L"website", NULL } },
        { ICS_CATEGORIES,  { L"categories", L"category", L"tags", L"type", NULL } },
        { ICS_STATUS,      { L"status", NULL } },
        { ICS_TRANSP,      { L"transp", L"transparency", L"show as", NULL } },
        { ICS_PRIORITY,    { L"priority", L"importance", NULL } },
        { ICS_START_TIME,  { L"start time", L"start_time", L"time", L"begins", NULL } },
        { ICS_END_TIME,    { L"end time", L"end_time", L"finish time", NULL } },
        { ICS_ALL_DAY,     { L"all day", L"all_day", L"allday", L"whole day", NULL } },
        { ICS_REMINDER,    { L"reminder", L"alarm", L"alert", NULL } },
        { ICS_RECURRENCE,  { L"recurrence", L"recurring", L"repeat", L"frequency", L"rrule", NULL } },
    };

    for (int m = 0; m < (int)(sizeof(autoMap) / sizeof(autoMap[0])); m++) {
        for (int c = 0; c < state->csv.col_count; c++) {
            wchar_t lower[256];
            int hlen = (int)wcslen(state->csv.header.fields[c]);
            for (int k = 0; k < hlen && k < 255; k++)
                lower[k] = towlower(state->csv.header.fields[c][k]);
            lower[hlen < 255 ? hlen : 255] = L'\0';

            for (int k = 0; autoMap[m].keywords[k]; k++) {
                if (wcscmp(lower, autoMap[m].keywords[k]) == 0) {
                    /* Check this column isn't already mapped */
                    bool alreadyMapped = false;
                    for (int f = 0; f < ICS_FIELD_COUNT; f++) {
                        if (state->fieldMapping[f] == c) { alreadyMapped = true; break; }
                    }
                    if (!alreadyMapped) {
                        state->fieldMapping[autoMap[m].field] = c;
                        goto nextField;
                    }
                }
            }
        }
        nextField:;
    }
}

static void UpdateIcsPreview(AppState* state) {
    if (state->csv.row_count == 0) {
        SetWindowText(state->hwndPreview, L"No data to preview");
        return;
    }

    wchar_t preview[4096] = L"";
    CsvRow* row = &state->csv.rows[0];
    int* mapping = state->fieldMapping;
    int datePref = state->dateFormatPref;

    wcscat(preview, L"BEGIN:VEVENT\r\n");

    /* DTSTART with optional separate time */
    if (mapping[ICS_DTSTART] >= 0 && mapping[ICS_DTSTART] < row->field_count) {
        ParsedDateTime dt = ParseDateTime(row->fields[mapping[ICS_DTSTART]], datePref);
        if (dt.valid) {
            /* Check All Day */
            const wchar_t* allDayStr = GetField(row, mapping, ICS_ALL_DAY);
            bool forceAllDay = allDayStr && IsBoolTrue(allDayStr);

            if (!forceAllDay) {
                const wchar_t* timeStr = GetField(row, mapping, ICS_START_TIME);
                if (timeStr && *timeStr) {
                    int h, m, s;
                    if (TryParseTime(timeStr, &h, &m, &s)) {
                        dt.hour = h; dt.minute = m; dt.second = s;
                        dt.has_time = true;
                    }
                }
            }
            if (forceAllDay) dt.has_time = false;

            wchar_t line[128];
            if (dt.has_time) {
                swprintf(line, 128, L"DTSTART:%04d%02d%02dT%02d%02d%02d\r\n",
                         dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second);
            } else {
                swprintf(line, 128, L"DTSTART;VALUE=DATE:%04d%02d%02d\r\n",
                         dt.year, dt.month, dt.day);
            }
            wcscat(preview, line);
        } else {
            wcscat(preview, L"DTSTART: (invalid date)\r\n");
        }
    }

    /* DTEND with optional separate time */
    if (mapping[ICS_DTEND] >= 0 && mapping[ICS_DTEND] < row->field_count) {
        ParsedDateTime dt = ParseDateTime(row->fields[mapping[ICS_DTEND]], datePref);
        if (dt.valid) {
            const wchar_t* timeStr = GetField(row, mapping, ICS_END_TIME);
            if (timeStr && *timeStr) {
                int h, m, s;
                if (TryParseTime(timeStr, &h, &m, &s)) {
                    dt.hour = h; dt.minute = m; dt.second = s;
                    dt.has_time = true;
                }
            }
            wchar_t line[128];
            if (dt.has_time) {
                swprintf(line, 128, L"DTEND:%04d%02d%02dT%02d%02d%02d\r\n",
                         dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second);
            } else {
                swprintf(line, 128, L"DTEND;VALUE=DATE:%04d%02d%02d\r\n",
                         dt.year, dt.month, dt.day);
            }
            wcscat(preview, line);
        }
    }

    /* Text properties */
    struct { int field; const wchar_t* name; } fields[] = {
        { ICS_SUMMARY, L"SUMMARY" }, { ICS_DESCRIPTION, L"DESCRIPTION" },
        { ICS_LOCATION, L"LOCATION" }, { ICS_URL, L"URL" },
        { ICS_CATEGORIES, L"CATEGORIES" }, { ICS_STATUS, L"STATUS" },
        { ICS_TRANSP, L"TRANSP" }, { ICS_PRIORITY, L"PRIORITY" },
    };

    for (int i = 0; i < (int)(sizeof(fields) / sizeof(fields[0])); i++) {
        const wchar_t* val = GetField(row, mapping, fields[i].field);
        if (val && *val) {
            wchar_t escaped[1024];
            IcsEscapeText(val, escaped, 1024);
            wchar_t line[1200];
            swprintf(line, 1200, L"%s:%s\r\n", fields[i].name, escaped);
            if (wcslen(preview) + wcslen(line) < 3800)
                wcscat(preview, line);
        }
    }

    /* Recurrence */
    const wchar_t* recurStr = GetField(row, mapping, ICS_RECURRENCE);
    if (recurStr) {
        const char* rrule = ParseRecurrence(recurStr);
        if (rrule) {
            wchar_t line[256];
            swprintf(line, 256, L"RRULE:%hs\r\n", rrule);
            wcscat(preview, line);
        }
    }

    /* Reminder from CSV */
    const wchar_t* reminderStr = GetField(row, mapping, ICS_REMINDER);
    if (reminderStr) {
        int mins = ParseReminderMinutes(reminderStr);
        if (mins >= 0) {
            wchar_t line[128];
            swprintf(line, 128, L"BEGIN:VALARM ... TRIGGER:-PT%dM\r\n", mins);
            wcscat(preview, line);
        }
    }

    wcscat(preview, L"END:VEVENT");
    SetWindowText(state->hwndPreview, preview);
}

static void DoOpenFile(AppState* state) {
    wchar_t filePath[MAX_PATH] = L"";

    OPENFILENAME ofn = {0};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = state->hwndMain;
    ofn.lpstrFilter = L"CSV Files (*.csv)\0*.csv\0All Files (*.*)\0*.*\0";
    ofn.lpstrFile = filePath;
    ofn.nMaxFile = MAX_PATH;
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST;
    ofn.lpstrTitle = L"Open CSV File";

    if (!GetOpenFileName(&ofn)) return;

    wcscpy(state->csvFilePath, filePath);

    if (CsvLoad(state, filePath)) {
        /* Update path label */
        wchar_t* fileName = wcsrchr(filePath, L'\\');
        if (!fileName) fileName = filePath; else fileName++;
        wchar_t label[MAX_PATH + 32];
        swprintf(label, MAX_PATH + 32, L"%s (%d rows, %d columns)",
                 fileName, state->csv.row_count, state->csv.col_count);
        SetWindowText(state->hwndLblPath, label);

        /* Set header checkbox */
        SendMessage(state->hwndChkHeader, BM_SETCHECK,
                    state->csv.has_header ? BST_CHECKED : BST_UNCHECKED, 0);

        /* Auto-map fields */
        AutoMapFields(state);

        /* Populate list view */
        PopulateListView(state);
    } else {
        MessageBox(state->hwndMain, L"Failed to load CSV file.", APP_NAME, MB_ICONERROR);
    }
}

static void DoExport(AppState* state) {
    if (state->csv.row_count == 0) {
        MessageBox(state->hwndMain, L"No events to export.", APP_NAME, MB_ICONWARNING);
        return;
    }

    if (state->fieldMapping[ICS_DTSTART] < 0) {
        MessageBox(state->hwndMain, L"DTSTART field must be mapped.", APP_NAME, MB_ICONWARNING);
        return;
    }

    /* Read reminder and travel settings */
    state->defaultReminder1 = (int)SendMessage(state->hwndComboRemind1, CB_GETCURSEL, 0, 0);
    state->defaultReminder2 = (int)SendMessage(state->hwndComboRemind2, CB_GETCURSEL, 0, 0);
    state->appleTravel = (SendMessage(state->hwndChkTravel, BM_GETCHECK, 0, 0) == BST_CHECKED);

    bool single = (SendMessage(state->hwndRadioSingle, BM_GETCHECK, 0, 0) == BST_CHECKED);

    if (single) {
        /* Default to original CSV filename with .ics extension */
        wchar_t filePath[MAX_PATH];
        wcscpy(filePath, state->csvFilePath);
        wchar_t* dot = wcsrchr(filePath, L'.');
        if (dot) wcscpy(dot, L".ics");
        else wcscat(filePath, L".ics");

        OPENFILENAME ofn = {0};
        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner = state->hwndMain;
        ofn.lpstrFilter = L"ICS Files (*.ics)\0*.ics\0All Files (*.*)\0*.*\0";
        ofn.lpstrFile = filePath;
        ofn.nMaxFile = MAX_PATH;
        ofn.Flags = OFN_OVERWRITEPROMPT;
        ofn.lpstrDefExt = L"ics";
        ofn.lpstrTitle = L"Save ICS File";

        if (!GetSaveFileName(&ofn)) return;

        SetWindowText(state->hwndLblStatus, L"Exporting...");
        if (IcsExportSingle(filePath, state)) {
            wchar_t msg[MAX_PATH + 64];
            swprintf(msg, MAX_PATH + 64, L"Exported successfully to:\n%s", filePath);
            SetWindowText(state->hwndLblStatus, msg);
            MessageBox(state->hwndMain, msg, APP_NAME, MB_ICONINFORMATION);
        } else {
            SetWindowText(state->hwndLblStatus, L"Export failed!");
            MessageBox(state->hwndMain, L"Export failed. Check that dates are valid.", APP_NAME, MB_ICONERROR);
        }
    } else {
        /* Folder picker */
        BROWSEINFO bi = {0};
        bi.hwndOwner = state->hwndMain;
        bi.lpszTitle = L"Select folder for ICS files";
        bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;

        LPITEMIDLIST pidl = SHBrowseForFolder(&bi);
        if (!pidl) return;

        wchar_t folderPath[MAX_PATH];
        SHGetPathFromIDList(pidl, folderPath);
        CoTaskMemFree(pidl);

        SetWindowText(state->hwndLblStatus, L"Exporting...");
        if (IcsExportSeparate(folderPath, state)) {
            wchar_t msg[MAX_PATH + 64];
            swprintf(msg, MAX_PATH + 64, L"Exported %d events to:\n%s", state->csv.row_count, folderPath);
            SetWindowText(state->hwndLblStatus, msg);
            MessageBox(state->hwndMain, msg, APP_NAME, MB_ICONINFORMATION);
        } else {
            SetWindowText(state->hwndLblStatus, L"Export failed!");
            MessageBox(state->hwndMain, L"Export failed.", APP_NAME, MB_ICONERROR);
        }
    }
}

/*============================================================================
 *  WINDOW PROCEDURE
 *============================================================================*/

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    AppState* state = (AppState*)GetWindowLongPtr(hwnd, GWLP_USERDATA);

    switch (msg) {
    case WM_CREATE: {
        CREATESTRUCT* cs = (CREATESTRUCT*)lParam;
        state = (AppState*)cs->lpCreateParams;
        SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)state);
        state->hwndMain = hwnd;

        state->hFont = CreateFont(-14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, 0, 0, CLEARTYPE_QUALITY, 0, L"Segoe UI");

        CreatePage1Controls(state, hwnd);
        CreatePage2Controls(state, hwnd);
        CreatePage3Controls(state, hwnd);

        /* Navigation buttons */
        state->hwndBtnBack = CreateBtn(hwnd, L"< Back", CLIENT_WIDTH - 240, NAV_Y, 110, 32, IDC_BTN_BACK, state->hFont);
        state->hwndBtnNext = CreateBtn(hwnd, L"Next >", CLIENT_WIDTH - 120, NAV_Y, 110, 32, IDC_BTN_NEXT, state->hFont);

        /* Add About to system menu */
        HMENU hSysMenu = GetSystemMenu(hwnd, FALSE);
        AppendMenu(hSysMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hSysMenu, MF_STRING, IDM_ABOUT, L"About CSV2ICS...");

        ShowPage(state, PAGE_FILE);
        return 0;
    }

    case WM_SYSCOMMAND:
        if (wParam == IDM_ABOUT) {
            MessageBox(hwnd,
                L"CSV to ICS Converter v1.0\r\n\r\n"
                L"Author: Simon Craig\r\n"
                L"Code: Entirely generated by Claude (Anthropic)\r\n\r\n"
                L"Converts CSV files to RFC 5545 compliant\r\n"
                L"iCalendar (.ics) files.\r\n\r\n"
                L"Features:\r\n"
                L"  \x2022  Auto-detects CSV headers\r\n"
                L"  \x2022  Flexible date/time parsing\r\n"
                L"  \x2022  Field mapping with preview\r\n"
                L"  \x2022  Single or per-event export\r\n"
                L"  \x2022  Reminders (VALARM)\r\n"
                L"  \x2022  Recurrence rules\r\n"
                L"  \x2022  Apple Calendar support\r\n\r\n"
                L"No external dependencies. Single executable\r\n"
                L"built with the Windows SDK and pure C.",
                L"About CSV to ICS Converter",
                MB_OK | MB_ICONINFORMATION);
            return 0;
        }
        break;

    case WM_COMMAND: {
        if (!state) break;
        int id = LOWORD(wParam);
        int code = HIWORD(wParam);

        switch (id) {
        case IDC_BTN_OPEN:
            DoOpenFile(state);
            break;

        case IDC_CHK_HEADER: {
            bool checked = (SendMessage(state->hwndChkHeader, BM_GETCHECK, 0, 0) == BST_CHECKED);
            if (checked != state->csv.has_header) {
                state->csv.has_header = checked;
                PopulateListView(state);
                AutoMapFields(state);
            }
            break;
        }

        case IDC_RADIO_MDY:
            state->dateFormatPref = 0;
            break;

        case IDC_RADIO_DMY:
            state->dateFormatPref = 1;
            break;

        case IDC_RADIO_ISO:
            state->dateFormatPref = 2;
            break;

        case IDC_BTN_NEXT:
            if (state->currentPage == PAGE_FILE) {
                if (state->csv.col_count == 0) {
                    MessageBox(hwnd, L"Please open a CSV file first.", APP_NAME, MB_ICONWARNING);
                    break;
                }
                ShowPage(state, PAGE_MAP);
            } else if (state->currentPage == PAGE_EXPORT) {
                if (MessageBox(hwnd, L"Are you sure you want to close?", APP_NAME,
                               MB_YESNO | MB_ICONQUESTION) == IDYES) {
                    DestroyWindow(hwnd);
                }
                break;
            } else if (state->currentPage == PAGE_MAP) {
                /* Read combo selections */
                for (int i = 0; i < ICS_FIELD_COUNT; i++) {
                    int sel = (int)SendMessage(state->hwndCombo[i], CB_GETCURSEL, 0, 0);
                    if (i == ICS_DTSTART) {
                        state->fieldMapping[i] = (sel >= 0) ? sel : -1;
                    } else {
                        state->fieldMapping[i] = (sel > 0) ? sel - 1 : -1;
                    }
                }

                if (state->fieldMapping[ICS_DTSTART] < 0) {
                    MessageBox(hwnd, L"DTSTART must be mapped to a column.", APP_NAME, MB_ICONWARNING);
                    break;
                }
                ShowPage(state, PAGE_EXPORT);
            }
            break;

        case IDC_BTN_BACK:
            if (state->currentPage == PAGE_MAP) {
                ShowPage(state, PAGE_FILE);
            } else if (state->currentPage == PAGE_EXPORT) {
                ShowPage(state, PAGE_MAP);
            }
            break;

        case IDC_BTN_EXPORT:
            DoExport(state);
            break;

        case IDC_BTN_STARTOVER:
            CsvFree(&state->csv);
            for (int i = 0; i < ICS_FIELD_COUNT; i++) state->fieldMapping[i] = -1;
            state->csvFilePath[0] = L'\0';
            SetWindowText(state->hwndLblPath, L"No file selected");
            ListView_DeleteAllItems(state->hwndListView);
            while (ListView_DeleteColumn(state->hwndListView, 0)) {}
            ShowPage(state, PAGE_FILE);
            break;

        default:
            /* Handle combo box changes for preview update */
            if (code == CBN_SELCHANGE && state->currentPage == PAGE_MAP) {
                for (int i = 0; i < ICS_FIELD_COUNT; i++) {
                    int sel = (int)SendMessage(state->hwndCombo[i], CB_GETCURSEL, 0, 0);
                    if (i == ICS_DTSTART) {
                        state->fieldMapping[i] = (sel >= 0) ? sel : -1;
                    } else {
                        state->fieldMapping[i] = (sel > 0) ? sel - 1 : -1;
                    }
                }
                UpdateIcsPreview(state);
            }
            break;
        }
        return 0;
    }

    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)wParam;
        SetBkMode(hdc, TRANSPARENT);
        return (LRESULT)GetStockObject(WHITE_BRUSH);
    }

    case WM_ERASEBKGND: {
        HDC hdc = (HDC)wParam;
        RECT rc;
        GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, (HBRUSH)GetStockObject(WHITE_BRUSH));
        return 1;
    }

    case WM_DESTROY:
        if (state) {
            CsvFree(&state->csv);
            DeleteObject(state->hFont);
        }
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProc(hwnd, msg, wParam, lParam);
}

/*============================================================================
 *  WINMAIN
 *============================================================================*/

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
                   LPSTR lpCmdLine, int nCmdShow) {
    (void)hPrevInstance;
    (void)lpCmdLine;

    /* DPI awareness */
    SetProcessDPIAware();

    /* Initialize common controls */
    INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_LISTVIEW_CLASSES | ICC_STANDARD_CLASSES };
    InitCommonControlsEx(&icc);

    /* Initialize COM for folder browser */
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    /* Allocate app state */
    AppState* state = (AppState*)calloc(1, sizeof(AppState));
    if (!state) return 1;
    for (int i = 0; i < ICS_FIELD_COUNT; i++) state->fieldMapping[i] = -1;
    state->dateFormatPref = 1; /* DD/MM/YYYY default */

    /* Register window class */
    WNDCLASSEX wc = {0};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)GetStockObject(WHITE_BRUSH);
    wc.lpszClassName = APP_CLASS;
    wc.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(1));
    wc.hIconSm = LoadIcon(hInstance, MAKEINTRESOURCE(1));
    RegisterClassEx(&wc);

    /* Calculate window size for desired client area */
    DWORD dwStyle = WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX;
    RECT rc = { 0, 0, CLIENT_WIDTH, CLIENT_HEIGHT };
    AdjustWindowRect(&rc, dwStyle, FALSE);
    int winW = rc.right - rc.left;
    int winH = rc.bottom - rc.top;

    /* Center window on screen */
    int screenW = GetSystemMetrics(SM_CXSCREEN);
    int screenH = GetSystemMetrics(SM_CYSCREEN);
    int x = (screenW - winW) / 2;
    int y = (screenH - winH) / 2;

    /* Create main window */
    HWND hwnd = CreateWindowEx(0, APP_CLASS, APP_NAME, dwStyle,
        x, y, winW, winH,
        NULL, NULL, hInstance, state);

    if (!hwnd) { free(state); return 1; }

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    /* Message loop */
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0) > 0) {
        if (!IsDialogMessage(hwnd, &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    free(state);
    CoUninitialize();
    return (int)msg.wParam;
}
