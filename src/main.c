/*
 * Jun PDF Tools - Split & Merge
 */

#include <windows.h>
#include <commctrl.h>
#include <commdlg.h>
#include <shlobj.h>
#include <shellapi.h>
#include <stdio.h>

#include "pdf_tools.h"

#pragma comment(lib, "comctl32.lib")

/* Control IDs */
#define ID_TAB_MAIN         100
#define ID_STATIC_STATUS    101

/* Split tab IDs */
#define ID_SPLIT_PDF_PATH   200
#define ID_SPLIT_OUT_PATH   201
#define ID_SPLIT_BTN_PDF    202
#define ID_SPLIT_BTN_OUT    203
#define ID_SPLIT_LIST       204
#define ID_SPLIT_NAME       205
#define ID_SPLIT_START      206
#define ID_SPLIT_END        207
#define ID_SPLIT_BTN_ADD    208
#define ID_SPLIT_BTN_DEL    209
#define ID_SPLIT_BTN_CLR    210
#define ID_SPLIT_BTN_RUN    211
#define ID_SPLIT_PAGE_INFO  212
#define ID_SPLIT_PROGRESS   213

/* Merge tab IDs */
#define ID_MERGE_LIST       300
#define ID_MERGE_BTN_ADD    301
#define ID_MERGE_BTN_DEL    302
#define ID_MERGE_BTN_UP     303
#define ID_MERGE_BTN_DOWN   304
#define ID_MERGE_BTN_RUN    305
#define ID_MERGE_OUT_PATH   306
#define ID_MERGE_BTN_OUT    307
#define ID_MERGE_PROGRESS   308

#define MAX_CHAPTERS        100
#define MAX_MERGE_FILES     50
#define NAME_LENGTH         64

#define TAB_SPLIT           0
#define TAB_MERGE           1
#define TAB_COUNT           2

typedef struct chapter {
    WCHAR name[NAME_LENGTH];
    int start_page;
    int end_page;
} chapter_t;

/* Main window */
static HWND s_hwnd_main;
static HWND s_hwnd_tab;
static HWND s_hwnd_status;
static HFONT s_hfont_ui;
static HFONT s_hfont_title;
static int s_current_tab = 0;
static float s_dpi_scale = 1.0f;

/* DPI 스케일링 함수 */
static int dpi(int value)
{
    return (int)(value * s_dpi_scale + 0.5f);
}

/* Split tab */
static HWND s_split_ctrls[32];
static int s_split_ctrl_count = 0;
static HWND s_hwnd_split_pdf_path;
static HWND s_hwnd_split_out_path;
static HWND s_hwnd_split_list;
static HWND s_hwnd_split_name;
static HWND s_hwnd_split_start;
static HWND s_hwnd_split_end;
static HWND s_hwnd_split_page_info;
static HWND s_hwnd_split_btn_run;
static HWND s_hwnd_split_progress;
static WCHAR s_split_pdf_path[MAX_PATH];
static WCHAR s_split_out_path[MAX_PATH];
static chapter_t s_chapters[MAX_CHAPTERS];
static int s_chapter_count = 0;
static int s_split_total_pages = 0;
static WNDPROC s_orig_edit_proc;

/* Merge tab */
static HWND s_merge_ctrls[20];
static int s_merge_ctrl_count = 0;
static HWND s_hwnd_merge_list;
static HWND s_hwnd_merge_out_path;
static HWND s_hwnd_merge_btn_run;
static HWND s_hwnd_merge_progress;
static WCHAR s_merge_files[MAX_MERGE_FILES][MAX_PATH];
static int s_merge_file_count = 0;
static WCHAR s_merge_out_path[MAX_PATH];

/* Subclass procedure for edit controls (Enter/Tab/Shift+Tab handling) */
static LRESULT CALLBACK edit_subclass_proc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam)
{
    if (msg == WM_KEYDOWN) {
        /* Enter key */
        if (wparam == VK_RETURN) {
            if (hwnd == s_hwnd_split_name) {
                SetFocus(s_hwnd_split_start);
                return 0;
            } else if (hwnd == s_hwnd_split_start) {
                SetFocus(s_hwnd_split_end);
                return 0;
            } else if (hwnd == s_hwnd_split_end) {
                SendMessage(s_hwnd_main, WM_COMMAND, ID_SPLIT_BTN_ADD, 0);
                return 0;
            }
        }
        /* Tab key */
        if (wparam == VK_TAB) {
            int shift_pressed = GetKeyState(VK_SHIFT) & 0x8000;
            if (shift_pressed) {
                /* Shift+Tab: backward */
                if (hwnd == s_hwnd_split_end) {
                    SetFocus(s_hwnd_split_start);
                    return 0;
                } else if (hwnd == s_hwnd_split_start) {
                    SetFocus(s_hwnd_split_name);
                    return 0;
                }
            } else {
                /* Tab: forward */
                if (hwnd == s_hwnd_split_name) {
                    SetFocus(s_hwnd_split_start);
                    return 0;
                } else if (hwnd == s_hwnd_split_start) {
                    SetFocus(s_hwnd_split_end);
                    return 0;
                }
            }
        }
    }
    /* WM_CHAR에서 Enter/Tab 문자 무시 (경고음 방지) */
    if (msg == WM_CHAR) {
        if (wparam == '\r' || wparam == '\n' || wparam == '\t') {
            return 0;
        }
    }
    return CallWindowProc(s_orig_edit_proc, hwnd, msg, wparam, lparam);
}

/* Function declarations */
static LRESULT CALLBACK window_proc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam);
static void create_fonts(void);
static void create_controls(HWND hwnd);
static void create_split_tab(HWND hwnd, HINSTANCE hinst);
static void create_merge_tab(HWND hwnd, HINSTANCE hinst);
static void show_tab(int tab_index);
static void set_control_font(HWND hwnd, HFONT font);
static void update_status(const WCHAR* message);
static void handle_drop_files(HDROP hdrop);

/* Split functions */
static void split_select_pdf(HWND hwnd);
static void split_select_output(HWND hwnd);
static void split_add_chapter(HWND hwnd);
static void split_del_chapter(void);
static void split_clear_chapters(void);
static void split_refresh_list(void);
static void split_run(HWND hwnd);
static void split_load_pdf(const WCHAR* path);

/* Merge functions */
static void merge_add_files(HWND hwnd);
static void merge_add_file(const WCHAR* path);
static void merge_del_file(void);
static void merge_move_up(void);
static void merge_move_down(void);
static void merge_select_output(HWND hwnd);
static void merge_run(HWND hwnd);
static void merge_refresh_list(void);

int WINAPI wWinMain(HINSTANCE hinstance, HINSTANCE hprev_instance, LPWSTR cmd_line, int cmd_show)
{
    WNDCLASSEXW wc;
    MSG msg;
    INITCOMMONCONTROLSEX icex;
    HDC hdc;

    (void)hprev_instance;
    (void)cmd_line;

    /* DPI Awareness 설정 - 고해상도 디스플레이에서 선명하게 표시 */
    SetProcessDPIAware();

    /* DPI 스케일 계산 (기준: 96 DPI) */
    hdc = GetDC(NULL);
    s_dpi_scale = GetDeviceCaps(hdc, LOGPIXELSX) / 96.0f;
    ReleaseDC(NULL, hdc);

    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_TAB_CLASSES | ICC_LISTVIEW_CLASSES | ICC_STANDARD_CLASSES;
    InitCommonControlsEx(&icex);

    create_fonts();

    memset(&wc, 0, sizeof(wc));
    wc.cbSize = sizeof(WNDCLASSEXW);
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = window_proc;
    wc.hInstance = hinstance;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = L"JunPdfTools";
    wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);

    RegisterClassExW(&wc);

    s_hwnd_main = CreateWindowExW(
        WS_EX_ACCEPTFILES,
        L"JunPdfTools", L"Jun PDF Tools",
        WS_OVERLAPPEDWINDOW & ~WS_THICKFRAME & ~WS_MAXIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, dpi(720), dpi(560),
        NULL, NULL, hinstance, NULL
    );

    ShowWindow(s_hwnd_main, cmd_show);
    UpdateWindow(s_hwnd_main);

    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    DeleteObject(s_hfont_ui);
    DeleteObject(s_hfont_title);

    return (int)msg.wParam;
}

static void create_fonts(void)
{
    s_hfont_ui = CreateFontW(-dpi(14), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"맑은 고딕");

    s_hfont_title = CreateFontW(-dpi(16), 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"맑은 고딕");
}

static void set_control_font(HWND hwnd, HFONT font)
{
    SendMessage(hwnd, WM_SETFONT, (WPARAM)font, TRUE);
}

static LRESULT CALLBACK window_proc(HWND hwnd, UINT msg, WPARAM wparam, LPARAM lparam)
{
    NMHDR* nmhdr;

    switch (msg) {
    case WM_CREATE:
        create_controls(hwnd);
        break;

    case WM_DROPFILES:
        handle_drop_files((HDROP)wparam);
        break;

    case WM_NOTIFY:
        nmhdr = (NMHDR*)lparam;
        if (nmhdr->idFrom == ID_TAB_MAIN && nmhdr->code == TCN_SELCHANGE) {
            show_tab(TabCtrl_GetCurSel(s_hwnd_tab));
        }
        break;

    case WM_COMMAND:
        switch (LOWORD(wparam)) {
        /* Split tab */
        case ID_SPLIT_BTN_PDF: split_select_pdf(hwnd); break;
        case ID_SPLIT_BTN_OUT: split_select_output(hwnd); break;
        case ID_SPLIT_BTN_ADD: split_add_chapter(hwnd); break;
        case ID_SPLIT_BTN_DEL: split_del_chapter(); break;
        case ID_SPLIT_BTN_CLR: split_clear_chapters(); break;
        case ID_SPLIT_BTN_RUN: split_run(hwnd); break;
        /* Merge tab */
        case ID_MERGE_BTN_ADD: merge_add_files(hwnd); break;
        case ID_MERGE_BTN_DEL: merge_del_file(); break;
        case ID_MERGE_BTN_UP: merge_move_up(); break;
        case ID_MERGE_BTN_DOWN: merge_move_down(); break;
        case ID_MERGE_BTN_OUT: merge_select_output(hwnd); break;
        case ID_MERGE_BTN_RUN: merge_run(hwnd); break;
        }
        break;

    case WM_CTLCOLORSTATIC:
        {
            int ctrl_id = GetDlgCtrlID((HWND)lparam);
            /* Read-only edit controls need white background */
            if (ctrl_id == ID_SPLIT_PDF_PATH || ctrl_id == ID_SPLIT_OUT_PATH ||
                ctrl_id == ID_MERGE_OUT_PATH) {
                SetBkColor((HDC)wparam, RGB(255, 255, 255));
                return (LRESULT)GetStockObject(WHITE_BRUSH);
            }
            /* Page info label needs proper background to clear on update */
            if (ctrl_id == ID_SPLIT_PAGE_INFO) {
                SetBkColor((HDC)wparam, GetSysColor(COLOR_BTNFACE));
                return (LRESULT)GetSysColorBrush(COLOR_BTNFACE);
            }
        }
        SetBkMode((HDC)wparam, TRANSPARENT);
        return (LRESULT)GetStockObject(NULL_BRUSH);

    case WM_DESTROY:
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hwnd, msg, wparam, lparam);
    }
    return 0;
}

static void handle_drop_files(HDROP hdrop)
{
    UINT file_count;
    UINT i;
    WCHAR file_path[MAX_PATH];
    WCHAR* ext;
    WCHAR* filename;
    int pdf_count = 0;
    int non_pdf_count = 0;
    WCHAR first_non_pdf[MAX_PATH] = L"";

    file_count = DragQueryFileW(hdrop, 0xFFFFFFFF, NULL, 0);

    /* 먼저 PDF 파일 개수 확인 */
    for (i = 0; i < file_count; i++) {
        DragQueryFileW(hdrop, i, file_path, MAX_PATH);
        ext = wcsrchr(file_path, L'.');

        if (ext && _wcsicmp(ext, L".pdf") == 0) {
            pdf_count++;
        } else {
            non_pdf_count++;
            if (wcslen(first_non_pdf) == 0) {
                filename = wcsrchr(file_path, L'\\');
                wcscpy_s(first_non_pdf, MAX_PATH, filename ? filename + 1 : file_path);
            }
        }
    }

    /* 비PDF 파일이 있으면 알림 */
    if (non_pdf_count > 0) {
        WCHAR msg[512];
        if (non_pdf_count == 1) {
            swprintf_s(msg, 512, L"PDF 파일만 지원합니다.\n\n지원하지 않는 파일: %s", first_non_pdf);
        } else {
            swprintf_s(msg, 512, L"PDF 파일만 지원합니다.\n\n지원하지 않는 파일 %d개가 무시되었습니다.", non_pdf_count);
        }
        MessageBoxW(s_hwnd_main, msg, L"파일 형식 오류", MB_OK | MB_ICONWARNING);
    }

    /* 분할 탭에서 여러 PDF 파일 드롭 시 알림 */
    if (s_current_tab == TAB_SPLIT && pdf_count > 1) {
        MessageBoxW(s_hwnd_main, L"분할 모드에서는 1개의 PDF 파일만 선택할 수 있습니다.\n\n첫 번째 파일만 로드됩니다.", L"알림", MB_OK | MB_ICONINFORMATION);
    }

    /* PDF 파일 처리 */
    for (i = 0; i < file_count; i++) {
        DragQueryFileW(hdrop, i, file_path, MAX_PATH);
        ext = wcsrchr(file_path, L'.');

        if (ext && _wcsicmp(ext, L".pdf") == 0) {
            if (s_current_tab == TAB_SPLIT) {
                split_load_pdf(file_path);
                break;  /* 분할 모드에서는 첫 번째 PDF만 로드 */
            } else if (s_current_tab == TAB_MERGE) {
                merge_add_file(file_path);
            }
        }
    }

    DragFinish(hdrop);
}

static void create_controls(HWND hwnd)
{
    HINSTANCE hinst;
    TCITEMW tie;

    hinst = (HINSTANCE)GetWindowLongPtr(hwnd, GWLP_HINSTANCE);

    s_hwnd_tab = CreateWindowExW(0, WC_TABCONTROLW, L"",
        WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS,
        0, 0, dpi(704), dpi(521), hwnd, (HMENU)ID_TAB_MAIN, hinst, NULL);
    set_control_font(s_hwnd_tab, s_hfont_ui);

    memset(&tie, 0, sizeof(tie));
    tie.mask = TCIF_TEXT;
    tie.pszText = L"   분할   "; TabCtrl_InsertItem(s_hwnd_tab, TAB_SPLIT, &tie);
    tie.pszText = L"   병합   "; TabCtrl_InsertItem(s_hwnd_tab, TAB_MERGE, &tie);

    s_hwnd_status = CreateWindowW(L"STATIC", L"PDF 파일을 드래그하거나 선택하세요",
        WS_CHILD | WS_VISIBLE | SS_LEFT,
        dpi(10), dpi(490), dpi(684), dpi(20), hwnd, (HMENU)ID_STATIC_STATUS, hinst, NULL);
    set_control_font(s_hwnd_status, s_hfont_ui);

    create_split_tab(hwnd, hinst);
    create_merge_tab(hwnd, hinst);

    show_tab(TAB_SPLIT);
}

static void show_tab(int tab_index)
{
    int i;
    int show_split = (tab_index == TAB_SPLIT) ? SW_SHOW : SW_HIDE;
    int show_merge = (tab_index == TAB_MERGE) ? SW_SHOW : SW_HIDE;

    for (i = 0; i < s_split_ctrl_count; i++) {
        ShowWindow(s_split_ctrls[i], show_split);
    }
    for (i = 0; i < s_merge_ctrl_count; i++) {
        ShowWindow(s_merge_ctrls[i], show_merge);
    }

    /* Always hide progress bars on tab switch (they show only during operation) */
    ShowWindow(s_hwnd_split_progress, SW_HIDE);
    ShowWindow(s_hwnd_merge_progress, SW_HIDE);

    s_current_tab = tab_index;

    if (tab_index == TAB_SPLIT) {
        update_status(L"PDF 파일을 드래그하거나 선택하세요");
    } else {
        update_status(L"병합할 PDF 파일들을 드래그하거나 추가하세요");
    }
}

static void update_status(const WCHAR* message)
{
    RECT rc;
    GetWindowRect(s_hwnd_status, &rc);
    MapWindowPoints(HWND_DESKTOP, s_hwnd_main, (LPPOINT)&rc, 2);
    InvalidateRect(s_hwnd_main, &rc, TRUE);
    SetWindowTextW(s_hwnd_status, message);
    UpdateWindow(s_hwnd_main);
}

/* ==================== Split Tab ==================== */
static void create_split_tab(HWND hwnd, HINSTANCE hinst)
{
    int y = dpi(38);
    int lm = dpi(10);
    int cw = dpi(684);
    HWND h;

    #define ADD_SPLIT_CTRL(x) s_split_ctrls[s_split_ctrl_count++] = (x)

    h = CreateWindowW(L"STATIC", L"PDF 분할", WS_CHILD | WS_VISIBLE,
        lm, y, dpi(200), dpi(22), hwnd, NULL, hinst, NULL);
    set_control_font(h, s_hfont_title); ADD_SPLIT_CTRL(h);
    y += dpi(28);

    h = CreateWindowW(L"STATIC", L"PDF 파일", WS_CHILD | WS_VISIBLE,
        lm, y + dpi(3), dpi(70), dpi(20), hwnd, NULL, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_SPLIT_CTRL(h);

    s_hwnd_split_pdf_path = CreateWindowW(L"EDIT", L"",
        WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
        lm + dpi(75), y, cw - dpi(175), dpi(24), hwnd, (HMENU)ID_SPLIT_PDF_PATH, hinst, NULL);
    set_control_font(s_hwnd_split_pdf_path, s_hfont_ui);
    SendMessage(s_hwnd_split_pdf_path, EM_SETREADONLY, TRUE, 0);
    ADD_SPLIT_CTRL(s_hwnd_split_pdf_path);

    h = CreateWindowW(L"BUTTON", L"파일 선택...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        lm + cw - dpi(95), y - dpi(1), dpi(95), dpi(26), hwnd, (HMENU)ID_SPLIT_BTN_PDF, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_SPLIT_CTRL(h);
    y += dpi(28);

    s_hwnd_split_page_info = CreateWindowW(L"STATIC", L"", WS_CHILD | WS_VISIBLE,
        lm + dpi(75), y, dpi(200), dpi(18), hwnd, (HMENU)ID_SPLIT_PAGE_INFO, hinst, NULL);
    set_control_font(s_hwnd_split_page_info, s_hfont_ui); ADD_SPLIT_CTRL(s_hwnd_split_page_info);
    y += dpi(22);

    h = CreateWindowW(L"STATIC", L"출력 폴더", WS_CHILD | WS_VISIBLE,
        lm, y + dpi(3), dpi(70), dpi(20), hwnd, NULL, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_SPLIT_CTRL(h);

    s_hwnd_split_out_path = CreateWindowW(L"EDIT", L"",
        WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
        lm + dpi(75), y, cw - dpi(175), dpi(24), hwnd, (HMENU)ID_SPLIT_OUT_PATH, hinst, NULL);
    set_control_font(s_hwnd_split_out_path, s_hfont_ui);
    SendMessage(s_hwnd_split_out_path, EM_SETREADONLY, TRUE, 0);
    ADD_SPLIT_CTRL(s_hwnd_split_out_path);

    h = CreateWindowW(L"BUTTON", L"폴더 선택...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        lm + cw - dpi(95), y - dpi(1), dpi(95), dpi(26), hwnd, (HMENU)ID_SPLIT_BTN_OUT, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_SPLIT_CTRL(h);
    y += dpi(35);

    h = CreateWindowW(L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_ETCHEDHORZ,
        lm, y, cw, dpi(2), hwnd, NULL, hinst, NULL);
    ADD_SPLIT_CTRL(h);
    y += dpi(12);

    h = CreateWindowW(L"STATIC", L"챕터 추가", WS_CHILD | WS_VISIBLE,
        lm, y, dpi(100), dpi(20), hwnd, NULL, hinst, NULL);
    set_control_font(h, s_hfont_title); ADD_SPLIT_CTRL(h);
    y += dpi(25);

    h = CreateWindowW(L"STATIC", L"이름", WS_CHILD | WS_VISIBLE,
        lm, y + dpi(3), dpi(35), dpi(20), hwnd, NULL, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_SPLIT_CTRL(h);

    s_hwnd_split_name = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
        WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
        lm + dpi(40), y, dpi(130), dpi(24), hwnd, (HMENU)ID_SPLIT_NAME, hinst, NULL);
    set_control_font(s_hwnd_split_name, s_hfont_ui); ADD_SPLIT_CTRL(s_hwnd_split_name);

    h = CreateWindowW(L"STATIC", L"시작", WS_CHILD | WS_VISIBLE,
        lm + dpi(185), y + dpi(3), dpi(35), dpi(20), hwnd, NULL, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_SPLIT_CTRL(h);

    s_hwnd_split_start = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
        WS_CHILD | WS_VISIBLE | ES_NUMBER,
        lm + dpi(225), y, dpi(60), dpi(24), hwnd, (HMENU)ID_SPLIT_START, hinst, NULL);
    set_control_font(s_hwnd_split_start, s_hfont_ui); ADD_SPLIT_CTRL(s_hwnd_split_start);

    h = CreateWindowW(L"STATIC", L"끝", WS_CHILD | WS_VISIBLE,
        lm + dpi(300), y + dpi(3), dpi(25), dpi(20), hwnd, NULL, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_SPLIT_CTRL(h);

    s_hwnd_split_end = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
        WS_CHILD | WS_VISIBLE | ES_NUMBER,
        lm + dpi(330), y, dpi(60), dpi(24), hwnd, (HMENU)ID_SPLIT_END, hinst, NULL);
    set_control_font(s_hwnd_split_end, s_hfont_ui); ADD_SPLIT_CTRL(s_hwnd_split_end);

    h = CreateWindowW(L"BUTTON", L"추가", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        lm + dpi(410), y - dpi(1), dpi(70), dpi(26), hwnd, (HMENU)ID_SPLIT_BTN_ADD, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_SPLIT_CTRL(h);

    h = CreateWindowW(L"BUTTON", L"삭제", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        lm + dpi(490), y - dpi(1), dpi(70), dpi(26), hwnd, (HMENU)ID_SPLIT_BTN_DEL, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_SPLIT_CTRL(h);

    h = CreateWindowW(L"BUTTON", L"초기화", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        lm + dpi(570), y - dpi(1), dpi(70), dpi(26), hwnd, (HMENU)ID_SPLIT_BTN_CLR, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_SPLIT_CTRL(h);
    y += dpi(35);

    h = CreateWindowW(L"STATIC", L"챕터 목록", WS_CHILD | WS_VISIBLE,
        lm, y, dpi(100), dpi(20), hwnd, NULL, hinst, NULL);
    set_control_font(h, s_hfont_title); ADD_SPLIT_CTRL(h);
    y += dpi(22);

    s_hwnd_split_list = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | LBS_NOINTEGRALHEIGHT | LBS_NOTIFY,
        lm, y, cw, dpi(150), hwnd, (HMENU)ID_SPLIT_LIST, hinst, NULL);
    set_control_font(s_hwnd_split_list, s_hfont_ui); ADD_SPLIT_CTRL(s_hwnd_split_list);
    y += dpi(158);

    /* Progress bar */
    s_hwnd_split_progress = CreateWindowExW(0, PROGRESS_CLASSW, L"",
        WS_CHILD | PBS_SMOOTH,
        lm, y, cw, dpi(20), hwnd, (HMENU)ID_SPLIT_PROGRESS, hinst, NULL);
    ADD_SPLIT_CTRL(s_hwnd_split_progress);
    y += dpi(28);

    s_hwnd_split_btn_run = CreateWindowW(L"BUTTON", L"PDF 분할 실행",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        lm, y, cw, dpi(36), hwnd, (HMENU)ID_SPLIT_BTN_RUN, hinst, NULL);
    set_control_font(s_hwnd_split_btn_run, s_hfont_title); ADD_SPLIT_CTRL(s_hwnd_split_btn_run);

    /* Subclass edit controls for Enter key */
    s_orig_edit_proc = (WNDPROC)SetWindowLongPtr(s_hwnd_split_name, GWLP_WNDPROC, (LONG_PTR)edit_subclass_proc);
    SetWindowLongPtr(s_hwnd_split_start, GWLP_WNDPROC, (LONG_PTR)edit_subclass_proc);
    SetWindowLongPtr(s_hwnd_split_end, GWLP_WNDPROC, (LONG_PTR)edit_subclass_proc);

    #undef ADD_SPLIT_CTRL
}

static void split_load_pdf(const WCHAR* path)
{
    WCHAR* last_slash;
    WCHAR msg[256];
    int page_count;
    HWND hwnd_pdf;
    HWND hwnd_out;
    pdf_error_t error = PDF_OK;

    /* Clear chapter list when loading new PDF */
    s_chapter_count = 0;
    split_refresh_list();
    SetWindowTextW(s_hwnd_split_name, L"");
    SetWindowTextW(s_hwnd_split_start, L"");
    SetWindowTextW(s_hwnd_split_end, L"");

    /* Reset progress bar */
    ShowWindow(s_hwnd_split_progress, SW_HIDE);
    SendMessageW(s_hwnd_split_progress, PBM_SETPOS, 0, 0);

    /* Use GetDlgItem to get controls by ID */
    hwnd_pdf = GetDlgItem(s_hwnd_main, ID_SPLIT_PDF_PATH);
    hwnd_out = GetDlgItem(s_hwnd_main, ID_SPLIT_OUT_PATH);

    wcscpy_s(s_split_pdf_path, MAX_PATH, path);
    if (hwnd_pdf) {
        SetWindowTextW(hwnd_pdf, s_split_pdf_path);
    }

    wcscpy_s(s_split_out_path, MAX_PATH, path);
    last_slash = wcsrchr(s_split_out_path, L'\\');
    if (last_slash) {
        *last_slash = L'\0';
    }
    if (hwnd_out) {
        SetWindowTextW(hwnd_out, s_split_out_path);
    }

    page_count = pdf_get_page_count(s_split_pdf_path, &error);
    if (page_count > 0) {
        s_split_total_pages = page_count;
        swprintf_s(msg, 256, L"총 %d 페이지", page_count);
        SetWindowTextW(s_hwnd_split_page_info, msg);
        update_status(L"PDF 로드 완료");
    } else {
        s_split_total_pages = 0;
        SetWindowTextW(s_hwnd_split_page_info, L"");
        /* 구체적인 오류 메시지 표시 */
        update_status(pdf_error_message(error));
        MessageBoxW(s_hwnd_main, pdf_error_message(error), L"PDF 로드 오류", MB_OK | MB_ICONERROR);
    }
}

static void split_select_pdf(HWND hwnd)
{
    OPENFILENAMEW ofn;
    WCHAR file_path[MAX_PATH];

    memset(file_path, 0, sizeof(file_path));
    memset(&ofn, 0, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFile = file_path;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFilter = L"PDF 파일 (*.pdf)\0*.pdf\0";
    ofn.Flags = OFN_FILEMUSTEXIST;

    if (GetOpenFileNameW(&ofn)) {
        split_load_pdf(file_path);
    }
}

static void split_select_output(HWND hwnd)
{
    BROWSEINFOW bi;
    PIDLIST_ABSOLUTE pidl;
    WCHAR folder_path[MAX_PATH];
    HWND hwnd_out;

    memset(&bi, 0, sizeof(bi));
    bi.hwndOwner = hwnd;
    bi.lpszTitle = L"출력 폴더 선택";
    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;

    pidl = SHBrowseForFolderW(&bi);
    if (pidl) {
        if (SHGetPathFromIDListW(pidl, folder_path)) {
            wcscpy_s(s_split_out_path, MAX_PATH, folder_path);
            hwnd_out = GetDlgItem(s_hwnd_main, ID_SPLIT_OUT_PATH);
            if (hwnd_out) {
                SetWindowTextW(hwnd_out, s_split_out_path);
            }
        }
        CoTaskMemFree(pidl);
    }
}

static void split_add_chapter(HWND hwnd)
{
    WCHAR name[NAME_LENGTH], start_str[16], end_str[16];
    WCHAR msg[256];
    int start, end;

    if (s_chapter_count >= MAX_CHAPTERS) {
        MessageBoxW(hwnd, L"최대 100개까지 추가 가능합니다.", L"알림", MB_OK | MB_ICONWARNING);
        return;
    }

    GetWindowTextW(s_hwnd_split_name, name, NAME_LENGTH);
    GetWindowTextW(s_hwnd_split_start, start_str, 16);
    GetWindowTextW(s_hwnd_split_end, end_str, 16);
    start = _wtoi(start_str);
    end = _wtoi(end_str);

    /* 입력값 유효성 검사 - 구체적인 오류 메시지 */
    if (wcslen(name) == 0) {
        MessageBoxW(hwnd, L"챕터 이름을 입력해주세요.", L"입력 오류", MB_OK | MB_ICONWARNING);
        SetFocus(s_hwnd_split_name);
        return;
    }

    if (wcslen(start_str) == 0 || start <= 0) {
        MessageBoxW(hwnd, L"시작 페이지를 1 이상의 숫자로 입력해주세요.", L"입력 오류", MB_OK | MB_ICONWARNING);
        SetFocus(s_hwnd_split_start);
        return;
    }

    if (wcslen(end_str) == 0 || end <= 0) {
        MessageBoxW(hwnd, L"끝 페이지를 1 이상의 숫자로 입력해주세요.", L"입력 오류", MB_OK | MB_ICONWARNING);
        SetFocus(s_hwnd_split_end);
        return;
    }

    if (end < start) {
        swprintf_s(msg, 256, L"끝 페이지(%d)가 시작 페이지(%d)보다 작습니다.\n페이지 범위를 확인해주세요.", end, start);
        MessageBoxW(hwnd, msg, L"입력 오류", MB_OK | MB_ICONWARNING);
        SetFocus(s_hwnd_split_end);
        return;
    }

    /* PDF가 로드된 경우 페이지 범위 검증 */
    if (s_split_total_pages > 0) {
        if (start > s_split_total_pages) {
            swprintf_s(msg, 256, L"시작 페이지(%d)가 총 페이지 수(%d)를 초과합니다.", start, s_split_total_pages);
            MessageBoxW(hwnd, msg, L"페이지 범위 오류", MB_OK | MB_ICONWARNING);
            SetFocus(s_hwnd_split_start);
            return;
        }
        if (end > s_split_total_pages) {
            swprintf_s(msg, 256, L"끝 페이지(%d)가 총 페이지 수(%d)를 초과합니다.", end, s_split_total_pages);
            MessageBoxW(hwnd, msg, L"페이지 범위 오류", MB_OK | MB_ICONWARNING);
            SetFocus(s_hwnd_split_end);
            return;
        }
    }

    wcscpy_s(s_chapters[s_chapter_count].name, NAME_LENGTH, name);
    s_chapters[s_chapter_count].start_page = start;
    s_chapters[s_chapter_count].end_page = end;
    s_chapter_count++;

    split_refresh_list();

    SetWindowTextW(s_hwnd_split_name, L"");
    SetWindowTextW(s_hwnd_split_start, L"");
    SetWindowTextW(s_hwnd_split_end, L"");
    SetFocus(s_hwnd_split_name);
}

static void split_del_chapter(void)
{
    int sel = (int)SendMessageW(s_hwnd_split_list, LB_GETCURSEL, 0, 0);
    int i;
    if (sel == LB_ERR) return;
    for (i = sel; i < s_chapter_count - 1; i++) {
        s_chapters[i] = s_chapters[i + 1];
    }
    s_chapter_count--;
    split_refresh_list();

    /* Select next item (or last item if deleted the last one) */
    if (s_chapter_count > 0) {
        if (sel >= s_chapter_count) {
            sel = s_chapter_count - 1;
        }
        SendMessageW(s_hwnd_split_list, LB_SETCURSEL, sel, 0);
    }
}

static void split_clear_chapters(void)
{
    s_chapter_count = 0;
    split_refresh_list();
    SetWindowTextW(s_hwnd_split_name, L"");
    SetWindowTextW(s_hwnd_split_start, L"");
    SetWindowTextW(s_hwnd_split_end, L"");
}

static void split_refresh_list(void)
{
    int i;
    WCHAR item[128];
    SendMessageW(s_hwnd_split_list, LB_RESETCONTENT, 0, 0);
    for (i = 0; i < s_chapter_count; i++) {
        swprintf_s(item, 128, L"  %s    (페이지 %d ~ %d)",
            s_chapters[i].name, s_chapters[i].start_page, s_chapters[i].end_page);
        SendMessageW(s_hwnd_split_list, LB_ADDSTRING, 0, (LPARAM)item);
    }
}

/* 실패한 챕터 정보 저장 구조체 */
typedef struct failed_chapter {
    WCHAR name[NAME_LENGTH];
    int start_page;
    int end_page;
    pdf_error_t error;
} failed_chapter_t;

static void split_run(HWND hwnd)
{
    int i, success = 0, existing_count = 0, fail_count = 0;
    WCHAR out_path[MAX_PATH], msg[1024];
    MSG winmsg;
    pdf_error_t error;
    failed_chapter_t failed_chapters[MAX_CHAPTERS];

    if (wcslen(s_split_pdf_path) == 0) {
        MessageBoxW(hwnd, L"PDF 파일을 선택하세요.", L"오류", MB_OK | MB_ICONERROR);
        return;
    }
    if (s_chapter_count == 0) {
        MessageBoxW(hwnd, L"챕터를 추가하세요.", L"오류", MB_OK | MB_ICONERROR);
        return;
    }

    /* Check for existing files */
    {
        WCHAR existing_files[1024] = L"";
        for (i = 0; i < s_chapter_count; i++) {
            swprintf_s(out_path, MAX_PATH, L"%s\\%s.pdf", s_split_out_path, s_chapters[i].name);
            if (GetFileAttributesW(out_path) != INVALID_FILE_ATTRIBUTES) {
                if (existing_count > 0) {
                    wcscat_s(existing_files, 1024, L", ");
                }
                wcscat_s(existing_files, 1024, s_chapters[i].name);
                wcscat_s(existing_files, 1024, L".pdf");
                existing_count++;
            }
        }
        if (existing_count > 0) {
            swprintf_s(msg, 1024, L"다음 파일이 이미 존재합니다:\n%s\n\n덮어쓰시겠습니까?", existing_files);
            if (MessageBoxW(hwnd, msg, L"확인", MB_YESNO | MB_ICONQUESTION) != IDYES) {
                return;
            }
        }
    }

    EnableWindow(s_hwnd_split_btn_run, FALSE);

    /* Show and setup progress bar */
    ShowWindow(s_hwnd_split_progress, SW_SHOW);
    SendMessageW(s_hwnd_split_progress, PBM_SETRANGE32, 0, s_chapter_count);
    SendMessageW(s_hwnd_split_progress, PBM_SETPOS, 0, 0);

    for (i = 0; i < s_chapter_count; i++) {
        /* Update progress */
        SendMessageW(s_hwnd_split_progress, PBM_SETPOS, i + 1, 0);
        swprintf_s(msg, 1024, L"분할 중... (%d/%d)", i + 1, s_chapter_count);
        update_status(msg);

        /* Process messages to keep UI responsive */
        while (PeekMessage(&winmsg, NULL, 0, 0, PM_REMOVE)) {
            TranslateMessage(&winmsg);
            DispatchMessage(&winmsg);
        }

        swprintf_s(out_path, MAX_PATH, L"%s\\%s.pdf", s_split_out_path, s_chapters[i].name);
        error = PDF_OK;
        if (pdf_split(s_split_pdf_path, out_path, s_chapters[i].start_page, s_chapters[i].end_page, &error)) {
            success++;
        } else {
            /* 실패한 챕터 정보 저장 */
            if (fail_count < MAX_CHAPTERS) {
                wcscpy_s(failed_chapters[fail_count].name, NAME_LENGTH, s_chapters[i].name);
                failed_chapters[fail_count].start_page = s_chapters[i].start_page;
                failed_chapters[fail_count].end_page = s_chapters[i].end_page;
                failed_chapters[fail_count].error = error;
                fail_count++;
            }
        }
    }

    /* Hide progress bar */
    ShowWindow(s_hwnd_split_progress, SW_HIDE);
    EnableWindow(s_hwnd_split_btn_run, TRUE);

    swprintf_s(msg, 256, L"완료: %d/%d 챕터 분할됨", success, s_chapter_count);
    update_status(msg);

    /* 결과 표시 */
    if (fail_count > 0) {
        WCHAR result_msg[2048];
        int offset;

        if (success > 0) {
            offset = swprintf_s(result_msg, 2048, L"분할 완료\n\n성공: %d/%d 챕터\n실패: %d개\n\n",
                               success, s_chapter_count, fail_count);
        } else {
            offset = swprintf_s(result_msg, 2048, L"분할 실패\n\n모든 챕터(%d개)가 실패했습니다.\n\n",
                               s_chapter_count);
        }

        /* 실패한 챕터 목록 (최대 5개까지만 표시) */
        wcscat_s(result_msg, 2048, L"실패한 챕터:\n");
        for (i = 0; i < fail_count && i < 5; i++) {
            WCHAR fail_item[256];
            swprintf_s(fail_item, 256, L"  - %s (페이지 %d~%d): %s\n",
                      failed_chapters[i].name,
                      failed_chapters[i].start_page,
                      failed_chapters[i].end_page,
                      pdf_error_message(failed_chapters[i].error));
            wcscat_s(result_msg, 2048, fail_item);
        }
        if (fail_count > 5) {
            WCHAR more_msg[64];
            swprintf_s(more_msg, 64, L"  ... 외 %d개\n", fail_count - 5);
            wcscat_s(result_msg, 2048, more_msg);
        }

        MessageBoxW(hwnd, result_msg, success > 0 ? L"분할 완료" : L"분할 실패",
                    success > 0 ? MB_OK | MB_ICONWARNING : MB_OK | MB_ICONERROR);

        if (success > 0 && MessageBoxW(hwnd, L"폴더를 열까요?", L"확인", MB_YESNO) == IDYES) {
            ShellExecuteW(NULL, L"open", s_split_out_path, NULL, NULL, SW_SHOWNORMAL);
        }
    } else if (success > 0) {
        if (MessageBoxW(hwnd, L"분할 완료! 폴더를 열까요?", L"완료", MB_YESNO) == IDYES) {
            ShellExecuteW(NULL, L"open", s_split_out_path, NULL, NULL, SW_SHOWNORMAL);
        }
    }
}

/* ==================== Merge Tab ==================== */
static void create_merge_tab(HWND hwnd, HINSTANCE hinst)
{
    int y = dpi(38);
    int lm = dpi(10);
    int cw = dpi(684);
    HWND h;

    #define ADD_MERGE_CTRL(x) s_merge_ctrls[s_merge_ctrl_count++] = (x)

    h = CreateWindowW(L"STATIC", L"PDF 병합", WS_CHILD,
        lm, y, dpi(200), dpi(22), hwnd, NULL, hinst, NULL);
    set_control_font(h, s_hfont_title); ADD_MERGE_CTRL(h);
    y += dpi(28);

    h = CreateWindowW(L"STATIC", L"병합할 파일 목록 (위에서 아래 순서로 병합됩니다)", WS_CHILD,
        lm, y, dpi(400), dpi(20), hwnd, NULL, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_MERGE_CTRL(h);
    y += dpi(25);

    s_hwnd_merge_list = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
        WS_CHILD | WS_VSCROLL | LBS_NOINTEGRALHEIGHT | LBS_NOTIFY,
        lm, y, cw - dpi(90), dpi(250), hwnd, (HMENU)ID_MERGE_LIST, hinst, NULL);
    set_control_font(s_hwnd_merge_list, s_hfont_ui); ADD_MERGE_CTRL(s_hwnd_merge_list);

    h = CreateWindowW(L"BUTTON", L"추가", WS_CHILD | BS_PUSHBUTTON,
        lm + cw - dpi(80), y, dpi(80), dpi(28), hwnd, (HMENU)ID_MERGE_BTN_ADD, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_MERGE_CTRL(h);

    h = CreateWindowW(L"BUTTON", L"삭제", WS_CHILD | BS_PUSHBUTTON,
        lm + cw - dpi(80), y + dpi(35), dpi(80), dpi(28), hwnd, (HMENU)ID_MERGE_BTN_DEL, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_MERGE_CTRL(h);

    h = CreateWindowW(L"BUTTON", L"위로", WS_CHILD | BS_PUSHBUTTON,
        lm + cw - dpi(80), y + dpi(80), dpi(80), dpi(28), hwnd, (HMENU)ID_MERGE_BTN_UP, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_MERGE_CTRL(h);

    h = CreateWindowW(L"BUTTON", L"아래로", WS_CHILD | BS_PUSHBUTTON,
        lm + cw - dpi(80), y + dpi(115), dpi(80), dpi(28), hwnd, (HMENU)ID_MERGE_BTN_DOWN, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_MERGE_CTRL(h);
    y += dpi(265);

    h = CreateWindowW(L"STATIC", L"출력 파일", WS_CHILD,
        lm, y + dpi(3), dpi(70), dpi(20), hwnd, NULL, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_MERGE_CTRL(h);

    s_hwnd_merge_out_path = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
        WS_CHILD | ES_AUTOHSCROLL,
        lm + dpi(75), y, cw - dpi(175), dpi(24), hwnd, (HMENU)ID_MERGE_OUT_PATH, hinst, NULL);
    set_control_font(s_hwnd_merge_out_path, s_hfont_ui);
    ADD_MERGE_CTRL(s_hwnd_merge_out_path);

    h = CreateWindowW(L"BUTTON", L"저장 위치...", WS_CHILD | BS_PUSHBUTTON,
        lm + cw - dpi(95), y - dpi(1), dpi(95), dpi(26), hwnd, (HMENU)ID_MERGE_BTN_OUT, hinst, NULL);
    set_control_font(h, s_hfont_ui); ADD_MERGE_CTRL(h);
    y += dpi(35);

    /* Progress bar */
    s_hwnd_merge_progress = CreateWindowExW(0, PROGRESS_CLASSW, L"",
        WS_CHILD | PBS_SMOOTH,
        lm, y, cw, dpi(20), hwnd, (HMENU)ID_MERGE_PROGRESS, hinst, NULL);
    ADD_MERGE_CTRL(s_hwnd_merge_progress);
    y += dpi(30);

    s_hwnd_merge_btn_run = CreateWindowW(L"BUTTON", L"PDF 병합 실행", WS_CHILD | BS_PUSHBUTTON,
        lm, y, cw, dpi(36), hwnd, (HMENU)ID_MERGE_BTN_RUN, hinst, NULL);
    set_control_font(s_hwnd_merge_btn_run, s_hfont_title); ADD_MERGE_CTRL(s_hwnd_merge_btn_run);

    #undef ADD_MERGE_CTRL
}

static void merge_add_file(const WCHAR* path)
{
    WCHAR msg[128];
    if (s_merge_file_count >= MAX_MERGE_FILES) {
        swprintf_s(msg, 128, L"최대 %d개 파일까지 추가할 수 있습니다.", MAX_MERGE_FILES);
        MessageBoxW(s_hwnd_main, msg, L"파일 추가 제한", MB_OK | MB_ICONWARNING);
        return;
    }
    wcscpy_s(s_merge_files[s_merge_file_count], MAX_PATH, path);
    s_merge_file_count++;
    merge_refresh_list();
}

static void merge_add_files(HWND hwnd)
{
    OPENFILENAMEW ofn;
    WCHAR file_buf[4096];
    WCHAR* p;
    WCHAR dir[MAX_PATH];
    WCHAR full_path[MAX_PATH];

    memset(file_buf, 0, sizeof(file_buf));
    memset(&ofn, 0, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFile = file_buf;
    ofn.nMaxFile = 4096;
    ofn.lpstrFilter = L"PDF 파일 (*.pdf)\0*.pdf\0";
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT | OFN_EXPLORER;

    if (GetOpenFileNameW(&ofn)) {
        wcscpy_s(dir, MAX_PATH, file_buf);
        p = file_buf + wcslen(file_buf) + 1;

        if (*p == L'\0') {
            /* Single file */
            merge_add_file(file_buf);
        } else {
            /* Multiple files */
            while (*p && s_merge_file_count < MAX_MERGE_FILES) {
                swprintf_s(full_path, MAX_PATH, L"%s\\%s", dir, p);
                merge_add_file(full_path);
                p += wcslen(p) + 1;
            }
        }
    }
}

static void merge_del_file(void)
{
    int sel = (int)SendMessageW(s_hwnd_merge_list, LB_GETCURSEL, 0, 0);
    int i;
    if (sel == LB_ERR) return;
    for (i = sel; i < s_merge_file_count - 1; i++) {
        wcscpy_s(s_merge_files[i], MAX_PATH, s_merge_files[i + 1]);
    }
    s_merge_file_count--;
    merge_refresh_list();

    /* Select next item (or last item if deleted the last one) */
    if (s_merge_file_count > 0) {
        if (sel >= s_merge_file_count) {
            sel = s_merge_file_count - 1;
        }
        SendMessageW(s_hwnd_merge_list, LB_SETCURSEL, sel, 0);
    }
}

static void merge_move_up(void)
{
    int sel = (int)SendMessageW(s_hwnd_merge_list, LB_GETCURSEL, 0, 0);
    WCHAR temp[MAX_PATH];
    if (sel <= 0) return;
    wcscpy_s(temp, MAX_PATH, s_merge_files[sel - 1]);
    wcscpy_s(s_merge_files[sel - 1], MAX_PATH, s_merge_files[sel]);
    wcscpy_s(s_merge_files[sel], MAX_PATH, temp);
    merge_refresh_list();
    SendMessageW(s_hwnd_merge_list, LB_SETCURSEL, sel - 1, 0);
}

static void merge_move_down(void)
{
    int sel = (int)SendMessageW(s_hwnd_merge_list, LB_GETCURSEL, 0, 0);
    WCHAR temp[MAX_PATH];
    if (sel < 0 || sel >= s_merge_file_count - 1) return;
    wcscpy_s(temp, MAX_PATH, s_merge_files[sel + 1]);
    wcscpy_s(s_merge_files[sel + 1], MAX_PATH, s_merge_files[sel]);
    wcscpy_s(s_merge_files[sel], MAX_PATH, temp);
    merge_refresh_list();
    SendMessageW(s_hwnd_merge_list, LB_SETCURSEL, sel + 1, 0);
}

static void merge_select_output(HWND hwnd)
{
    OPENFILENAMEW ofn;
    WCHAR file_path[MAX_PATH];
    HWND hwnd_out;

    wcscpy_s(file_path, MAX_PATH, L"merged.pdf");

    memset(&ofn, 0, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFilter = L"PDF 파일\0*.pdf\0모든 파일\0*.*\0";
    ofn.lpstrFile = file_path;
    ofn.nMaxFile = MAX_PATH;
    ofn.lpstrDefExt = L"pdf";
    ofn.lpstrTitle = L"병합 파일 저장";
    ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST;

    if (GetSaveFileNameW(&ofn)) {
        wcscpy_s(s_merge_out_path, MAX_PATH, file_path);
        hwnd_out = GetDlgItem(s_hwnd_main, ID_MERGE_OUT_PATH);
        if (hwnd_out) {
            SetWindowTextW(hwnd_out, s_merge_out_path);
        }
    }
}

static void merge_refresh_list(void)
{
    int i;
    WCHAR* filename;
    WCHAR msg[64];

    SendMessageW(s_hwnd_merge_list, LB_RESETCONTENT, 0, 0);
    for (i = 0; i < s_merge_file_count; i++) {
        filename = wcsrchr(s_merge_files[i], L'\\');
        filename = filename ? filename + 1 : s_merge_files[i];
        SendMessageW(s_hwnd_merge_list, LB_ADDSTRING, 0, (LPARAM)filename);
    }

    swprintf_s(msg, 64, L"%d개 파일 추가됨", s_merge_file_count);
    update_status(msg);
}

/* Progress callback for merge */
static void merge_progress_callback(int current, int total, void* user_data)
{
    WCHAR msg[64];
    (void)user_data;

    /* Update progress bar */
    SendMessageW(s_hwnd_merge_progress, PBM_SETRANGE32, 0, total);
    SendMessageW(s_hwnd_merge_progress, PBM_SETPOS, current, 0);

    /* Update status text */
    swprintf_s(msg, 64, L"병합 중... (%d/%d)", current, total);
    update_status(msg);

    /* Process messages to keep UI responsive */
    MSG winmsg;
    while (PeekMessage(&winmsg, NULL, 0, 0, PM_REMOVE)) {
        TranslateMessage(&winmsg);
        DispatchMessage(&winmsg);
    }
}

/* 출력 경로 검증 헬퍼 함수 */
static int validate_output_path(HWND hwnd, const WCHAR* path)
{
    WCHAR parent_dir[MAX_PATH];
    WCHAR* last_slash;
    WCHAR* ext;
    WCHAR* filename;
    const WCHAR* invalid_chars = L"<>:\"|?*";
    DWORD attr;

    /* 경로가 비어있는지 확인 */
    if (wcslen(path) == 0) {
        MessageBoxW(hwnd, L"출력 파일 경로를 입력해주세요.", L"경로 오류", MB_OK | MB_ICONWARNING);
        return 0;
    }

    /* .pdf 확장자 확인 */
    ext = wcsrchr(path, L'.');
    if (!ext || _wcsicmp(ext, L".pdf") != 0) {
        MessageBoxW(hwnd, L"출력 파일 이름은 .pdf 확장자로 끝나야 합니다.\n\n예: output.pdf", L"경로 오류", MB_OK | MB_ICONWARNING);
        return 0;
    }

    /* 파일명에 금지 문자 확인 */
    last_slash = wcsrchr(path, L'\\');
    filename = last_slash ? last_slash + 1 : (WCHAR*)path;
    if (wcspbrk(filename, invalid_chars) != NULL) {
        MessageBoxW(hwnd, L"파일 이름에 사용할 수 없는 문자가 포함되어 있습니다.\n\n사용 불가: < > : \" | ? *", L"경로 오류", MB_OK | MB_ICONWARNING);
        return 0;
    }

    /* 부모 폴더 존재 여부 확인 */
    wcscpy_s(parent_dir, MAX_PATH, path);
    last_slash = wcsrchr(parent_dir, L'\\');
    if (last_slash) {
        *last_slash = L'\0';
        attr = GetFileAttributesW(parent_dir);
        if (attr == INVALID_FILE_ATTRIBUTES || !(attr & FILE_ATTRIBUTE_DIRECTORY)) {
            WCHAR msg[512];
            swprintf_s(msg, 512, L"저장 경로가 올바르지 않습니다.\n폴더가 존재하는지 확인해주세요.\n\n경로: %s", parent_dir);
            MessageBoxW(hwnd, msg, L"경로 오류", MB_OK | MB_ICONWARNING);
            return 0;
        }
    }

    return 1;
}

static void merge_run(HWND hwnd)
{
    const WCHAR* paths[MAX_MERGE_FILES];
    WCHAR msg[512];
    int i;
    pdf_error_t error = PDF_OK;
    int failed_index = -1;
    WCHAR* failed_filename;

    if (s_merge_file_count < 2) {
        MessageBoxW(hwnd, L"2개 이상의 PDF 파일을 추가하세요.", L"오류", MB_OK | MB_ICONERROR);
        return;
    }

    /* Get output path from edit control (user may have edited it) */
    GetWindowTextW(s_hwnd_merge_out_path, s_merge_out_path, MAX_PATH);

    /* 출력 경로 검증 */
    if (!validate_output_path(hwnd, s_merge_out_path)) {
        SetFocus(s_hwnd_merge_out_path);
        return;
    }

    for (i = 0; i < s_merge_file_count; i++) {
        paths[i] = s_merge_files[i];
    }

    EnableWindow(s_hwnd_merge_btn_run, FALSE);

    /* Show and reset progress bar */
    ShowWindow(s_hwnd_merge_progress, SW_SHOW);
    SendMessageW(s_hwnd_merge_progress, PBM_SETPOS, 0, 0);
    update_status(L"병합 시작...");

    if (pdf_merge(paths, s_merge_file_count, s_merge_out_path, merge_progress_callback, NULL, &error, &failed_index)) {
        /* Hide progress bar on success */
        ShowWindow(s_hwnd_merge_progress, SW_HIDE);
        swprintf_s(msg, 512, L"병합 완료: %d개 파일", s_merge_file_count);
        update_status(msg);
        if (MessageBoxW(hwnd, L"병합 완료! 폴더를 열까요?", L"완료", MB_YESNO) == IDYES) {
            /* Open folder and select the merged file */
            WCHAR cmd[MAX_PATH + 16];
            swprintf_s(cmd, MAX_PATH + 16, L"/select,\"%s\"", s_merge_out_path);
            ShellExecuteW(NULL, L"open", L"explorer.exe", cmd, NULL, SW_SHOWNORMAL);
        }
    } else {
        ShowWindow(s_hwnd_merge_progress, SW_HIDE);
        update_status(L"병합 실패");

        /* 구체적인 오류 메시지 생성 */
        if (failed_index >= 0 && failed_index < s_merge_file_count) {
            failed_filename = wcsrchr(s_merge_files[failed_index], L'\\');
            failed_filename = failed_filename ? failed_filename + 1 : s_merge_files[failed_index];
            swprintf_s(msg, 512, L"병합에 실패했습니다.\n\n문제 파일: %s\n\n%s",
                       failed_filename, pdf_error_message(error));
        } else {
            swprintf_s(msg, 512, L"병합에 실패했습니다.\n\n%s", pdf_error_message(error));
        }
        MessageBoxW(hwnd, msg, L"병합 오류", MB_OK | MB_ICONERROR);
    }

    EnableWindow(s_hwnd_merge_btn_run, TRUE);
}
