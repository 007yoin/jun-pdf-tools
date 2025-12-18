/*
 * pdf_tools.c - PDF processing using QPDF (Memory-optimized sequential merge)
 */

#include "pdf_tools.h"
#include <qpdf/qpdf-c.h>
#include <stdlib.h>
#include <stdio.h>
#include <string.h>

static FILE* g_log = NULL;
static void log_msg(const char* msg) {
    if (!g_log) g_log = _wfopen(L"C:\\Users\\mm\\Desktop\\pdf_debug.log", L"w");
    if (g_log) { fprintf(g_log, "%s\n", msg); fflush(g_log); }
}

/* 오류 코드를 사용자 친화적 메시지로 변환 */
const WCHAR* pdf_error_message(pdf_error_t error)
{
    switch (error) {
        case PDF_OK:
            return L"성공";
        case PDF_ERR_FILE_NOT_FOUND:
            return L"파일을 찾을 수 없습니다.";
        case PDF_ERR_ACCESS_DENIED:
            return L"파일에 접근할 수 없습니다.\n다른 프로그램에서 사용 중인지 확인하세요.";
        case PDF_ERR_INVALID_PDF:
            return L"PDF 파일이 손상되었거나 올바른 형식이 아닙니다.";
        case PDF_ERR_PASSWORD_PROTECTED:
            return L"암호로 보호된 PDF는 지원하지 않습니다.";
        case PDF_ERR_PAGE_OUT_OF_RANGE:
            return L"페이지 범위가 올바르지 않습니다.";
        case PDF_ERR_WRITE_FAILED:
            return L"파일을 저장할 수 없습니다.\n디스크 공간이 부족하거나 쓰기 권한이 없을 수 있습니다.";
        case PDF_ERR_MEMORY:
            return L"메모리가 부족합니다.\n파일 개수를 줄여서 다시 시도해주세요.";
        case PDF_ERR_TEMP_FILE:
            return L"임시 파일을 생성할 수 없습니다.\n디스크 공간을 확인해주세요.";
        case PDF_ERR_UNKNOWN:
        default:
            return L"알 수 없는 오류가 발생했습니다.";
    }
}

/* 에러 코드 설정 헬퍼 매크로 */
#define SET_ERROR(err_ptr, code) do { if (err_ptr) *(err_ptr) = (code); } while(0)

/* Copy file using Windows API (handles Korean paths) */
static int copy_file_w(const WCHAR* src, const WCHAR* dst)
{
    return CopyFileW(src, dst, FALSE) ? 1 : 0;
}

/* Generate temp file path (ASCII-safe for QPDF) */
static int get_temp_file(WCHAR* out_path, const WCHAR* prefix)
{
    WCHAR temp_dir[MAX_PATH];
    if (GetTempPathW(MAX_PATH, temp_dir) == 0) return 0;
    if (GetTempFileNameW(temp_dir, prefix, 0, out_path) == 0) return 0;
    return 1;
}

/* Convert wide string to narrow string (for temp paths which are ASCII) */
static int wchar_to_utf8(const WCHAR* wstr, char* str, int len)
{
    return WideCharToMultiByte(CP_UTF8, 0, wstr, -1, str, len, NULL, NULL);
}

int pdf_get_page_count(const WCHAR* pdf_path, pdf_error_t* error)
{
    WCHAR temp_path[MAX_PATH];
    char temp_path_a[MAX_PATH];
    qpdf_data qpdf;
    int page_count;
    DWORD file_attr;
    int qpdf_status;

    SET_ERROR(error, PDF_OK);

    /* 파일 존재 여부 확인 */
    file_attr = GetFileAttributesW(pdf_path);
    if (file_attr == INVALID_FILE_ATTRIBUTES) {
        DWORD err = GetLastError();
        if (err == ERROR_FILE_NOT_FOUND || err == ERROR_PATH_NOT_FOUND) {
            SET_ERROR(error, PDF_ERR_FILE_NOT_FOUND);
        } else if (err == ERROR_ACCESS_DENIED) {
            SET_ERROR(error, PDF_ERR_ACCESS_DENIED);
        } else {
            SET_ERROR(error, PDF_ERR_UNKNOWN);
        }
        return -1;
    }

    /* Copy to temp file (ASCII path for QPDF) */
    if (!get_temp_file(temp_path, L"pdf")) {
        SET_ERROR(error, PDF_ERR_TEMP_FILE);
        return -1;
    }

    if (!copy_file_w(pdf_path, temp_path)) {
        DWORD err = GetLastError();
        DeleteFileW(temp_path);
        if (err == ERROR_ACCESS_DENIED || err == ERROR_SHARING_VIOLATION) {
            SET_ERROR(error, PDF_ERR_ACCESS_DENIED);
        } else {
            SET_ERROR(error, PDF_ERR_UNKNOWN);
        }
        return -1;
    }

    wchar_to_utf8(temp_path, temp_path_a, MAX_PATH);

    qpdf = qpdf_init();
    if (qpdf == NULL) {
        DeleteFileW(temp_path);
        SET_ERROR(error, PDF_ERR_MEMORY);
        return -1;
    }

    page_count = -1;

    /* Stream read from file (low memory) */
    qpdf_status = qpdf_read(qpdf, temp_path_a, NULL);
    if (qpdf_status < 2) {
        page_count = qpdf_get_num_pages(qpdf);
    } else {
        /* QPDF 오류 - PDF 형식 문제 또는 암호화 */
        const char* qpdf_err = qpdf_get_error_full_text(qpdf, qpdf_get_error(qpdf));
        if (qpdf_err && strstr(qpdf_err, "password")) {
            SET_ERROR(error, PDF_ERR_PASSWORD_PROTECTED);
        } else {
            SET_ERROR(error, PDF_ERR_INVALID_PDF);
        }
    }

    qpdf_cleanup(&qpdf);
    DeleteFileW(temp_path);

    return page_count;
}

int pdf_split(const WCHAR* input_path, const WCHAR* output_path, int start_page, int end_page, pdf_error_t* error)
{
    WCHAR temp_in[MAX_PATH];
    WCHAR temp_out[MAX_PATH];
    char temp_in_a[MAX_PATH];
    char temp_out_a[MAX_PATH];
    qpdf_data qpdf_in;
    qpdf_data qpdf_out;
    qpdf_oh page;
    int i;
    int result;
    int total_pages;

    SET_ERROR(error, PDF_OK);

    /* Create temp files */
    if (!get_temp_file(temp_in, L"pin") || !get_temp_file(temp_out, L"pou")) {
        SET_ERROR(error, PDF_ERR_TEMP_FILE);
        return 0;
    }

    /* Copy input to temp */
    if (!copy_file_w(input_path, temp_in)) {
        DWORD err = GetLastError();
        DeleteFileW(temp_in);
        DeleteFileW(temp_out);
        if (err == ERROR_FILE_NOT_FOUND || err == ERROR_PATH_NOT_FOUND) {
            SET_ERROR(error, PDF_ERR_FILE_NOT_FOUND);
        } else if (err == ERROR_ACCESS_DENIED || err == ERROR_SHARING_VIOLATION) {
            SET_ERROR(error, PDF_ERR_ACCESS_DENIED);
        } else {
            SET_ERROR(error, PDF_ERR_UNKNOWN);
        }
        return 0;
    }

    wchar_to_utf8(temp_in, temp_in_a, MAX_PATH);
    wchar_to_utf8(temp_out, temp_out_a, MAX_PATH);

    qpdf_in = qpdf_init();
    qpdf_out = qpdf_init();

    if (qpdf_in == NULL || qpdf_out == NULL) {
        if (qpdf_in) qpdf_cleanup(&qpdf_in);
        if (qpdf_out) qpdf_cleanup(&qpdf_out);
        DeleteFileW(temp_in);
        DeleteFileW(temp_out);
        SET_ERROR(error, PDF_ERR_MEMORY);
        return 0;
    }

    result = 0;

    /* Stream read from file */
    if (qpdf_read(qpdf_in, temp_in_a, NULL) < 2) {
        /* 페이지 범위 검증 */
        total_pages = qpdf_get_num_pages(qpdf_in);
        if (start_page < 1 || end_page > total_pages || start_page > end_page) {
            qpdf_cleanup(&qpdf_in);
            qpdf_cleanup(&qpdf_out);
            DeleteFileW(temp_in);
            DeleteFileW(temp_out);
            SET_ERROR(error, PDF_ERR_PAGE_OUT_OF_RANGE);
            return 0;
        }

        qpdf_empty_pdf(qpdf_out);

        /* Copy pages one by one (0-indexed) */
        for (i = start_page - 1; i < end_page; i++) {
            page = qpdf_get_page_n(qpdf_in, i);
            qpdf_add_page(qpdf_out, qpdf_in, page, QPDF_FALSE);
        }

        /* Stream write to file (low memory) */
        qpdf_init_write(qpdf_out, temp_out_a);
        qpdf_set_compress_streams(qpdf_out, QPDF_TRUE);
        qpdf_set_object_stream_mode(qpdf_out, qpdf_o_generate);

        if (qpdf_write(qpdf_out) < 2) {
            result = 1;
        } else {
            SET_ERROR(error, PDF_ERR_WRITE_FAILED);
        }
    } else {
        SET_ERROR(error, PDF_ERR_INVALID_PDF);
    }

    qpdf_cleanup(&qpdf_in);
    qpdf_cleanup(&qpdf_out);

    /* Copy result to final destination */
    if (result) {
        if (!copy_file_w(temp_out, output_path)) {
            DWORD err = GetLastError();
            result = 0;
            if (err == ERROR_ACCESS_DENIED) {
                SET_ERROR(error, PDF_ERR_ACCESS_DENIED);
            } else if (err == ERROR_DISK_FULL) {
                SET_ERROR(error, PDF_ERR_WRITE_FAILED);
            } else {
                SET_ERROR(error, PDF_ERR_WRITE_FAILED);
            }
        }
    }

    DeleteFileW(temp_in);
    DeleteFileW(temp_out);

    return result;
}

/*
 * pdf_merge_two - Merge exactly 2 PDF files (internal function)
 * This keeps only 2 PDFs in memory at a time for low memory usage
 * @param which_failed: 0=none, 1=first file, 2=second file, 3=output
 */
static int pdf_merge_two(const WCHAR* path1, const WCHAR* path2, const WCHAR* output_path,
                         pdf_error_t* error, int* which_failed)
{
    WCHAR temp1[MAX_PATH], temp2[MAX_PATH], temp_out[MAX_PATH];
    char temp1_a[MAX_PATH], temp2_a[MAX_PATH], temp_out_a[MAX_PATH];
    qpdf_data qpdf1 = NULL, qpdf2 = NULL, qpdf_out = NULL;
    int i, page_count, result = 0;
    char buf[128];
    pdf_error_t local_error = PDF_OK;

    if (which_failed) *which_failed = 0;

    log_msg("pdf_merge_two: start");

    /* Create temp files */
    if (!get_temp_file(temp1, L"pm1") || !get_temp_file(temp2, L"pm2") || !get_temp_file(temp_out, L"pmo")) {
        log_msg("ERROR: failed to create temp files");
        local_error = PDF_ERR_TEMP_FILE;
        goto cleanup;
    }

    /* Copy inputs to temp */
    if (!copy_file_w(path1, temp1)) {
        log_msg("ERROR: copy path1 failed");
        DWORD err = GetLastError();
        if (err == ERROR_FILE_NOT_FOUND || err == ERROR_PATH_NOT_FOUND) {
            local_error = PDF_ERR_FILE_NOT_FOUND;
        } else if (err == ERROR_ACCESS_DENIED || err == ERROR_SHARING_VIOLATION) {
            local_error = PDF_ERR_ACCESS_DENIED;
        } else {
            local_error = PDF_ERR_UNKNOWN;
        }
        if (which_failed) *which_failed = 1;
        goto cleanup;
    }
    if (!copy_file_w(path2, temp2)) {
        log_msg("ERROR: copy path2 failed");
        DWORD err = GetLastError();
        if (err == ERROR_FILE_NOT_FOUND || err == ERROR_PATH_NOT_FOUND) {
            local_error = PDF_ERR_FILE_NOT_FOUND;
        } else if (err == ERROR_ACCESS_DENIED || err == ERROR_SHARING_VIOLATION) {
            local_error = PDF_ERR_ACCESS_DENIED;
        } else {
            local_error = PDF_ERR_UNKNOWN;
        }
        if (which_failed) *which_failed = 2;
        goto cleanup;
    }

    wchar_to_utf8(temp1, temp1_a, MAX_PATH);
    wchar_to_utf8(temp2, temp2_a, MAX_PATH);
    wchar_to_utf8(temp_out, temp_out_a, MAX_PATH);

    /* Initialize QPDF objects */
    qpdf1 = qpdf_init();
    qpdf2 = qpdf_init();
    qpdf_out = qpdf_init();

    if (!qpdf1 || !qpdf2 || !qpdf_out) {
        log_msg("ERROR: qpdf_init failed");
        local_error = PDF_ERR_MEMORY;
        goto cleanup;
    }

    /* Read both PDFs */
    if (qpdf_read(qpdf1, temp1_a, NULL) >= 2) {
        log_msg("ERROR: qpdf_read path1 failed");
        local_error = PDF_ERR_INVALID_PDF;
        if (which_failed) *which_failed = 1;
        goto cleanup;
    }
    if (qpdf_read(qpdf2, temp2_a, NULL) >= 2) {
        log_msg("ERROR: qpdf_read path2 failed");
        local_error = PDF_ERR_INVALID_PDF;
        if (which_failed) *which_failed = 2;
        goto cleanup;
    }

    /* Create output and add pages */
    qpdf_empty_pdf(qpdf_out);

    /* Add pages from first PDF */
    page_count = qpdf_get_num_pages(qpdf1);
    sprintf(buf, "Adding %d pages from file 1", page_count);
    log_msg(buf);
    for (i = 0; i < page_count; i++) {
        qpdf_add_page(qpdf_out, qpdf1, qpdf_get_page_n(qpdf1, i), QPDF_FALSE);
    }

    /* Add pages from second PDF */
    page_count = qpdf_get_num_pages(qpdf2);
    sprintf(buf, "Adding %d pages from file 2", page_count);
    log_msg(buf);
    for (i = 0; i < page_count; i++) {
        qpdf_add_page(qpdf_out, qpdf2, qpdf_get_page_n(qpdf2, i), QPDF_FALSE);
    }

    /* Write output */
    qpdf_init_write(qpdf_out, temp_out_a);
    qpdf_set_static_ID(qpdf_out, QPDF_TRUE);

    if (qpdf_write(qpdf_out) < 2) {
        log_msg("qpdf_write OK");
        if (!copy_file_w(temp_out, output_path)) {
            local_error = PDF_ERR_WRITE_FAILED;
            if (which_failed) *which_failed = 3;
        } else {
            result = 1;
        }
    } else {
        log_msg("ERROR: qpdf_write failed");
        local_error = PDF_ERR_WRITE_FAILED;
        if (which_failed) *which_failed = 3;
    }

cleanup:
    if (qpdf1) qpdf_cleanup(&qpdf1);
    if (qpdf2) qpdf_cleanup(&qpdf2);
    if (qpdf_out) qpdf_cleanup(&qpdf_out);
    DeleteFileW(temp1);
    DeleteFileW(temp2);
    DeleteFileW(temp_out);

    SET_ERROR(error, local_error);

    sprintf(buf, "pdf_merge_two: end, result=%d", result);
    log_msg(buf);

    return result;
}

/*
 * pdf_merge - Merge multiple PDF files using sequential merge
 * Memory usage: Only 2 PDFs in memory at any time
 *
 * Algorithm:
 *   [A, B, C, D, E] ->
 *   A + B -> temp1
 *   temp1 + C -> temp2
 *   temp2 + D -> temp1
 *   temp1 + E -> output
 */
int pdf_merge(const WCHAR** input_paths, int input_count, const WCHAR* output_path,
              pdf_progress_cb progress_cb, void* user_data, pdf_error_t* error, int* failed_index)
{
    WCHAR temp1[MAX_PATH], temp2[MAX_PATH];
    WCHAR* current;
    WCHAR* next;
    int i, result = 0;
    int total_steps;
    char buf[128];
    int which_failed = 0;

    SET_ERROR(error, PDF_OK);
    if (failed_index) *failed_index = -1;

    log_msg("=== MERGE START (SEQUENTIAL MODE) ===");
    sprintf(buf, "input_count: %d", input_count);
    log_msg(buf);

    if (input_count <= 0) {
        log_msg("ERROR: invalid input_count");
        SET_ERROR(error, PDF_ERR_UNKNOWN);
        return 0;
    }

    /* Calculate total steps: (input_count - 1) merges */
    total_steps = (input_count > 1) ? (input_count - 1) : 1;

    /* Single file: just copy */
    if (input_count == 1) {
        log_msg("Single file, copying...");
        if (progress_cb) progress_cb(1, 1, user_data);
        if (!copy_file_w(input_paths[0], output_path)) {
            SET_ERROR(error, PDF_ERR_WRITE_FAILED);
            return 0;
        }
        return 1;
    }

    /* Two files: direct merge */
    if (input_count == 2) {
        log_msg("Two files, direct merge...");
        if (progress_cb) progress_cb(1, 1, user_data);
        result = pdf_merge_two(input_paths[0], input_paths[1], output_path, error, &which_failed);
        if (!result && which_failed > 0 && which_failed <= 2 && failed_index) {
            *failed_index = which_failed - 1;  /* Convert to 0-based index */
        }
        return result;
    }

    /* 3+ files: sequential merge */
    log_msg("Multiple files, sequential merge...");

    /* Create temp file paths */
    if (!get_temp_file(temp1, L"seq") || !get_temp_file(temp2, L"seq")) {
        log_msg("ERROR: failed to create temp files");
        SET_ERROR(error, PDF_ERR_TEMP_FILE);
        return 0;
    }

    /* First merge: input[0] + input[1] -> temp1 */
    sprintf(buf, "Step 1: merging files 0 and 1");
    log_msg(buf);
    if (progress_cb) progress_cb(1, total_steps, user_data);
    if (!pdf_merge_two(input_paths[0], input_paths[1], temp1, error, &which_failed)) {
        log_msg("ERROR: first merge failed");
        if (which_failed > 0 && which_failed <= 2 && failed_index) {
            *failed_index = which_failed - 1;  /* 0 or 1 */
        }
        goto fail;
    }

    /* Sequential merge: temp + input[i] -> next_temp */
    for (i = 2; i < input_count; i++) {
        sprintf(buf, "Step %d: merging with file %d", i, i);
        log_msg(buf);

        /* Report progress */
        if (progress_cb) progress_cb(i, total_steps, user_data);

        /* Alternate between temp1 and temp2 */
        current = (i % 2 == 0) ? temp1 : temp2;
        next = (i % 2 == 0) ? temp2 : temp1;

        /* Last file: output to final destination */
        if (i == input_count - 1) {
            if (!pdf_merge_two(current, input_paths[i], output_path, error, &which_failed)) {
                log_msg("ERROR: final merge failed");
                if (which_failed == 2 && failed_index) {
                    *failed_index = i;  /* The current input file */
                }
                goto fail;
            }
        } else {
            if (!pdf_merge_two(current, input_paths[i], next, error, &which_failed)) {
                log_msg("ERROR: intermediate merge failed");
                if (which_failed == 2 && failed_index) {
                    *failed_index = i;  /* The current input file */
                }
                goto fail;
            }
        }

        /* Delete the temp file we just used as input */
        DeleteFileW(current);
    }

    result = 1;
    if (progress_cb) progress_cb(total_steps, total_steps, user_data);
    log_msg("=== MERGE END (SUCCESS) ===");

    /* Cleanup remaining temp files */
    DeleteFileW(temp1);
    DeleteFileW(temp2);
    return result;

fail:
    log_msg("=== MERGE END (FAILED) ===");
    DeleteFileW(temp1);
    DeleteFileW(temp2);
    return 0;
}
