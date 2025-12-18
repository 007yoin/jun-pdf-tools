/*
 * pdf_tools.h
 * PDF split/merge using QPDF
 */

#ifndef PDF_TOOLS_H
#define PDF_TOOLS_H

#include <windows.h>

/*
 * PDF 작업 오류 코드
 */
typedef enum {
    PDF_OK = 0,                     /* 성공 */
    PDF_ERR_FILE_NOT_FOUND = -1,    /* 파일을 찾을 수 없음 */
    PDF_ERR_ACCESS_DENIED = -2,     /* 파일 접근 거부 */
    PDF_ERR_INVALID_PDF = -3,       /* 손상되었거나 잘못된 PDF */
    PDF_ERR_PASSWORD_PROTECTED = -4,/* 암호로 보호된 PDF */
    PDF_ERR_PAGE_OUT_OF_RANGE = -5, /* 페이지 범위 초과 */
    PDF_ERR_WRITE_FAILED = -6,      /* 파일 쓰기 실패 */
    PDF_ERR_MEMORY = -7,            /* 메모리 부족 */
    PDF_ERR_TEMP_FILE = -8,         /* 임시 파일 생성 실패 */
    PDF_ERR_UNKNOWN = -99           /* 알 수 없는 오류 */
} pdf_error_t;

/*
 * 오류 코드를 사용자 친화적 메시지로 변환
 * @param error 오류 코드
 * @return 오류 메시지 (정적 문자열)
 */
const WCHAR* pdf_error_message(pdf_error_t error);

/*
 * Progress callback type for merge operations.
 * @param current current step (1-based)
 * @param total total steps
 * @param user_data user-provided data
 */
typedef void (*pdf_progress_cb)(int current, int total, void* user_data);

/*
 * Get page count of a PDF file.
 *
 * @param pdf_path PDF file path
 * @param error 오류 코드 출력 (NULL 가능)
 * @return page count (-1 on error)
 */
int pdf_get_page_count(const WCHAR* pdf_path, pdf_error_t* error);

/*
 * Split pages from a PDF file.
 *
 * @param input_path source PDF path
 * @param output_path output PDF path
 * @param start_page start page (1-based)
 * @param end_page end page (inclusive)
 * @param error 오류 코드 출력 (NULL 가능)
 * @return 1 on success, 0 on failure
 */
int pdf_split(const WCHAR* input_path, const WCHAR* output_path, int start_page, int end_page, pdf_error_t* error);

/*
 * Merge multiple PDF files into one.
 *
 * @param input_paths array of input PDF paths
 * @param input_count number of input files
 * @param output_path output PDF path
 * @param progress_cb progress callback (can be NULL)
 * @param user_data user data for callback
 * @param error 오류 코드 출력 (NULL 가능)
 * @param failed_index 실패한 파일 인덱스 출력 (NULL 가능, 입력 파일 오류 시)
 * @return 1 on success, 0 on failure
 */
int pdf_merge(const WCHAR** input_paths, int input_count, const WCHAR* output_path,
              pdf_progress_cb progress_cb, void* user_data, pdf_error_t* error, int* failed_index);

#endif /* PDF_TOOLS_H */
