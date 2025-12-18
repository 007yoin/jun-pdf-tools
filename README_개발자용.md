# Jun PDF Tools - 개발자 가이드

## 프로젝트 구조

```
jun-pdf-tools/
├── src/
│   ├── main.c           # Win32 GUI (탭, 버튼, 리스트박스 등)
│   ├── pdf_tools.c      # PDF 처리 로직 (QPDF 라이브러리 사용)
│   └── pdf_tools.h      # PDF 함수 헤더
├── CMakeLists.txt       # CMake 빌드 설정
├── README.md            # 사용자용 문서
└── README_개발자용.md   # 개발자용 문서 (이 파일)
```

## 기술 스택

- **언어**: C99
- **GUI**: Win32 API (순수 Windows API, 프레임워크 없음)
- **PDF 처리**: QPDF 라이브러리 (C API)
- **빌드 시스템**: CMake + vcpkg
- **컴파일러**: MSVC (Visual Studio Build Tools)

## 빌드 환경 설정

### 1. 필수 도구 설치

1. **Visual Studio Build Tools 2022**
   - [다운로드](https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022)
   - "C++ 빌드 도구" 워크로드 선택

2. **vcpkg** (패키지 관리자)
   ```cmd
   git clone https://github.com/microsoft/vcpkg.git C:\vcpkg
   cd C:\vcpkg
   bootstrap-vcpkg.bat
   ```

3. **CMake**
   - vcpkg가 자동으로 CMake를 다운로드합니다
   - 경로: `C:\vcpkg\downloads\tools\cmake-<버전>\cmake-<버전>-windows-x86_64\bin\cmake.exe`
   - 또는 [CMake 공식 사이트](https://cmake.org/download/)에서 별도 설치 후 PATH에 추가

### 2. QPDF 라이브러리 설치 (Static)

DLL 없이 단일 exe 파일을 만들기 위해 static 버전을 설치합니다:

```cmd
vcpkg install qpdf:x64-windows-static
```

## 빌드 방법

### 1. 저장소 클론

```cmd
git clone https://github.com/007yoin/jun-pdf-tools.git
cd jun-pdf-tools
```

### 2. CMake 설정

vcpkg의 CMake 사용 (경로는 버전에 따라 다를 수 있음):
```cmd
"C:\vcpkg\downloads\tools\cmake-3.31.10-windows\cmake-3.31.10-windows-x86_64\bin\cmake.exe" -B build -DCMAKE_TOOLCHAIN_FILE=C:/vcpkg/scripts/buildsystems/vcpkg.cmake -DVCPKG_TARGET_TRIPLET=x64-windows-static
```

또는 CMake가 PATH에 있는 경우:
```cmd
cmake -B build -DCMAKE_TOOLCHAIN_FILE=C:/vcpkg/scripts/buildsystems/vcpkg.cmake -DVCPKG_TARGET_TRIPLET=x64-windows-static
```

### 3. 빌드 실행

```cmd
"C:\vcpkg\downloads\tools\cmake-3.31.10-windows\cmake-3.31.10-windows-x86_64\bin\cmake.exe" --build build --config Release
```

또는 CMake가 PATH에 있는 경우:
```cmd
cmake --build build --config Release
```

### 4. 출력 파일

```
build\Release\jun-pdf-tools.exe
```

단일 exe 파일 (약 1.7MB), DLL 불필요.

## 코드 구조 설명

### main.c

Win32 GUI 전체 구현:

- **윈도우 프로시저**: `WndProc()` - 메시지 처리
- **탭 관리**: `show_tab()` - 분할/병합 탭 전환
- **분할 탭**:
  - `create_split_controls()` - 컨트롤 생성
  - `split_load_pdf()` - PDF 로드
  - `split_add_chapter()` - 챕터 추가
  - `split_execute()` - 분할 실행
- **병합 탭**:
  - `create_merge_controls()` - 컨트롤 생성
  - `merge_add_files()` - 파일 추가
  - `merge_execute()` - 병합 실행
- **드래그 앤 드롭**: `WM_DROPFILES` 처리

### pdf_tools.c

QPDF 라이브러리를 사용한 PDF 처리:

- `pdf_get_page_count()` - 페이지 수 조회
- `pdf_split()` - PDF 분할 (특정 페이지 범위 추출)
- `pdf_merge()` - PDF 병합 (순차적 2개씩 병합으로 메모리 최적화)
- `pdf_merge_two()` - 2개 PDF 병합 (내부 함수)

**한글 경로 처리**: QPDF는 한글 경로를 직접 처리하지 못하므로, 임시 파일(ASCII 경로)로 복사 후 처리.

### pdf_tools.h

```c
int pdf_get_page_count(const WCHAR* pdf_path);
int pdf_split(const WCHAR* input_path, const WCHAR* output_path,
              int start_page, int end_page);
int pdf_merge(const WCHAR** input_paths, int input_count,
              const WCHAR* output_path,
              pdf_progress_cb progress_cb, void* user_data);
```

## CMakeLists.txt 주요 설정

```cmake
# Static 런타임 (DLL 의존성 제거)
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

# Windows GUI 앱 (콘솔 창 없음)
set(CMAKE_WIN32_EXECUTABLE ON)

# QPDF 링크
find_package(qpdf CONFIG REQUIRED)
target_link_libraries(${PROJECT_NAME} PRIVATE qpdf::libqpdf)

# Windows 시스템 라이브러리
target_link_libraries(${PROJECT_NAME} PRIVATE
    comdlg32 shell32 ole32 user32 gdi32
)
```

## 주의사항

- **인코딩**: 모든 소스 파일은 UTF-8, 컴파일러 옵션 `/utf-8` 사용
- **유니코드**: `UNICODE`, `_UNICODE` 정의됨 (Wide 문자열 사용)
- **메모리**: 대용량 PDF 병합 시 순차적 2개씩 병합하여 메모리 사용 최소화
