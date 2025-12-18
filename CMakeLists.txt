cmake_minimum_required(VERSION 3.16)
project(jun-pdf-tools C)

set(CMAKE_C_STANDARD 99)
set(CMAKE_C_STANDARD_REQUIRED ON)

# Static runtime for single exe (no DLL dependencies)
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")

# Windows GUI
if(WIN32)
    set(CMAKE_WIN32_EXECUTABLE ON)
endif()

# QPDF
find_package(qpdf CONFIG REQUIRED)

# Source files
set(SOURCES
    src/main.c
    src/pdf_tools.c
)

set(HEADERS
    src/pdf_tools.h
)

# Executable
add_executable(${PROJECT_NAME} WIN32 ${SOURCES} ${HEADERS})

# Link QPDF
target_link_libraries(${PROJECT_NAME} PRIVATE qpdf::libqpdf)

# Windows libraries
if(WIN32)
    target_link_libraries(${PROJECT_NAME} PRIVATE
        comdlg32
        shell32
        ole32
        user32
        gdi32
    )
endif()

# Unicode
target_compile_definitions(${PROJECT_NAME} PRIVATE UNICODE _UNICODE)

# Compiler options
if(MSVC)
    target_compile_options(${PROJECT_NAME} PRIVATE /W4 /utf-8)
else()
    target_compile_options(${PROJECT_NAME} PRIVATE -Wall -Wextra)
endif()

