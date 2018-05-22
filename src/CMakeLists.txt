cmake_minimum_required(VERSION 3.5)
project(OpenXLSX.Library VERSION 0.5)

# C/C++ language version
set(CMAKE_C_STANDARD 99)
set(CMAKE_C_STANDARD_REQUIRED ON)
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Project metadata
set(PROJECT_VENDOR "Kenneth Troldal Balslev")
set(PROJECT_CONTACT "kenneth.balslev@gmail.com")
set(PROJECT_URL "https://github.com/troldal/OpenXLSX")
set(PROJECT_DESCRIPTION "A library for reading, writing and modifying Excel spreadsheets for C++17")

# Load source and header files
file(GLOB WORKBOOK_HEADERS XLWorkbook/*.h)
file(GLOB WORKBOOK_SOURCES XLWorkbook/*.cpp)
file(GLOB SHEET_HEADERS XLSheet/*.h)
file(GLOB SHEET_SOURCES XLSheet/*.cpp)
file(GLOB CELL_HEADERS XLCell/*.h)
file(GLOB CELL_SOURCES XLCell/*.cpp)
file(GLOB XML_HEADERS Utilities/XML/*.h)
file(GLOB XML_SOURCES Utilities/XML/*.cpp)
file(GLOB ZIP_HEADERS Utilities/Zip/*.h)
file(GLOB ZIP_SOURCES Utilities/Zip/*.c)
file(GLOB ZIP_SOURCES Utilities/Zip/*.cpp)

set(OPENXLSX_HEADERS OpenXLSX.h Utilities/Zip/miniz.h
        ${WORKBOOK_HEADERS}
        ${SHEET_HEADERS}
        ${CELL_HEADERS}
        ${XML_HEADERS}
        ${ZIP_HEADERS} Utilities/Zip/miniz-cpp.h)

set(OPENXLSX_SOURCES Utilities/Zip/miniz.c
        ${WORKBOOK_SOURCES}
        ${SHEET_SOURCES}
        ${CELL_SOURCES}
        ${XML_SOURCES}
        ${ZIP_SOURCES})

# Build library
add_library(OpenXLSX STATIC ${OPENXLSX_HEADERS} ${OPENXLSX_SOURCES})
target_include_directories(OpenXLSX PUBLIC "${LIBZIP_INCLUDE_DIR}")
target_link_libraries(OpenXLSX PUBLIC ${LIBZIP_LIBRARY})

# Install directories
install (TARGETS OpenXLSX DESTINATION lib)
install (FILES ${WORKBOOK_HEADERS} ${SHEET_HEADERS} ${ZIP_HEADERS} DESTINATION include/OpenXLSX)
install (FILES ${CELL_SOURCES} DESTINATION include/OpenXLSX/Zip)

