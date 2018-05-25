cmake_minimum_required(VERSION 3.5)
project(OpenXLSX.Library VERSION 0.5)

add_subdirectory(@thirdparty/@zlib)
add_subdirectory(@thirdparty/@libzip)

# C/C++ language version
set(CMAKE_C_STANDARD 99)
set(CMAKE_C_STANDARD_REQUIRED ON)
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)


set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -O3")
set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -flto")
set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -fno-use-linker-plugin")
set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -fwhole-program")

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
file(GLOB XML_HEADERS @thirdparty/XML/*.h)
file(GLOB XML_SOURCES @thirdparty/XML/*.cpp)
file(GLOB ZIP_HEADERS @thirdparty/Zip/*.h)
file(GLOB ZIP_SOURCES @thirdparty/Zip/*.cpp)

set(OPENXLSX_HEADERS OpenXLSX.h
        ${WORKBOOK_HEADERS}
        ${SHEET_HEADERS}
        ${CELL_HEADERS}
        ${XML_HEADERS}
        ${ZIP_HEADERS})

set(OPENXLSX_SOURCES
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

