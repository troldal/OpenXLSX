cmake_minimum_required(VERSION 3.5)
project(OpenXLSX.Library VERSION 0.5)

#=======================================================================================================================
# Determine C++ compiler vendor
#=======================================================================================================================
if ("${CMAKE_CXX_COMPILER_ID}" MATCHES "Clang")
    set(IS_CLANG TRUE)
elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
    set(IS_GCC TRUE)
elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Intel")
    set(IS_INTEL TRUE)
elseif ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "MSVC")
    set(IS_MSVC TRUE)
endif()

add_subdirectory(@zip)
add_subdirectory(@xml)

#=======================================================================================================================
# Set C/C++ compiler version
#=======================================================================================================================
set(CMAKE_C_STANDARD 99)
set(CMAKE_C_STANDARD_REQUIRED ON)
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

#=======================================================================================================================
# Set compiler flags
#=======================================================================================================================
IF(CMAKE_BUILD_TYPE STREQUAL "Debug")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wall")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wextra")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -pedantic")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wno-unused-value")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wno-unused-parameter")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Weffc++")
ELSEIF(CMAKE_BUILD_TYPE STREQUAL "Release")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -O3")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -flto")
    IF (MINGW)
        set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -fno-use-linker-plugin")
        set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -fwhole-program")
    ENDIF(MINGW)
ENDIF(CMAKE_BUILD_TYPE STREQUAL "Debug")


#=======================================================================================================================
# Set project metadata
#=======================================================================================================================
set(PROJECT_VENDOR "Kenneth Troldal Balslev")
set(PROJECT_CONTACT "kenneth.balslev@gmail.com")
set(PROJECT_URL "https://github.com/troldal/OpenXLSX")
set(PROJECT_DESCRIPTION "A library for reading, writing and modifying Excel spreadsheets for C++17")

#=======================================================================================================================
# Find project files
#=======================================================================================================================
file(GLOB WORKBOOK_HEADERS XLWorkbook/*.h)
file(GLOB WORKBOOK_SOURCES XLWorkbook/*.cpp)
file(GLOB SHEET_HEADERS XLSheet/*.h)
file(GLOB SHEET_SOURCES XLSheet/*.cpp)
file(GLOB CELL_HEADERS XLCell/*.h)
file(GLOB CELL_SOURCES XLCell/*.cpp)

set(OPENXLSX_HEADERS OpenXLSX.h
        ${WORKBOOK_HEADERS}
        ${SHEET_HEADERS}
        ${CELL_HEADERS})

set(OPENXLSX_SOURCES
        ${WORKBOOK_SOURCES}
        ${SHEET_SOURCES}
        ${CELL_SOURCES})

#=======================================================================================================================
# Build OpenXLSX library
#=======================================================================================================================
add_library(OpenXLSX STATIC ${OPENXLSX_HEADERS} ${OPENXLSX_SOURCES})
target_link_libraries(OpenXLSX PRIVATE zip++ xml)

#=======================================================================================================================
# Install OpenXLSX Library
#=======================================================================================================================
install (TARGETS OpenXLSX DESTINATION lib)
install (FILES ${WORKBOOK_HEADERS} ${SHEET_HEADERS} DESTINATION include/OpenXLSX)

