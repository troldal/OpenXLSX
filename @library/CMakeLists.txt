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
    #set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wall")
    #set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wextra")
    #set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -pedantic")
    #set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wno-unused-value")
    #set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wno-unused-parameter")
    #set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Weffc++")
    #set(CMAKE_VERBOSE_MAKEFILE ON)
ELSEIF(CMAKE_BUILD_TYPE STREQUAL "Release")
    #set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -O3")
    #set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -flto")
    IF (MINGW)
        #set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -fno-use-linker-plugin")
        #set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -fwhole-program")
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
set(OPENXLSX_HEADERS
        headers/XLAbstractXMLFile.h
        headers/XLAppProperties.h
        headers/XLColor.h
        headers/XLContentTypes.h
        headers/XLCoreProperties.h
        headers/XLDocument.h
        headers/XLException.h
        headers/XLFont.h
        headers/XLRelationships.h
        headers/XLSharedStrings.h
        headers/XLSpreadsheetElement.h
        headers/XLStyles.h
        headers/XLTokenizer.h
        headers/XLWorkbook.h
        headers/XLChartsheet.h
        headers/XLColumn.h
        headers/XLRow.h
        headers/XLSheet.h
        headers/XLTemplate.h
        headers/XLWorksheet.h
        headers/XLCell.h
        headers/XLCellFormats.h
        headers/XLCellFormula.h
        headers/XLCellIterator.h
        headers/XLCellRange.h
        headers/XLCellReference.h
        headers/XLCellType.h
        headers/XLCellValue.h)

set(OPENXLSX_SOURCES
        sources/XLAbstractXMLFile.cpp
        sources/XLAppProperties.cpp
        sources/XLColor.cpp
        sources/XLContentTypes.cpp
        sources/XLCoreProperties.cpp
        sources/XLDocument.cpp
        sources/XLFont.cpp
        sources/XLRelationships.cpp
        sources/XLSharedStrings.cpp
        sources/XLSpreadsheetElement.cpp
        sources/XLStyles.cpp
        sources/XLTokenizer.cpp
        sources/XLWorkbook.cpp
        sources/XLChartsheet.cpp
        sources/XLColumn.cpp
        sources/XLRow.cpp
        sources/XLSheet.cpp
        sources/XLWorksheet.cpp
        sources/XLCell.cpp
        sources/XLCellFormats.cpp
        sources/XLCellFormula.cpp
        sources/XLCellIterator.cpp
        sources/XLCellRange.cpp
        sources/XLCellReference.cpp
        sources/XLCellValue.cpp)

#=======================================================================================================================
# Build OpenXLSX library
#=======================================================================================================================
add_library(OpenXLSXDepend STATIC $<TARGET_OBJECTS:xml> $<TARGET_OBJECTS:zlibstatic> $<TARGET_OBJECTS:zip> $<TARGET_OBJECTS:zip++>)
add_library(OpenXLSX STATIC ${OPENXLSX_HEADERS} ${OPENXLSX_SOURCES})
target_link_libraries(OpenXLSX PRIVATE OpenXLSXDepend)

#=======================================================================================================================
# Install OpenXLSX Library
#=======================================================================================================================
install (TARGETS OpenXLSX DESTINATION lib)
install (FILES ${WORKBOOK_HEADERS} ${SHEET_HEADERS} DESTINATION include/OpenXLSX)

