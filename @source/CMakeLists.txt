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
set(WORKBOOK_HEADERS
        XLWorkbook/XLAbstractXMLFile.h
        XLWorkbook/XLAppProperties.h
        XLWorkbook/XLColor.h
        XLWorkbook/XLContentTypes.h
        XLWorkbook/XLCoreProperties.h
        XLWorkbook/XLDocument.h
        XLWorkbook/XLException.h
        XLWorkbook/XLFont.h
        XLWorkbook/XLRelationships.h
        XLWorkbook/XLSharedStrings.h
        XLWorkbook/XLSpreadsheetElement.h
        XLWorkbook/XLStyles.h
        XLWorkbook/XLTokenizer.h
        XLWorkbook/XLWorkbook.h)

set(WORKBOOK_SOURCES
        XLWorkbook/XLAbstractXMLFile.cpp
        XLWorkbook/XLAppProperties.cpp
        XLWorkbook/XLColor.cpp
        XLWorkbook/XLContentTypes.cpp
        XLWorkbook/XLCoreProperties.cpp
        XLWorkbook/XLDocument.cpp
        XLWorkbook/XLFont.cpp
        XLWorkbook/XLRelationships.cpp
        XLWorkbook/XLSharedStrings.cpp
        XLWorkbook/XLSpreadsheetElement.cpp
        XLWorkbook/XLStyles.cpp
        XLWorkbook/XLTokenizer.cpp
        XLWorkbook/XLWorkbook.cpp)

set(SHEET_HEADERS
        XLSheet/XLChartsheet.h
        XLSheet/XLColumn.h
        XLSheet/XLRow.h
        XLSheet/XLSheet.h
        XLSheet/XLTemplate.h
        XLSheet/XLWorksheet.h)

set(SHEET_SOURCES
        XLSheet/XLChartsheet.cpp
        XLSheet/XLColumn.cpp
        XLSheet/XLRow.cpp
        XLSheet/XLSheet.cpp
        XLSheet/XLWorksheet.cpp)

set(CELL_HEADERS
        XLCell/XLCell.h
        XLCell/XLCellFormats.h
        XLCell/XLCellFormula.h
        XLCell/XLCellIterator.h
        XLCell/XLCellRange.h
        XLCell/XLCellReference.h
        XLCell/XLCellType.h
        XLCell/XLCellValue.h
        XLCell/XLValue.h
        XLCell/XLValueBoolean.h
        XLCell/XLValueEmpty.h
        XLCell/XLValueError.h
        XLCell/XLValueNumber.h
        XLCell/XLValueString.h)

set(CELL_Sources
        XLCell/XLCell.cpp
        XLCell/XLCellFormats.cpp
        XLCell/XLCellFormula.cpp
        XLCell/XLCellIterator.cpp
        XLCell/XLCellRange.cpp
        XLCell/XLCellReference.cpp
        XLCell/XLCellValue.cpp
        XLCell/XLValue.cpp
        XLCell/XLValueBoolean.cpp
        XLCell/XLValueEmpty.cpp
        XLCell/XLValueError.cpp
        XLCell/XLValueNumber.cpp
        XLCell/XLValueString.cpp)

set(OPENXLSX_HEADERS
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
add_library(OpenXLSX STATIC $<TARGET_OBJECTS:xml> $<TARGET_OBJECTS:zlibstatic> $<TARGET_OBJECTS:zip> $<TARGET_OBJECTS:zip++> ${OPENXLSX_SOURCES})
#target_link_libraries(OpenXLSX PRIVATE zip++ xml)

#=======================================================================================================================
# Install OpenXLSX Library
#=======================================================================================================================
install (TARGETS OpenXLSX DESTINATION lib)
install (FILES ${WORKBOOK_HEADERS} ${SHEET_HEADERS} DESTINATION include/OpenXLSX)

