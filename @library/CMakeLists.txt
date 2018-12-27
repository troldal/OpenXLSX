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

add_library(OpenXLSX STATIC "")

add_subdirectory(@dependencies)

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
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLAbstractXMLFile.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLAppProperties.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLColor.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLContentTypes.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCoreProperties.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLDocument.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLException.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLFont.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLRelationships.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLSharedStrings.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLSpreadsheetElement.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLStyles.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLTokenizer.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLWorkbook.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLChartsheet.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLColumn.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLRow.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLSheet.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLTemplate.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLWorksheet.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCell.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellFormats.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellFormula.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellIterator.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellRange.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellReference.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellType.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLCellValue.h
        ${CMAKE_CURRENT_LIST_DIR}/headers/XLXml.h)

set(OPENXLSX_SOURCES
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLAbstractXMLFile.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLAppProperties.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLColor.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLContentTypes.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCoreProperties.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLDocument.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLFont.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLRelationships.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLSharedStrings.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLSpreadsheetElement.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLStyles.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLTokenizer.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLWorkbook.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLChartsheet.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLColumn.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLRow.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLSheet.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLWorksheet.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCell.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellFormats.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellFormula.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellIterator.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellRange.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellReference.cpp
        ${CMAKE_CURRENT_LIST_DIR}/sources/XLCellValue.cpp)

#=======================================================================================================================
# Build OpenXLSX library
#=======================================================================================================================
target_sources(OpenXLSX PUBLIC ${OPENXLSX_HEADERS} ${OPENXLSX_SOURCES})
target_link_libraries(OpenXLSX PRIVATE pugixml zip++)

#=======================================================================================================================
# Install OpenXLSX Library
#=======================================================================================================================
install (TARGETS OpenXLSX DESTINATION lib)
install (FILES ${WORKBOOK_HEADERS} ${SHEET_HEADERS} DESTINATION include/OpenXLSX)

