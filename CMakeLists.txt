cmake_minimum_required(VERSION 3.5)
project(RapidXLSX)

set(CMAKE_MODULE_PATH ${CMAKE_MODULE_PATH} "${CMAKE_SOURCE_DIR}/CMake/Modules/")

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED TRUE)
set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wall -Wextra -Weffc++ -pedantic -Wno-unused-value -Wno-unused-parameter")

find_package(libzip REQUIRED)
find_package(Boost COMPONENTS system filesystem REQUIRED)
set(Boost_USE_STATIC_LIBS ON)
link_libraries(${LIBZIP_LIBRARY} ${Boost_LIBRARIES})

set(SOURCE_FILES
        src/libzip++.h src/libzip++.cpp
        src/XML/rapidxml_iterators.hpp
        src/XML/rapidxml_print.hpp
        src/XML/rapidxml_utils.hpp
        src/XML/rapidxml.hpp
        src/XML/XMLNode.cpp src/XML/XMLNode.h
        src/XML/XMLAttribute.cpp src/XML/XMLAttribute.h
        src/XML/XMLDocument.cpp src/XML/XMLDocument.h
        src/XLDocument.cpp src/XLDocument.h
        src/XLRelationships.cpp src/XLRelationships.h
        src/XLAbstractXMLFile.cpp src/XLAbstractXMLFile.h
        src/XLContentTypes.cpp src/XLContentTypes.h
        src/XLDocumentAppProperties.cpp src/XLDocumentAppProperties.h
        src/XLDocumentCoreProperties.cpp src/XLDocumentCoreProperties.h
        src/XLSharedStrings.cpp src/XLSharedStrings.h
        src/XLStyles.cpp src/XLStyles.h
        src/XLWorkbook.cpp src/XLWorkbook.h
        src/XLAbstractSheet.cpp src/XLAbstractSheet.h
        src/XLCell.cpp src/XLCell.h
        src/XLWorksheet.cpp src/XLWorksheet.h
        src/XLCellReference.cpp src/XLCellReference.h
        src/XLCellRange.cpp src/XLCellRange.h
        src/XLArchive.cpp src/XLArchive.h
        src/XLCellIterator.cpp src/XLCellIterator.h
        src/XLChartsheet.cpp src/XLChartsheet.h
        src/XLTemplate.h
        src/XLRow.cpp src/XLRow.h
        src/XLColumn.cpp src/XLColumn.h
        src/XLColor.cpp src/XLColor.h
        src/XLFont.cpp src/XLFont.h
        src/XLCellFormats.cpp src/XLCellFormats.h
        src/XLException.h
        src/XLCellValue.cpp src/XLCellValue.h
        src/XLValueString.cpp src/XLValueString.h
        src/XLSpreadsheetElement.cpp src/XLSpreadsheetElement.h
        src/XLCellFormula.cpp src/XLCellFormula.h
        src/XLCellType.h
        src/XLValueNumber.cpp src/XLValueNumber.h
        src/XLValue.cpp src/XLValue.h
        src/XLValueBoolean.cpp src/XLValueBoolean.h
        src/XLValueEmpty.cpp src/XLValueEmpty.h
        src/XLValueError.cpp src/XLValueError.h)

add_library(RapidXLSX ${SOURCE_FILES})
