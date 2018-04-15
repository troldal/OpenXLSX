cmake_minimum_required(VERSION 3.5)
project(OpenXLSX)

set(CMAKE_MODULE_PATH ${CMAKE_MODULE_PATH} "${CMAKE_SOURCE_DIR}/CMake/Modules/")

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED TRUE)
IF(CMAKE_BUILD_TYPE STREQUAL "Debug")
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wall -Wextra -Weffc++ -pedantic -Wno-unused-value -Wno-unused-parameter")
ENDIF(CMAKE_BUILD_TYPE STREQUAL "Debug")
#set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -pg")
#SET(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} -pg")
#SET(CMAKE_SHARED_LINKER_FLAGS "${CMAKE_SHARED_LINKER_FLAGS} -pg")

find_package(libzip REQUIRED)
find_package(Boost COMPONENTS system filesystem REQUIRED)
set(Boost_USE_STATIC_LIBS ON)

set(SOURCE_FILES
        src/Zip/libzip++.h src/Zip/libzip++.cpp
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
        src/XLAppProperties.cpp src/XLAppProperties.h
        src/XLCoreProperties.cpp src/XLCoreProperties.h
        src/XLSharedStrings.cpp src/XLSharedStrings.h
        src/XLStyles.cpp src/XLStyles.h
        src/XLWorkbook.cpp src/XLWorkbook.h
        src/XLAbstractSheet.cpp src/XLAbstractSheet.h
        src/XLCell.cpp src/XLCell.h
        src/XLWorksheet.cpp src/XLWorksheet.h
        src/XLCellReference.cpp src/XLCellReference.h
        src/XLCellRange.cpp src/XLCellRange.h
        src/Zip/XLArchive.cpp src/Zip/XLArchive.h
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
        src/XLValueError.cpp src/XLValueError.h src/OpenXLSX.h)

add_library(OpenXLSX ${SOURCE_FILES})
target_include_directories(OpenXLSX PUBLIC "${LIBZIP_INCLUDE_DIR}")
target_link_libraries(OpenXLSX PUBLIC ${LIBZIP_LIBRARY} ${Boost_LIBRARIES})

install (TARGETS OpenXLSX DESTINATION lib)
install (FILES
        src/OpenXLSX.h
        src/XLDocument.h
        src/XLRelationships.h
        src/XLAbstractXMLFile.h
        src/XLContentTypes.h
        src/XLAppProperties.h
        src/XLCoreProperties.h
        src/XLSharedStrings.h
        src/XLStyles.h
        src/XLWorkbook.h
        src/XLAbstractSheet.h
        src/XLCell.h
        src/XLWorksheet.h
        src/XLCellReference.h
        src/XLCellRange.h
        src/XLCellIterator.h
        src/XLChartsheet.h
        src/XLTemplate.h
        src/XLRow.h
        src/XLColumn.h
        src/XLColor.h
        src/XLFont.h
        src/XLCellFormats.h
        src/XLException.h
        src/XLCellValue.h
        src/XLValueString.h
        src/XLSpreadsheetElement.h
        src/XLCellFormula.h
        src/XLCellType.h
        src/XLValueNumber.h
        src/XLValue.h
        src/XLValueBoolean.h
        src/XLValueEmpty.h
        src/XLValueError.h
        DESTINATION include/OpenXLSX)
install (FILES
        src/XML/rapidxml_iterators.hpp
        src/XML/rapidxml_print.hpp
        src/XML/rapidxml_utils.hpp
        src/XML/rapidxml.hpp
        src/XML/XMLNode.h
        src/XML/XMLAttribute.h
        src/XML/XMLDocument.h
        DESTINATION include/OpenXLSX/XML)
install (FILES
        src/Zip/libzip++.h
        src/Zip/XLArchive.h
        DESTINATION include/OpenXLSX/Zip)

