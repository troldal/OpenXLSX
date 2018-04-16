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
        Zip/libzip++.h Zip/libzip++.cpp
        XML/rapidxml_iterators.hpp
        XML/rapidxml_print.hpp
        XML/rapidxml_utils.hpp
        XML/rapidxml.hpp
        XML/XMLNode.cpp XML/XMLNode.h
        XML/XMLAttribute.cpp XML/XMLAttribute.h
        XML/XMLDocument.cpp XML/XMLDocument.h
        XLDocument.cpp XLDocument.h
        XLRelationships.cpp XLRelationships.h
        XLAbstractXMLFile.cpp XLAbstractXMLFile.h
        XLContentTypes.cpp XLContentTypes.h
        XLAppProperties.cpp XLAppProperties.h
        XLCoreProperties.cpp XLCoreProperties.h
        XLSharedStrings.cpp XLSharedStrings.h
        XLStyles.cpp XLStyles.h
        XLWorkbook.cpp XLWorkbook.h
        XLAbstractSheet.cpp XLAbstractSheet.h
        XLCell.cpp XLCell.h
        XLWorksheet.cpp XLWorksheet.h
        XLCellReference.cpp XLCellReference.h
        XLCellRange.cpp XLCellRange.h
        Zip/XLArchive.cpp Zip/XLArchive.h
        XLCellIterator.cpp XLCellIterator.h
        XLChartsheet.cpp XLChartsheet.h
        XLTemplate.h
        XLRow.cpp XLRow.h
        XLColumn.cpp XLColumn.h
        XLColor.cpp XLColor.h
        XLFont.cpp XLFont.h
        XLCellFormats.cpp XLCellFormats.h
        XLException.h
        XLCellValue.cpp XLCellValue.h
        XLValueString.cpp XLValueString.h
        XLSpreadsheetElement.cpp XLSpreadsheetElement.h
        XLCellFormula.cpp XLCellFormula.h
        XLCellType.h
        XLValueNumber.cpp XLValueNumber.h
        XLValue.cpp XLValue.h
        XLValueBoolean.cpp XLValueBoolean.h
        XLValueEmpty.cpp XLValueEmpty.h
        XLValueError.cpp XLValueError.h OpenXLSX.h)

add_library(OpenXLSX ${SOURCE_FILES})
target_include_directories(OpenXLSX PUBLIC "${LIBZIP_INCLUDE_DIR}")
target_link_libraries(OpenXLSX PUBLIC ${LIBZIP_LIBRARY} ${Boost_LIBRARIES})

install (TARGETS OpenXLSX DESTINATION lib)
install (FILES
        OpenXLSX.h
        XLDocument.h
        XLRelationships.h
        XLAbstractXMLFile.h
        XLContentTypes.h
        XLAppProperties.h
        XLCoreProperties.h
        XLSharedStrings.h
        XLStyles.h
        XLWorkbook.h
        XLAbstractSheet.h
        XLCell.h
        XLWorksheet.h
        XLCellReference.h
        XLCellRange.h
        XLCellIterator.h
        XLChartsheet.h
        XLTemplate.h
        XLRow.h
        XLColumn.h
        XLColor.h
        XLFont.h
        XLCellFormats.h
        XLException.h
        XLCellValue.h
        XLValueString.h
        XLSpreadsheetElement.h
        XLCellFormula.h
        XLCellType.h
        XLValueNumber.h
        XLValue.h
        XLValueBoolean.h
        XLValueEmpty.h
        XLValueError.h
        DESTINATION include/OpenXLSX)
install (FILES
        XML/rapidxml_iterators.hpp
        XML/rapidxml_print.hpp
        XML/rapidxml_utils.hpp
        XML/rapidxml.hpp
        XML/XMLNode.h
        XML/XMLAttribute.h
        XML/XMLDocument.h
        DESTINATION include/OpenXLSX/XML)
install (FILES
        Zip/libzip++.h
        Zip/XLArchive.h
        DESTINATION include/OpenXLSX/Zip)

