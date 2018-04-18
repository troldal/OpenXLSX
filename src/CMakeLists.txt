set(SOURCE_FILES
        Zip/libzip++.h Zip/libzip++.cpp
        XML/rapidxml_iterators.hpp
        XML/rapidxml_print.hpp
        XML/rapidxml_utils.hpp
        XML/rapidxml.hpp
        XML/XMLNode.cpp XML/XMLNode.h
        XML/XMLAttribute.cpp XML/XMLAttribute.h
        XML/XMLDocument.cpp XML/XMLDocument.h
        OpenXLSX.h
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
        XLValueError.cpp XLValueError.h
        XLTokenizer.cpp XLTokenizer.h)

add_library(OpenXLSX STATIC ${SOURCE_FILES})
target_include_directories(OpenXLSX PUBLIC "${LIBZIP_INCLUDE_DIR}")
target_link_libraries(OpenXLSX PUBLIC ${LIBZIP_LIBRARY})

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
        DESTINATION include/OpenXLSX/Zip)

