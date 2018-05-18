set(SOURCE_FILES
        OpenXLSX.h

        XLWorkbook/XLDocument.cpp XLWorkbook/XLDocument.h
        XLWorkbook/XLRelationships.cpp XLWorkbook/XLRelationships.h
        XLWorkbook/XLAbstractXMLFile.cpp XLWorkbook/XLAbstractXMLFile.h
        XLWorkbook/XLContentTypes.cpp XLWorkbook/XLContentTypes.h
        XLWorkbook/XLAppProperties.cpp XLWorkbook/XLAppProperties.h
        XLWorkbook/XLCoreProperties.cpp XLWorkbook/XLCoreProperties.h
        XLWorkbook/XLSharedStrings.cpp XLWorkbook/XLSharedStrings.h
        XLWorkbook/XLStyles.cpp XLWorkbook/XLStyles.h
        XLWorkbook/XLWorkbook.cpp XLWorkbook/XLWorkbook.h
        XLWorkbook/XLColor.cpp XLWorkbook/XLColor.h
        XLWorkbook/XLFont.cpp XLWorkbook/XLFont.h
        XLWorkbook/XLSpreadsheetElement.cpp XLWorkbook/XLSpreadsheetElement.h
        XLWorkbook/XLTokenizer.cpp XLWorkbook/XLTokenizer.h
        XLWorkbook/XLException.h

        XLSheet/XLSheet.cpp XLSheet/XLSheet.h
        XLSheet/XLWorksheet.cpp XLSheet/XLWorksheet.h
        XLSheet/XLChartsheet.cpp XLSheet/XLChartsheet.h
        XLSheet/XLTemplate.h
        XLSheet/XLRow.cpp XLSheet/XLRow.h
        XLSheet/XLColumn.cpp XLSheet/XLColumn.h

        XLCell/XLCell.cpp XLCell/XLCell.h
        XLCell/XLCellReference.cpp XLCell/XLCellReference.h
        XLCell/XLCellRange.cpp XLCell/XLCellRange.h
        XLCell/XLCellIterator.cpp XLCell/XLCellIterator.h
        XLCell/XLCellFormats.cpp XLCell/XLCellFormats.h
        XLCell/XLCellValue.cpp XLCell/XLCellValue.h
        XLCell/XLValueString.cpp XLCell/XLValueString.h
        XLCell/XLCellFormula.cpp XLCell/XLCellFormula.h
        XLCell/XLCellType.h
        XLCell/XLValueNumber.cpp XLCell/XLValueNumber.h
        XLCell/XLValue.cpp XLCell/XLValue.h
        XLCell/XLValueBoolean.cpp XLCell/XLValueBoolean.h
        XLCell/XLValueEmpty.cpp XLCell/XLValueEmpty.h
        XLCell/XLValueError.cpp XLCell/XLValueError.h

        Utilities/Zip/libzip++.h Utilities/Zip/libzip++.cpp
        Utilities/XML/pugiconfig.hpp Utilities/XML/pugixml.hpp Utilities/XML/pugixml.cpp
        Utilities/XML/XML.h)


add_library(OpenXLSX STATIC ${SOURCE_FILES})
target_include_directories(OpenXLSX PUBLIC "${LIBZIP_INCLUDE_DIR}")
target_link_libraries(OpenXLSX PUBLIC ${LIBZIP_LIBRARY})

install (TARGETS OpenXLSX DESTINATION lib)
install (FILES
        OpenXLSX.h
        XLWorkbook/XLDocument.h
        XLWorkbook/XLRelationships.h
        XLWorkbook/XLAbstractXMLFile.h
        XLWorkbook/XLContentTypes.h
        XLWorkbook/XLAppProperties.h
        XLWorkbook/XLCoreProperties.h
        XLWorkbook/XLSharedStrings.h
        XLWorkbook/XLStyles.h
        XLWorkbook/XLWorkbook.h
        XLWorkbook/XLColor.h
        XLWorkbook/XLFont.h
        XLWorkbook/XLSpreadsheetElement.h
        XLWorkbook/XLException.h

        XLSheet/XLSheet.h
        XLSheet/XLWorksheet.h
        XLSheet/XLChartsheet.h
        XLSheet/XLTemplate.h
        XLSheet/XLRow.h
        XLSheet/XLColumn.h

        XLCell/XLCell.h
        XLCell/XLCellReference.h
        XLCell/XLCellRange.h
        XLCell/XLCellIterator.h
        XLCell/XLCellFormats.h
        XLCell/XLCellValue.h
        XLCell/XLValueString.h
        XLCell/XLCellFormula.h
        XLCell/XLCellType.h
        XLCell/XLValueNumber.h
        XLCell/XLValue.h
        XLCell/XLValueBoolean.h
        XLCell/XLValueEmpty.h
        XLCell/XLValueError.h

        DESTINATION include/OpenXLSX)

install (FILES
        Utilities/Zip/libzip++.h
        DESTINATION include/OpenXLSX/Zip)

