//
// Created by Troldal on 04/10/16.
//

#include "XLChartsheet_Impl.hpp"
#include "XLDocument_Impl.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
Impl::XLChartsheet::XLChartsheet(XLWorkbook& parent, const std::string& sheetRID, XMLAttribute name, const std::string& filePath,
                                 const std::string& xmlData)
        : XLSheet(parent, sheetRID, name, filePath, xmlData) {

}

Impl::XLSheet* Impl::XLChartsheet::Clone(const std::string& newName) {

    return nullptr;
}

bool Impl::XLChartsheet::ParseXMLData() {

    return false;
}
