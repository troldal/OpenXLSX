//
// Created by Troldal on 04/10/16.
//

#include "XLChartsheet.hpp"

#include "XLDocument.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
XLChartsheet::XLChartsheet(XLXmlData* xmlData) : XLSheetBase(xmlData) {}

XLChartsheet XLChartsheet::Clone(const std::string& newName) {}

bool XLChartsheet::ParseXMLData()
{
    return false;
}
