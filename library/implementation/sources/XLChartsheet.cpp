//
// Created by Troldal on 04/10/16.
//

#include "XLChartsheet.hpp"

#include "XLDocument.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
Impl::XLChartsheet::XLChartsheet(XLXmlData* xmlData) : XLSheetBase(xmlData) {}

Impl::XLChartsheet Impl::XLChartsheet::Clone(const std::string& newName) {}

bool Impl::XLChartsheet::ParseXMLData()
{
    return false;
}
