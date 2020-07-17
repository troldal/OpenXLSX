//
// Created by Troldal on 04/10/16.
//

#include "XLChartsheet_Impl.hpp"

#include "XLDocument_Impl.hpp"

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
