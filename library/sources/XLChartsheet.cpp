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

XLChartsheet XLChartsheet::clone(const std::string& newName) {}
