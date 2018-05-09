//
// Created by Troldal on 04/10/16.
//

#include "XLChartsheet.h"
#include "../XLWorkbook/XLDocument.h"

using namespace OpenXLSX;

/**
 * @details
 */
XLChartsheet::XLChartsheet(XLWorkbook &parent,
                           const std::string &name,
                           const std::string &filePath,
                           const std::string &xmlData)
    : XLSheet(parent, name, filePath, xmlData)
{

}
