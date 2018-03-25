//
// Created by Troldal on 04/10/16.
//

#include "XLChartsheet.h"
#include "XLDocument.h"

using namespace RapidXLSX;

/**
 * @details
 */
XLChartsheet::XLChartsheet(XLWorkbook &parent,
                           const std::string &name,
                           const std::string &filePath)
    : XLAbstractSheet(parent, name, filePath)
{

}
