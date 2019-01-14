//
// Created by Troldal on 04/10/16.
//

#include "XLChartsheet_Impl.h"
#include "XLDocument_Impl.h"

using namespace OpenXLSX;

/**
 * @details
 */
Impl::XLChartsheet::XLChartsheet(XLWorkbook &parent,
                           const std::string &name,
                           const std::string &filePath,
                           const std::string &xmlData)
    : XLSheet(parent, name, filePath, xmlData)
{

}
