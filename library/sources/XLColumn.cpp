//
// Created by Kenneth Balslev on 02/06/2017.
//

#include "XLColumn.hpp"

#include "XLWorkbook.hpp"
#include "XLWorksheet.hpp"

using namespace std;
using namespace OpenXLSX;

/**
 * @details Assumes each node only has data for one column.
 */
XLColumn::XLColumn(XMLNode columnNode) : m_columnNode(columnNode) {}

/**
 * @details
 */
float XLColumn::width() const
{
    return columnNode().attribute("width").as_float();
}

/**
 * @details
 */
void XLColumn::setWidth(float width)
{
    // Set the 'Width' attribute for the Cell. If it does not exist, create it.
    auto widthAtt = columnNode().attribute("width");
    if (!widthAtt) widthAtt = columnNode().append_attribute("width");

    widthAtt.set_value(width);

    // Set the 'customWidth' attribute for the Cell. If it does not exist, create it.
    auto customAtt = columnNode().attribute("customWidth");
    if (!customAtt) customAtt = columnNode().append_attribute("customWidth");

    customAtt.set_value("1");
}

/**
 * @details
 */
bool XLColumn::isHidden() const
{
    return columnNode().attribute("hidden").as_bool();
}

/**
 * @details
 */
void XLColumn::setHidden(bool state)
{
    auto hiddenAtt = columnNode().attribute("hidden");
    if (!hiddenAtt) hiddenAtt = columnNode().append_attribute("hidden");

    if (state)
        hiddenAtt.set_value("1");
    else
        hiddenAtt.set_value("0");
}

/**
 * @details
 */
XMLNode XLColumn::columnNode() const
{
    return m_columnNode;
}
