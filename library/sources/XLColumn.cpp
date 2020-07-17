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
float XLColumn::Width() const
{
    return ColumnNode().attribute("width").as_float();
}

/**
 * @details
 */
void XLColumn::SetWidth(float width)
{
    // Set the 'Width' attribute for the Cell. If it does not exist, create it.
    auto widthAtt = ColumnNode().attribute("width");
    if (!widthAtt) widthAtt = ColumnNode().append_attribute("width");

    widthAtt.set_value(width);

    // Set the 'customWidth' attribute for the Cell. If it does not exist, create it.
    auto customAtt = ColumnNode().attribute("customWidth");
    if (!customAtt) customAtt = ColumnNode().append_attribute("customWidth");

    customAtt.set_value("1");
}

/**
 * @details
 */
bool XLColumn::IsHidden() const
{
    return ColumnNode().attribute("hidden").as_bool();
}

/**
 * @details
 */
void XLColumn::SetHidden(bool state)
{
    auto hiddenAtt = ColumnNode().attribute("hidden");
    if (!hiddenAtt) hiddenAtt = ColumnNode().append_attribute("hidden");

    if (state)
        hiddenAtt.set_value("1");
    else
        hiddenAtt.set_value("0");
}

/**
 * @details
 */
XMLNode XLColumn::ColumnNode() const
{
    return m_columnNode;
}
