//
// Created by Kenneth Balslev on 02/06/2017.
//

#include <pugixml.hpp>

#include "XLWorkbook_Impl.hpp"
#include "XLWorksheet_Impl.hpp"
#include "XLColumn_Impl.hpp"

using namespace std;
using namespace OpenXLSX;

/**
 * @details Assumes each node only has data for one column.
 */
Impl::XLColumn::XLColumn(XMLNode columnNode)
        : m_columnNode(columnNode) {}

/**
 * @details
 */
float Impl::XLColumn::Width() const {

    return ColumnNode().attribute("width").as_float();
}

/**
 * @details
 */
void Impl::XLColumn::SetWidth(float width) {

    // Set the 'Width' attribute for the Cell. If it does not exist, create it.
    auto widthAtt = ColumnNode().attribute("width");
    if (!widthAtt)
        widthAtt = ColumnNode().append_attribute("width");

    widthAtt.set_value(width);

    // Set the 'customWidth' attribute for the Cell. If it does not exist, create it.
    auto customAtt = ColumnNode().attribute("customWidth");
    if (!customAtt)
        customAtt =ColumnNode().append_attribute("customWidth");

    customAtt.set_value("1");
}

/**
 * @details
 */
bool Impl::XLColumn::IsHidden() const {

    return ColumnNode().attribute("hidden").as_bool();
}

/**
 * @details
 */
void Impl::XLColumn::SetHidden(bool state) {

    auto hiddenAtt = ColumnNode().attribute("hidden");
    if (!hiddenAtt)
        hiddenAtt = ColumnNode().append_attribute("hidden");

    if (state)
        hiddenAtt.set_value("1");
    else
        hiddenAtt.set_value("0");
}

/**
 * @details
 */
XMLNode Impl::XLColumn::ColumnNode() const {

    return m_columnNode;
}
