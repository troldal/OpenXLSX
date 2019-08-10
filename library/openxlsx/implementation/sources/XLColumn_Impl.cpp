//
// Created by Kenneth Balslev on 02/06/2017.
//

#include <pugixml.hpp>

#include "XLWorkbook_Impl.h"
#include "XLWorksheet_Impl.h"
#include "XLColumn_Impl.h"

using namespace std;
using namespace OpenXLSX;

/**
 * @details Assumes each node only has data for one column.
 */
Impl::XLColumn::XLColumn(XLWorksheet& parent, XMLNode columnNode)
        : m_parentDocument(parent.Workbook()->Document()),
          m_columnNode(std::make_unique<XMLNode>(columnNode)),
          m_width(10),
          m_hidden(false),
          m_column(0) {
    // Read the 'min' attribute of the Column
    auto minmaxAtt = m_columnNode->attribute("min");
    if (minmaxAtt != nullptr)
        m_column = stoul(minmaxAtt.value());

    // Read the 'Width' attribute of the Column
    auto widthAtt = m_columnNode->attribute("width");
    if (widthAtt != nullptr)
        m_width = stof(widthAtt.value());

    // Read the 'hidden' attribute of the Column.
    auto hiddenAtt = m_columnNode->attribute("hidden");
    if (hiddenAtt != nullptr) {
        if (string(hiddenAtt.value()) == "1")
            m_hidden = true;
        else
            m_hidden = false;
    }
}

/**
 * @details
 */
float Impl::XLColumn::Width() const {

    return m_width;
}

/**
 * @details
 */
void Impl::XLColumn::SetWidth(float width) {

    m_width = width;

    // Set the 'Width' attribute for the Cell. If it does not exist, create it.
    auto widthAtt = m_columnNode->attribute("width");
    if (!widthAtt)
        m_columnNode->append_attribute("width") = m_width;
    else
        widthAtt.set_value(m_width);

    // Set the 'customWidth' attribute for the Cell. If it does not exist, create it.
    auto customAtt = m_columnNode->attribute("customWidth");
    if (!customAtt)
        m_columnNode->append_attribute("customWidth") = "1";
    else
        customAtt.set_value("1");
}

/**
 * @details
 */
bool Impl::XLColumn::IsHidden() const {

    return m_hidden;
}

/**
 * @details
 */
void Impl::XLColumn::SetHidden(bool state) {

    m_hidden = state;

    std::string hiddenstate;
    if (m_hidden)
        hiddenstate = "1";
    else
        hiddenstate = "0";

    // Set the 'hidden' attribute for the Cell. If it does not exist, create it.
    auto hiddenAtt = m_columnNode->attribute("hidden");
    if (!hiddenAtt)
        m_columnNode->append_attribute("hidden") = hiddenstate.c_str();
    else
        hiddenAtt.set_value(hiddenstate.c_str());
}

/**
 * @details
 */
XMLNode Impl::XLColumn::ColumnNode() {

    return *m_columnNode;
}
