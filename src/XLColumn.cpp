//
// Created by Kenneth Balslev on 02/06/2017.
//

#include "XLColumn.h"
#include "XLWorksheet.h"

using namespace std;
using namespace OpenXLSX;

/**
 * @details Assumes each node only has data for one column.
 */
XLColumn::XLColumn(XLWorksheet &parent,
                   XMLNode &columnNode)
    : m_parentWorksheet(&parent),
      m_parentDocument(parent.ParentDocument()),
      m_columnNode(&columnNode),
      m_width(10),
      m_hidden(false),
      m_column(0)
{
    // Read the 'min' attribute of the Column
    auto minmaxAtt = m_columnNode->Attribute("min");
    if (minmaxAtt != nullptr) m_column = stoul(minmaxAtt->Value());

    // Read the 'Width' attribute of the Column
    auto widthAtt = m_columnNode->Attribute("width");
    if (widthAtt != nullptr) m_width = stof(widthAtt->Value());

    // Read the 'hidden' attribute of the Column.
    auto hiddenAtt = m_columnNode->Attribute("hidden");
    if (hiddenAtt != nullptr) {
        if (hiddenAtt->Value() == "1") m_hidden = true;
        else m_hidden = false;
    }
}

/**
 * @details
 */
float XLColumn::Width() const
{
    return m_width;
}

/**
 * @details
 */
void XLColumn::SetWidth(float width)
{
    m_width = width;

    // Set the 'Width' attribute for the Cell. If it does not exist, create it.
    auto widthAtt = m_columnNode->Attribute("width");
    if (widthAtt == nullptr) {
        widthAtt = m_parentWorksheet->XmlDocument()->CreateAttribute("width", to_string(m_width));
        m_columnNode->AppendAttribute(widthAtt);
    }
    else {
        widthAtt->SetValue(m_width);
    }

    // Set the 'customWidth' attribute for the Cell. If it does not exist, create it.
    auto customAtt = m_columnNode->Attribute("customWidth");
    if (customAtt == nullptr) {
        customAtt = m_parentWorksheet->XmlDocument()->CreateAttribute("customWidth", "1");
        m_columnNode->AppendAttribute(customAtt);
    }
    else {
        customAtt->SetValue("1");
    }
}

/**
 * @details
 */
bool XLColumn::Ishidden() const
{
    return m_hidden;
}

/**
 * @details
 */
void XLColumn::SetHidden(bool state)
{
    m_hidden = state;

    std::string hiddenstate;
    if (m_hidden) hiddenstate = "1";
    else hiddenstate = "0";

    // Set the 'hidden' attribute for the Cell. If it does not exist, create it.
    auto hiddenAtt = m_columnNode->Attribute("hidden");
    if (hiddenAtt == nullptr) {
        hiddenAtt = m_parentWorksheet->XmlDocument()->CreateAttribute("hidden", hiddenstate);
        m_columnNode->AppendAttribute(hiddenAtt);
    }
    else {
        hiddenAtt->SetValue(hiddenstate);
    }
}

/**
 * @details
 */
XMLNode *XLColumn::ColumnNode()
{
    return m_columnNode;
}
