//
// Created by KBA012 on 02/06/2017.
//

#include "XLRow_Impl.hpp"

#include "XLCellReference_Impl.hpp"

using namespace std;
using namespace OpenXLSX;

/**
 * @details Constructs a new XLRow object from information in the underlying XML file. A pointer to the corresponding
 * node in the underlying XML file must be provided.
 */
Impl::XLRow::XLRow(XMLNode rowNode) : m_rowNode(rowNode) {}

/**
 * @details Returns the m_height member by value.
 */
double Impl::XLRow::Height() const
{
    return m_rowNode.attribute("ht").as_double(15.0);
}

/**
 * @details Set the height of the row. This is done by setting the value of the 'ht' attribute and setting the
 * 'customHeight' attribute to true.
 */
void Impl::XLRow::SetHeight(float height)
{
    // Set the 'ht' attribute for the Cell. If it does not exist, create it.
    if (!m_rowNode.attribute("ht"))
        m_rowNode.append_attribute("ht") = height;
    else
        m_rowNode.attribute("ht").set_value(height);

    // Set the 'customHeight' attribute. If it does not exist, create it.
    if (!m_rowNode.attribute("customHeight"))
        m_rowNode.append_attribute("customHeight") = 1;
    else
        m_rowNode.attribute("customHeight").set_value(1);
}

/**
 * @details Return the m_descent member by value.
 */
float Impl::XLRow::Descent() const
{
    return m_rowNode.attribute("x14ac:dyDescent").as_float(0.25);
}

/**
 * @details Set the descent by setting the 'x14ac:dyDescent' attribute in the XML file
 */
void Impl::XLRow::SetDescent(float descent)
{
    // Set the 'x14ac:dyDescent' attribute. If it does not exist, create it.
    if (!m_rowNode.attribute("x14ac:dyDescent"))
        m_rowNode.append_attribute("x14ac:dyDescent") = descent;
    else
        m_rowNode.attribute("x14ac:dyDescent") = descent;
}

/**
 * @details Determine if the row is hidden or not.
 */
bool Impl::XLRow::IsHidden() const
{
    return m_rowNode.attribute("hidden").as_bool(false);
}

/**
 * @details Set the hidden state by setting the 'hidden' attribute to true or false.
 */
void Impl::XLRow::SetHidden(bool state)
{
    // Set the 'hidden' attribute. If it does not exist, create it.
    if (!m_rowNode.attribute("hidden"))
        m_rowNode.append_attribute("hidden") = static_cast<int>(state);
    else
        m_rowNode.attribute("hidden").set_value(static_cast<int>(state));
}

/**
 * @details
 */
int64_t Impl::XLRow::RowNumber() const
{
    return m_rowNode.attribute("r").as_ullong();
}

/**
 * @details Get the number of cells in the row, by returning the size of the m_cells vector.
 */
unsigned int Impl::XLRow::CellCount() const
{
    return XLCellReference(m_rowNode.last_child().attribute("r").value()).Column();
}
