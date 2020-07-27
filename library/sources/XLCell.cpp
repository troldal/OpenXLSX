//
// Created by Troldal on 02/09/16.
//

#include "XLCell.hpp"

#include "XLCellRange.hpp"

using namespace std;
using namespace OpenXLSX;

XLCell::XLCell() : m_cellNode(XMLNode()), m_sharedStrings(nullptr) {}

/**
 * @details This constructor creates a XLCell object based on the cell XMLNode input parameter, and is
 * intended for use when the corresponding cell XMLNode already exist.
 * If a cell XMLNode does not exist (i.e., the cell is empty), use the relevant constructor to create an XLCell
 * from a XLCellReference parameter.
 * The initialization algorithm is as follows:
 *  -# Is the parameter a nullptr? If yes, throw a invalid_argument exception.
 *  -# Read the formula and value child nodes. These may be nullptr, if the cell is empty.
 *  -# Read the address, type and style attributes to the cellNode. The type attribute may be nullptr
 *  -# set the m_cellReference property using the value of the m_addressAttribute.
 *  -# Set the cell type, using the previous information:
 *      -# If there is no type attribute and no value node, the cell is empty.
 *      -# If there is no type attribute but there is a value node, the cell has a number value.
 *      -# Otherwise, determine the celltype based on the type attribute.
 */
XLCell::XLCell(XMLNode cellNode, XLSharedStrings* sharedStrings) : m_cellNode(cellNode), m_sharedStrings(sharedStrings) {}

/**
 * @details This methods copies a range into a new location, with the top left cell being located in the target cell.
 * The copying is done in the following way:
 * - Define a range with the top left cell being the target cell.
 * - Calculate the size of the range from the originating range and set the new range to the same size.
 * - Copy the contents of the original range to the new range.
 * - Return a reference to the first cell in the new range.
 * @pre
 * @post
 * @todo Consider what happens if the target range extends beyond the maximum spreadsheet limits
 */
XLCell& XLCell::operator=(const XLCellRange& range)
{
    auto            first = cellReference();
    XLCellReference last(first.row() + range.numRows() - 1, first.column() + range.numColumns() - 1);
    XLCellRange     rng(m_cellNode.parent().parent(), first, last);
    rng = range;

    return *this;
}

/**
 * @details
 */
XLValueType XLCell::valueType() const
{
    return XLCellValue(m_cellNode, nullptr).valueType();
}

/**
 * @details
 */
XLCellValue XLCell::value() const
{
    return XLCellValue(m_cellNode, m_sharedStrings);
}

/**
 * @details This function returns a const reference to the cellReference property.
 */
XLCellReference XLCell::cellReference() const
{
    return XLCellReference(m_cellNode.attribute("r").value());
}

/**
 * @details
 */
bool XLCell::hasFormula() const
{
    return m_cellNode.child("f") != nullptr;
}

/**
 * @details
 */
std::string XLCell::formula() const
{
    return m_cellNode.child("f").text().get();
}

/**
 * @details
 */
void XLCell::setFormula(const std::string& newFormula)
{
    m_cellNode.child("f").text().set(newFormula.c_str());
}

void XLCell::reset(XMLNode cellNode)
{
    m_cellNode = cellNode;
}
