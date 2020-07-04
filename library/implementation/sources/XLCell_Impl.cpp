//
// Created by Troldal on 02/09/16.
//

//#include <pugixml.hpp>
#include <memory>

#include "XLCell_Impl.hpp"
#include "XLWorksheet_Impl.hpp"
#include "XLCellRange_Impl.hpp"

using namespace std;
using namespace OpenXLSX;

Impl::XLCell::XLCell()
        : m_parentWorksheet(nullptr),
          m_cellNode(XMLNode()) {

}

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
Impl::XLCell::XLCell(XLWorksheet& parent, XMLNode cellNode)
        : m_parentWorksheet(&parent),
          m_cellNode(cellNode) {

    // Empty constructor body
}

Impl::XLCell::XLCell(Impl::XLCell const& other)
        : m_parentWorksheet(other.m_parentWorksheet),
          m_cellNode(other.m_cellNode) {

    // Empty constructor body
}

Impl::XLCell::XLCell(Impl::XLCell&& other) noexcept
        : m_parentWorksheet(std::move(other.m_parentWorksheet)),
          m_cellNode(std::move(other.m_cellNode)) {

    // Empty constructor body
}

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
Impl::XLCell& Impl::XLCell::operator=(const XLCellRange& range) {

    auto first = CellReference();
    XLCellReference last(first.Row() + range.NumRows() - 1, first.Column() + range.NumColumns() - 1);
    XLCellRange rng(*m_parentWorksheet, first, last);
    rng = range;

    return *this;
}

/**
 * @details
 */
Impl::XLValueType Impl::XLCell::ValueType() const {

    return XLCellValue(*this).ValueType();
    //return m_value.ValueType();
}

/**
 * @details
 */
Impl::XLCellValue Impl::XLCell::Value() const {

    return XLCellValue(*this);
}

/**
 * @details This function returns a const reference to the cellReference property.
 */
Impl::XLCellReference Impl::XLCell::CellReference() const {

    return XLCellReference(m_cellNode.attribute("r").value());
}

/**
 * @details
 */
bool Impl::XLCell::HasFormula() const {
    return m_cellNode.child("f") != nullptr;
}

/**
 * @details
 */
std::string Impl::XLCell::GetFormula() const {
    return m_cellNode.child("f").text().get();
}

/**
 * @details
 */
void Impl::XLCell::SetFormula(const std::string& newFormula) {
    m_cellNode.child("f").text().set(newFormula.c_str());
}

void Impl::XLCell::reset(XMLNode cellNode) {
    m_cellNode = cellNode;
}
