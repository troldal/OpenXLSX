//
// Created by Troldal on 02/09/16.
//

#include <pugixml.hpp>

#include "XLCell_Impl.h"
#include "XLWorksheet_Impl.h"
#include "XLCellRange_Impl.h"

using namespace std;
using namespace OpenXLSX;

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
        : XLSpreadsheetElement(*parent.ParentDocument()),
          m_parentWorksheet(&parent),
          m_cellNode(cellNode),
          m_cellReference(XLCellReference(cellNode.attribute("r").value())),
          m_value(*this) {
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

    auto            first = this->CellReference();
    XLCellReference last(first->Row() + range.NumRows() - 1, first->Column() + range.NumColumns() - 1);
    XLCellRange     rng(*ParentWorksheet(), *first, last);
    rng = range;

    return *this;
}

/**
 * @details
 */
XLValueType Impl::XLCell::ValueType() const {

    return m_value.ValueType();
}

/**
 * @details
 */
const Impl::XLCellValue& Impl::XLCell::Value() const {

    return m_value;
}

/**
 * @details
 */
Impl::XLCellValue& Impl::XLCell::Value() {

    return m_value;
}

/**
 * @details This function returns a const reference to the cellReference property.
 */
const Impl::XLCellReference* Impl::XLCell::CellReference() const {

    return &m_cellReference;
}

/**
 * @details Creates a new std::unique_ptr using a raw pointer to the new object
 * @warning The std::make_unique function doesn't work, because the constructor is protected. Therefore,
 * a new unique_ptr is created manually, which is unfortunate, as it requires a call to operator new which is unsafe.
 */
std::unique_ptr<Impl::XLCell> Impl::XLCell::CreateCell(XLWorksheet& parent, XMLNode cellNode) {

    return unique_ptr<XLCell>(new XLCell(parent, cellNode));
}

/**
 * @details
 */
void Impl::XLCell::SetTypeAttribute(const std::string& typeString) {

    if (typeString.empty()) {
        if (m_cellNode.attribute("t") != nullptr)
            m_cellNode.remove_attribute("t");
    }
    else {
        if (m_cellNode.attribute("t") == nullptr)
            m_cellNode.append_attribute("t") = typeString.c_str();
        else
            m_cellNode.attribute("t") = typeString.c_str();
    }
}

/**
 * @details
 */
void Impl::XLCell::DeleteTypeAttribute() {

    if (m_cellNode.attribute("t").as_bool()) {
        m_cellNode.remove_attribute("t");
    }
}

/**
 * @details
 */
Impl::XLWorksheet* Impl::XLCell::ParentWorksheet() {

    return m_parentWorksheet;
}

/**
 * @details
 */
const Impl::XLWorksheet* Impl::XLCell::ParentWorksheet() const {

    return m_parentWorksheet;
}

/**
 * @details
 */
XMLDocument* Impl::XLCell::XmlDocument() {

    return ParentWorksheet()->XmlDocument();
}

/**
 * @details
 */
const XMLDocument* Impl::XLCell::XmlDocument() const {

    return ParentWorksheet()->XmlDocument();
}

/**
 * @details
 */
XMLNode Impl::XLCell::CellNode() {

    return m_cellNode;
}

/**
 * @details
 */
const XMLNode Impl::XLCell::CellNode() const {

    return m_cellNode;
}

/**
 * @details
 */
XMLNode Impl::XLCell::CreateValueNode() {

    if (!m_cellNode.child("v"))
        m_cellNode.append_child("v");
    return m_cellNode.child("v");
}

/**
 * @details
 */
bool Impl::XLCell::HasTypeAttribute() const {

    return m_cellNode.attribute("t") != nullptr;
}

/**
 * @details Return the cell type attribute, by querying the attribute named "t" in the XML node for the cell. 
 * If the cell has no attribute (i.e. is empty or holds a number), a nullptr will be returned.
 */
const XMLAttribute Impl::XLCell::TypeAttribute() const {

    return m_cellNode.attribute("t");
}