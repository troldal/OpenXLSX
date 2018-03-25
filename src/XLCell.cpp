//
// Created by Troldal on 02/09/16.
//

#include "XLCell.h"
#include "XLCellValue.h"
#include "XLWorksheet.h"
#include "XLCellRange.h"
#include "XLException.h"


using namespace std;
using namespace RapidXLSX;


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
XLCell::XLCell(XLWorksheet &parent,
               XMLNode &cellNode)
    : XLSpreadsheetElement(parent.ParentDocument()),
      m_parentDocument(&parent.ParentDocument()),
      m_parentWorkbook(&parent.ParentWorkbook()),
      m_parentWorksheet(&parent),
      m_cellReference(XLCellReference(cellNode.attribute("r")->value())),
      m_rowNode(cellNode.parent()),
      m_cellNode(&cellNode),
      m_formulaNode(cellNode.childNode("f")),
      m_value(std::make_unique<XLCellValue>(*this))
{
}

/**
 * @details The copy assignment operator simply copies the value (and type) of the original cell object.
 * The target object already has a parent sheet and row node, and will handle creation of new cell node etc.
 * if required.
 * @pre The target and source XLCell objects are valid, i.e. are properly created and initialized.
 * @post The target XLCell has a type and value of the source object.
 */
XLCell &XLCell::operator=(const XLCell &other)
{
    *m_value = *other.m_value;
    //SetTypeAttribute(other.TypeString());
    //SetValue(other.CellType(), other.ValueString());
    return *this;
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
XLCell &XLCell::operator=(const XLCellRange &range)
{

    auto first = this->CellReference();
    XLCellReference last(first.Row() + range.NumRows() - 1, first.Column() + range.NumColumns() - 1);
    XLCellRange rng(ParentWorksheet(), first, last);
    rng = range;

    return *this;
}

XLValueType XLCell::ValueType() const
{
    return (*m_value).ValueType();
}

/**
 * @details
 */
const XLCellValue &XLCell::Value() const
{
    if (m_value == nullptr) throw XLException("Cell " + m_cellReference.Address() + " has no value.");
    return *m_value;
}

/**
 * @details
 */
XLCellValue &XLCell::Value()
{
    if (m_value == nullptr) m_value = std::make_unique<XLCellValue>(*this);
    return *m_value;
}

/**
 * @details
 */
void XLCell::SetModified()
{
    ParentWorksheet().SetModified();
}

/**
 * @details This function returns a const reference to the cellReference property.
 */
const XLCellReference &XLCell::CellReference() const
{
    return m_cellReference;
}

/**
 * @details Creates a new std::unique_ptr using a raw pointer to the new object
 * @warning The std::make_unique function doesn't work, because the constructor is protected. Therefore,
 * a new unique_ptr is created manually, which is unfortunate, as it requires a call to operator new which is unsafe.
 */
std::unique_ptr<XLCell> XLCell::CreateCell(XLWorksheet &parent,
                                           XMLNode &cellNode)
{
    return unique_ptr<XLCell>(new XLCell(parent, cellNode));
}

/**
 * @details
 */
void XLCell::SetTypeAttribute(const std::string &typeString)
{

    if (typeString.empty()) {
        if (m_cellNode->attribute("t"))
            m_cellNode->attribute("t")->deleteAttribute();
    }
    else {
        if (m_cellNode->attribute("t") == nullptr)
            m_cellNode->appendAttribute(XmlDocument().createAttribute("t", typeString));
        else
            m_cellNode->attribute("t")->setValue(typeString);
    }

    /*
    std::string typeString;

    switch (type) {
        case XLCellType::Empty:
            DeleteTypeAttribute();
            break;

        case XLCellType::Number :
            DeleteTypeAttribute();
            break;

        case XLCellType::String :
            typeString = "str";
            break;

        case XLCellType::InlineString :
        case XLCellType::SharedString :
            typeString = "s";
            break;

        case XLCellType::Boolean :
            typeString = "b";
            break;

        case XLCellType::Error :
            typeString = "e";
            break;
    }

    if (typeString.empty()) return;

    if (m_cellNode->attribute("t") == nullptr) {
        m_cellNode->appendAttribute(XmlDocument().createAttribute("t", typeString));
    }
    else {
        m_cellNode->attribute("t")->setValue(typeString);
    }*/
}

/**
 * @details
 */
void XLCell::DeleteTypeAttribute()
{
    if (m_cellNode->attribute("t") != nullptr) {
        m_cellNode->attribute("t")->deleteAttribute();
    }
}

/**
 * @details
 */
XLWorksheet &XLCell::ParentWorksheet()
{
    return *m_parentWorksheet;
}

/**
 * @details
 */
const XLWorksheet &XLCell::ParentWorksheet() const
{
    return *m_parentWorksheet;
}

/**
 * @details
 */
XMLDocument &XLCell::XmlDocument()
{
    return ParentWorksheet().XmlDocument();
}

/**
 * @details
 */
const XMLDocument &XLCell::XmlDocument() const
{
    return ParentWorksheet().XmlDocument();
}

/**
 * @details
 */
XMLNode *XLCell::CellNode()
{
    return m_cellNode;
}

/**
 * @details
 */
const XMLNode *XLCell::CellNode() const
{
    return m_cellNode;
}

XMLNode *XLCell::CreateValueNode()
{
    if (m_cellNode->childNode("v") == nullptr) m_cellNode->appendNode(XmlDocument().createNode("v"));
    return m_cellNode->childNode("v");
}

bool XLCell::HasTypeAttribute() const
{
    if (m_cellNode->attribute("t") == nullptr)
        return false;
    else
        return true;
}

/**
 * @details
 */
const XMLAttribute *XLCell::TypeAttribute() const
{
    return m_cellNode->attribute("t");
}










