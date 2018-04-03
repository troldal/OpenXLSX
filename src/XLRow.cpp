//
// Created by KBA012 on 02/06/2017.
//

#include "XLRow.h"
#include <boost/algorithm/string.hpp>
#include "XLWorksheet.h"
#include "XLCell.h"
#include "XLCellValue.h"
#include "XLException.h"

using namespace std;
using namespace OpenXLSX;
using namespace boost::algorithm;


/**
 * @details Constructs a new XLRow object from information in the underlying XML file. A pointer to the corresponding
 * node in the underlying XML file must be provided.
 */
XLRow::XLRow(XLWorksheet &parent,
             XMLNode &rowNode)
    : m_parentWorksheet(parent),
      m_parentDocument(*parent.ParentDocument()),
      m_rowNode(rowNode),
      m_height(15),
      m_descent(0.25),
      m_hidden(false),
      m_rowNumber(0),
      m_cells(0)
{
    // Read the Row number attribute
    auto rowAtt = m_rowNode.attribute("r");
    if (rowAtt != nullptr) m_rowNumber = stoul(rowAtt->value());

    // Read the Descent attribute
    auto descentAtt = m_rowNode.attribute("x14ac:dyDescent");
    if (descentAtt != nullptr) SetDescent(stof(descentAtt->value()));

    // Read the Row Height attribute
    auto heightAtt = m_rowNode.attribute("ht");
    if (heightAtt != nullptr) SetHeight(stof(heightAtt->value()));

    // Read the hidden attribute
    auto hiddenAtt = m_rowNode.attribute("hidden");
    if (hiddenAtt != nullptr) {
        if (hiddenAtt->value() == "1") SetHidden(true);
        else SetHidden(false);
    }

    // Iterate throught the Cell nodes and add cells to the m_cells vector
    auto spansAtt = m_rowNode.attribute("spans");
    if (spansAtt != nullptr) {
        XMLNode *currentCell = rowNode.childNode();
        while (currentCell != nullptr) {
            XLCellReference cellRef(currentCell->attribute("r")->value());
            Resize(cellRef.Column());
            m_cells.at(cellRef.Column() - 1) = XLCell::CreateCell(m_parentWorksheet, *currentCell);

            currentCell = currentCell->nextSibling();
        }
    }
}

/**
 * @details Resizes the m_cells vector holding the cells and updates the 'spans' attribure in the row node.
 */
void XLRow::Resize(unsigned int cellCount)
{
    m_cells.resize(cellCount);
    m_rowNode.attribute("spans")->setValue("1:" + to_string(cellCount));
}

/**
 * @details Returns the m_height member by value.
 */
float XLRow::Height() const
{
    return m_height;
}

/**
 * @details Set the height of the row. This is done by setting the value of the 'ht' attribute and setting the
 * 'customHeight' attribute to true.
 */
void XLRow::SetHeight(float height)
{
    m_height = height;

    // Set the 'ht' attribute for the Cell. If it does not exist, create it.
    auto heightAtt = m_rowNode.attribute("ht");
    if (heightAtt == nullptr) {
        heightAtt = m_parentWorksheet.XmlDocument().createAttribute("ht", to_string(m_height));
        m_rowNode.appendAttribute(heightAtt);
    }
    else {
        heightAtt->setValue(height);
    }

    // Set the 'customHeight' attribute. If it does not exist, create it.
    auto customAtt = m_rowNode.attribute("customHeight");
    if (customAtt == nullptr) {
        customAtt = m_parentWorksheet.XmlDocument().createAttribute("customHeight", "1");
        m_rowNode.appendAttribute(customAtt);
    }
    else {
        customAtt->setValue("1");
    }
}

/**
 * @details Return the m_descent member by value.
 */
float XLRow::Descent() const
{
    return m_descent;
}

/**
 * @details Set the descent by setting the 'x14ac:dyDescent' attribute in the XML file
 */
void XLRow::SetDescent(float descent)
{
    m_descent = descent;

    // Set the 'x14ac:dyDescent' attribute. If it does not exist, create it.
    auto descentAtt = m_rowNode.attribute("x14ac:dyDescent");
    if (descentAtt == nullptr) {
        descentAtt = m_parentWorksheet.XmlDocument().createAttribute("x14ac:dyDescent", to_string(m_descent));
        m_rowNode.appendAttribute(descentAtt);
    }
    else {
        descentAtt->setValue(descent);
    }
}

/**
 * @details Determine if the row is hidden or not.
 */
bool XLRow::Ishidden() const
{
    return m_hidden;
}

/**
 * @details Set the hidden state by setting the 'hidden' attribute to true or false.
 */
void XLRow::SetHidden(bool state)
{
    m_hidden = state;

    std::string hiddenstate;
    if (m_hidden) hiddenstate = "1";
    else hiddenstate = "0";

    // Set the 'hidden' attribute. If it does not exist, create it.
    auto hiddenAtt = m_rowNode.attribute("hidden");
    if (hiddenAtt == nullptr) {
        hiddenAtt = m_parentWorksheet.XmlDocument().createAttribute("hidden", hiddenstate);
        m_rowNode.appendAttribute(hiddenAtt);
    }
    else {
        hiddenAtt->setValue(hiddenstate);
    }
}

/**
 * @details Get the pointer to the row node in the underlying XML file but returning the m_rowNode member.
 */
XMLNode &XLRow::RowNode()
{
    return m_rowNode;
}

/**
 * @details Return a pointer to the XLCell object at the given column number. If the cell does not exist, it will be
 * created.
 */
XLCell &XLRow::Cell(unsigned int column)
{

    XLCell *result = nullptr;

    // If the requested Column number is higher than the number of Columns in the current Row,
    // create a new Cell node, append it to the Row node, resize the m_cells vector, and insert the new node.
    if (column > CellCount()) {

        // Create the new Cell node
        auto cellNode = m_parentWorksheet.XmlDocument().createNode("c");
        auto cellAttr = m_parentWorksheet.XmlDocument().createAttribute("r", XLCellReference(m_rowNumber,
                                                                                             column).Address());
        cellNode->appendAttribute(cellAttr);

        // Append the Cell node to the Row node, and create a new XLCell node and insert it in the m_cells vector.
        m_rowNode.appendNode(cellNode);
        Resize(column);
        m_cells.at(column - 1) = XLCell::CreateCell(m_parentWorksheet, *cellNode);

        result = m_cells.at(column - 1).get();

        // If the requested Column number is lower than the number of Columns in the current Row,
        // but the Cell does not exist, create a new node and insert it at the rigth position.
    }
    else if (m_cells.at(column - 1).get() == nullptr) {

        // Create the new Cell node.
        auto cellNode = m_parentWorksheet.XmlDocument().createNode("c");
        auto cellAttr = m_parentWorksheet.XmlDocument().createAttribute("r", XLCellReference(m_rowNumber,
                                                                                             column).Address());
        cellNode->appendAttribute(cellAttr);

        // Find the next Cell node and insert the new node at that position.
        auto cell = m_cells.at(column - 1).get();
        auto index = column - 1;
        while (cell == nullptr && index < m_cells.size()) cell = m_cells.at(index++).get();

        m_rowNode.insertNode(cell->CellNode(), cellNode);
        m_cells.at(column - 1) = XLCell::CreateCell(m_parentWorksheet, *cellNode);

        result = m_cells.at(column - 1).get();

        // If the Cell exists, look it up in the m_cells vector.
    }
    else {
        result = m_cells.at(column - 1).get();
    }

    return *result;
}

/**
 * @details
 */
const XLCell &XLRow::Cell(unsigned int column) const
{

    if (column > CellCount()) {
        throw XLException("Cell does not exist!");
    }
    else {
        return *m_cells.at(column - 1);
    }
}

/**
 * @details Get the number of cells in the row, by returning the size of the m_cells vector.
 */
unsigned int XLRow::CellCount() const
{
    return static_cast<unsigned int>(m_cells.size());
}

/**
 * @details Create an entirely new XLRow object (no existing node in the underlying XML file. The row and the first cell
 * in the row are created and inserted in the XML tree. A std::unique_ptr to the new XLRow object is returned
 * @todo What happens if the object is not caught and gets destroyed? The XML data still exists...
 */
unique_ptr<XLRow> XLRow::CreateRow(XLWorksheet &worksheet,
                                   unsigned long rowNumber)
{

    // Create the node
    auto nodeRow = worksheet.XmlDocument().createNode("row");
    auto attrRowNum = worksheet.XmlDocument().createAttribute("r", to_string(rowNumber));
    auto attrDescent = worksheet.XmlDocument().createAttribute("x14ac:dyDescent", "0.2");
    auto attrSpans = worksheet.XmlDocument().createAttribute("spans", "1:1");
    nodeRow->appendAttribute(attrRowNum);
    nodeRow->appendAttribute(attrSpans);
    nodeRow->appendAttribute(attrDescent);

    // Create the first Cell (at least one Cell must exist in each Row).
    auto cellNode = worksheet.XmlDocument().createNode("c");
    auto cellAttr = worksheet.XmlDocument().createAttribute("r", XLCellReference(rowNumber, 1).Address());
    cellNode->appendAttribute(cellAttr);
    nodeRow->appendNode(cellNode);

    // Insert the newly created Row and Cell in the right place in the XML structure.
    if (worksheet.SheetDataNode().childNode() == nullptr || rowNumber >= worksheet.RowCount()) {
        // If there are no Row nodes, or the requested Row number exceed the current maximum, append the the node
        // after the existing rownodes.
        worksheet.SheetDataNode().appendNode(nodeRow);
    }
    else {
        //Otherwise, search the Row nodes vector for the next node and insert there.
        auto index = rowNumber - 1; // vector is 0-based, Excel is 1-based; therefore rowNumber-1.
        XLRow *node = worksheet.Rows().at(index).get();
        while (node == nullptr) node = worksheet.Rows().at(index++).get();
        worksheet.SheetDataNode().insertNode(&node->RowNode(), nodeRow);
    }

    return make_unique<XLRow>(worksheet, *nodeRow);
}
