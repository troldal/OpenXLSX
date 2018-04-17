//
// Created by KBA012 on 02/06/2017.
//

#include "XLRow.h"
#include "XLWorksheet.h"
#include "XLCell.h"
#include "XLCellValue.h"
#include "XLException.h"

using namespace std;
using namespace OpenXLSX;


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
    auto rowAtt = m_rowNode.Attribute("r");
    if (rowAtt != nullptr) m_rowNumber = stoul(rowAtt->Value());

    // Read the Descent attribute
    auto descentAtt = m_rowNode.Attribute("x14ac:dyDescent");
    if (descentAtt != nullptr) SetDescent(stof(descentAtt->Value()));

    // Read the Row Height attribute
    auto heightAtt = m_rowNode.Attribute("ht");
    if (heightAtt != nullptr) SetHeight(stof(heightAtt->Value()));

    // Read the hidden attribute
    auto hiddenAtt = m_rowNode.Attribute("hidden");
    if (hiddenAtt != nullptr) {
        if (hiddenAtt->Value() == "1") SetHidden(true);
        else SetHidden(false);
    }

    // Iterate throught the Cell nodes and add cells to the m_cells vector
    auto spansAtt = m_rowNode.Attribute("spans");
    if (spansAtt != nullptr) {
        XMLNode *currentCell = rowNode.ChildNode();
        while (currentCell != nullptr) {
            XLCellReference cellRef(currentCell->Attribute("r")->Value());
            Resize(cellRef.Column());
            m_cells.at(cellRef.Column() - 1) = XLCell::CreateCell(m_parentWorksheet, *currentCell);

            currentCell = currentCell->NextSibling();
        }
    }
}

/**
 * @details Resizes the m_cells vector holding the cells and updates the 'spans' attribure in the row node.
 */
void XLRow::Resize(unsigned int cellCount)
{
    m_cells.resize(cellCount);
    m_rowNode.Attribute("spans")->SetValue("1:" + to_string(cellCount));
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
    auto heightAtt = m_rowNode.Attribute("ht");
    if (heightAtt == nullptr) {
        heightAtt = m_parentWorksheet.XmlDocument()->CreateAttribute("ht", to_string(m_height));
        m_rowNode.AppendAttribute(heightAtt);
    }
    else {
        heightAtt->SetValue(height);
    }

    // Set the 'customHeight' attribute. If it does not exist, create it.
    auto customAtt = m_rowNode.Attribute("customHeight");
    if (customAtt == nullptr) {
        customAtt = m_parentWorksheet.XmlDocument()->CreateAttribute("customHeight", "1");
        m_rowNode.AppendAttribute(customAtt);
    }
    else {
        customAtt->SetValue("1");
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
    auto descentAtt = m_rowNode.Attribute("x14ac:dyDescent");
    if (descentAtt == nullptr) {
        descentAtt = m_parentWorksheet.XmlDocument()->CreateAttribute("x14ac:dyDescent", to_string(m_descent));
        m_rowNode.AppendAttribute(descentAtt);
    }
    else {
        descentAtt->SetValue(descent);
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
    auto hiddenAtt = m_rowNode.Attribute("hidden");
    if (hiddenAtt == nullptr) {
        hiddenAtt = m_parentWorksheet.XmlDocument()->CreateAttribute("hidden", hiddenstate);
        m_rowNode.AppendAttribute(hiddenAtt);
    }
    else {
        hiddenAtt->SetValue(hiddenstate);
    }
}

/**
 * @details Get the pointer to the row node in the underlying XML file but returning the m_rowNode member.
 */
XMLNode *XLRow::RowNode()
{
    return &m_rowNode;
}

/**
 * @details Return a pointer to the XLCell object at the given column number. If the cell does not exist, it will be
 * created.
 */
XLCell *XLRow::Cell(unsigned int column)
{

    XLCell *result = nullptr;

    // If the requested Column number is higher than the number of Columns in the current Row,
    // create a new Cell node, append it to the Row node, resize the m_cells vector, and insert the new node.
    if (column > CellCount()) {

        // Create the new Cell node
        auto cellNode = m_parentWorksheet.XmlDocument()->CreateNode("c");
        auto cellAttr = m_parentWorksheet.XmlDocument()->CreateAttribute("r", XLCellReference(m_rowNumber,
                                                                                              column).Address());
        cellNode->AppendAttribute(cellAttr);

        // Append the Cell node to the Row node, and create a new XLCell node and insert it in the m_cells vector.
        m_rowNode.AppendNode(cellNode);
        Resize(column);
        m_cells.at(column - 1) = XLCell::CreateCell(m_parentWorksheet, *cellNode);

        result = m_cells.at(column - 1).get();

        // If the requested Column number is lower than the number of Columns in the current Row,
        // but the Cell does not exist, create a new node and insert it at the rigth position.
    }
    else if (m_cells.at(column - 1).get() == nullptr) {

        // Create the new Cell node.
        auto cellNode = m_parentWorksheet.XmlDocument()->CreateNode("c");
        auto cellAttr = m_parentWorksheet.XmlDocument()->CreateAttribute("r", XLCellReference(m_rowNumber,
                                                                                              column).Address());
        cellNode->AppendAttribute(cellAttr);

        // Find the next Cell node and insert the new node at that position.
        auto cell = m_cells.at(column - 1).get();
        auto index = column - 1;
        while (cell == nullptr && index < m_cells.size()) cell = m_cells.at(index++).get();

        m_rowNode.InsertNode(cell->CellNode(), cellNode);
        m_cells.at(column - 1) = XLCell::CreateCell(m_parentWorksheet, *cellNode);

        result = m_cells.at(column - 1).get();

        // If the Cell exists, look it up in the m_cells vector.
    }
    else {
        result = m_cells.at(column - 1).get();
    }

    return result;
}

/**
 * @details
 */
const XLCell *XLRow::Cell(unsigned int column) const
{

    if (column > CellCount()) {
        throw XLException("Cell does not exist!");
    }
    else {
        return m_cells.at(column - 1).get();
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
    auto nodeRow = worksheet.XmlDocument()->CreateNode("row");
    auto attrRowNum = worksheet.XmlDocument()->CreateAttribute("r", to_string(rowNumber));
    auto attrDescent = worksheet.XmlDocument()->CreateAttribute("x14ac:dyDescent", "0.2");
    auto attrSpans = worksheet.XmlDocument()->CreateAttribute("spans", "1:1");
    nodeRow->AppendAttribute(attrRowNum);
    nodeRow->AppendAttribute(attrSpans);
    nodeRow->AppendAttribute(attrDescent);

    // Create the first Cell (at least one Cell must exist in each Row).
    auto cellNode = worksheet.XmlDocument()->CreateNode("c");
    auto cellAttr = worksheet.XmlDocument()->CreateAttribute("r", XLCellReference(rowNumber, 1).Address());
    cellNode->AppendAttribute(cellAttr);
    nodeRow->AppendNode(cellNode);

    // Insert the newly created Row and Cell in the right place in the XML structure.
    if (worksheet.SheetDataNode()->ChildNode() == nullptr || rowNumber >= worksheet.RowCount()) {
        // If there are no Row nodes, or the requested Row number exceed the current maximum, append the the node
        // after the existing rownodes.
        worksheet.SheetDataNode()->AppendNode(nodeRow);
    }
    else {
        //Otherwise, search the Row nodes vector for the next node and insert there.
        auto index = rowNumber - 1; // vector is 0-based, Excel is 1-based; therefore rowNumber-1.
        XLRow *node = worksheet.Rows()->at(index).get();
        while (node == nullptr) node = worksheet.Rows()->at(index++).get();
        worksheet.SheetDataNode()->InsertNode(node->RowNode(), nodeRow);
    }

    return make_unique<XLRow>(worksheet, *nodeRow);
}
