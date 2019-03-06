//
// Created by KBA012 on 02/06/2017.
//

#include <algorithm>
#include <pugixml.hpp>

#include "XLRow_Impl.h"
#include "XLWorksheet_Impl.h"

using namespace std;
using namespace OpenXLSX;

/**
 * @details Constructs a new XLRow object from information in the underlying XML file. A pointer to the corresponding
 * node in the underlying XML file must be provided.
 */
Impl::XLRow::XLRow(XLWorksheet& parent, XMLNode rowNode)
        : m_parentWorksheet(parent),
          m_rowNode(rowNode) {

    // Iterate through the Cell nodes and add cells to the m_cells hashmap
    for (auto& cell : m_rowNode.children())
        m_cells.emplace(XLCellReference(cell.attribute("r").value()).Column() - 1,
                        XLCell::CreateCell(m_parentWorksheet, cell));

    m_rowNode.attribute("spans") = string("1:" + to_string(XLCellReference(m_rowNode.last_child().attribute("r").value()).Column())).c_str();
}

/**
 * @details Returns the m_height member by value.
 */
float Impl::XLRow::Height() const {

    return m_rowNode.attribute("ht").as_float(15.0);
}

/**
 * @details Set the height of the row. This is done by setting the value of the 'ht' attribute and setting the
 * 'customHeight' attribute to true.
 */
void Impl::XLRow::SetHeight(float height) {

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
float Impl::XLRow::Descent() const {

    return m_rowNode.attribute("x14ac:dyDescent").as_float(0.25);
}

/**
 * @details Set the descent by setting the 'x14ac:dyDescent' attribute in the XML file
 */
void Impl::XLRow::SetDescent(float descent) {

    // Set the 'x14ac:dyDescent' attribute. If it does not exist, create it.
    if (!m_rowNode.attribute("x14ac:dyDescent"))
        m_rowNode.append_attribute("x14ac:dyDescent") = descent;
    else
        m_rowNode.attribute("x14ac:dyDescent") = descent;
}

/**
 * @details Determine if the row is hidden or not.
 */
bool Impl::XLRow::IsHidden() const {

    return m_rowNode.attribute("hidden").as_bool(false);
}

/**
 * @details Set the hidden state by setting the 'hidden' attribute to true or false.
 */
void Impl::XLRow::SetHidden(bool state) {

    // Set the 'hidden' attribute. If it does not exist, create it.
    if (!m_rowNode.attribute("hidden"))
        m_rowNode.append_attribute("hidden") = static_cast<int>(state);
    else
        m_rowNode.attribute("hidden").set_value(static_cast<int>(state));
}

/**
 * @details
 */
int64_t Impl::XLRow::RowNumber() const {

    return m_rowNode.attribute("r").as_ullong();
}

/**
 * @details Get the pointer to the row node in the underlying XML file but returning the m_rowNode member.
 */
XMLNode Impl::XLRow::RowNode() const {

    return m_rowNode;
}

/**
 * @details Return a pointer to the XLCell object at the given column number. If the cell does not exist, it will be
 * created.
 */
Impl::XLCell* Impl::XLRow::Cell(unsigned int column) {

    // If the requested Column number is higher than the number of Columns in the current Row,
    // create a new Cell node, append it to the Row node, resize the m_cells vector, and insert the new node.
    XLCell* result = nullptr;
    auto item = m_cells.equal_range(column - 1).first;
    if (item == m_cells.end()) {
        // Create the new Cell node
        auto cellNode = m_rowNode.append_child("c");
        cellNode.append_attribute("r").set_value(XLCellReference(RowNumber(), column).Address().c_str());

        // Append the Cell node to the Row node, and create a new XLCell node and insert it in the m_cells vector.
        m_rowNode.attribute("spans") = string("1:" + to_string(column)).c_str();
        result = m_cells.emplace(column - 1, XLCell::CreateCell(m_parentWorksheet, cellNode)).first->second.get();

        // If the requested Column number is lower than the number of Columns in the current Row,
        // but the Cell does not exist, create a new node and insert it at the rigth position.
    }
    else if ((*item).second->CellReference()->Column() != column) {

        // Find the next Cell node and insert the new node at that position.
        auto cellNode = m_rowNode.insert_child_before("c", (*item).second->CellNode());
        cellNode.append_attribute("r") = XLCellReference(RowNumber(), column).Address().c_str();
        result = m_cells.emplace(column - 1, XLCell::CreateCell(m_parentWorksheet, cellNode)).first->second.get();
    }
    else {
        result = m_cells.at(column - 1).get();
    }

    return result;
}

/**
 * @details
 */
const Impl::XLCell* Impl::XLRow::Cell(unsigned int column) const {

    if (column > CellCount()) throw XLException("Cell does not exist!");

    return m_cells.at(column - 1).get();

}

/**
 * @details Get the number of cells in the row, by returning the size of the m_cells vector.
 */
unsigned int Impl::XLRow::CellCount() const {

    return static_cast<unsigned int>(m_cells.size());
}
