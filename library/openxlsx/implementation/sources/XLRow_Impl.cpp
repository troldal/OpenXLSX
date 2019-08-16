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
    for (const auto& cell : m_rowNode.children()) {
        auto& newCell = m_cells.emplace_back(XLCellData());
        newCell.cellIndex = XLCellReference(cell.attribute("r").value()).Column();
        newCell.cellItem = make_unique<XLCell>(m_parentWorksheet, cell);
    }
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

    // ===== Search for the cell at the requested column
    XLCellData searchItem;
    searchItem.cellIndex = column;
    auto dataItem = lower_bound(m_cells.begin(), m_cells.end(), searchItem,
                                [](const XLCellData& a, const XLCellData& b) {
                                    return a.cellIndex < b.cellIndex;
                                });

    // ===== If cell does not exist, create it...
    if (dataItem == m_cells.end() || dataItem->cellIndex > column) {

        auto cellNode = XMLNode();

        // ===== Append or insert new cell node in the XML file
        if (dataItem == m_cells.end())
            cellNode = m_rowNode.append_child("c");
        else
            cellNode = m_rowNode.insert_child_before("c", dataItem->cellItem->m_cellNode);

        // ===== Add the newly created node to the std::vector and update the dataItem iterator
        dataItem = m_cells.insert(dataItem, XLCellData());
        dataItem->cellIndex = column;
        dataItem->cellItem = make_unique<XLCell>(m_parentWorksheet, cellNode);

        // ===== Set the correct address of the newly created cell node
        cellNode.append_attribute("r").set_value(XLCellReference(RowNumber(), column).Address().c_str());

    }

    return dataItem->cellItem.get();
}

/**
 * @details
 */
const Impl::XLCell* Impl::XLRow::Cell(unsigned int column) const {

    auto result = find_if(m_cells.begin(), m_cells.end(), [=](const XLCellData& data) {
        return data.cellIndex == column;
    });

    if (result == m_cells.end())
        throw XLException("Cell does not exist!");
    return result->cellItem.get();

}

/**
 * @details Get the number of cells in the row, by returning the size of the m_cells vector.
 */
unsigned int Impl::XLRow::CellCount() const {

    return static_cast<unsigned int>(m_cells.size());
}
