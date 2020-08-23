//
// Created by Kenneth Balslev on 22/08/2020.
//

#include <pugixml.hpp>

#include "XLRowIterator.hpp"
#include "XLRowRange.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

namespace OpenXLSX
{
    XMLNode getRowNode(XMLNode sheetDataNode, uint32_t rowNumber);
}

/**
 * @details
 * @pre
 * @post
 */
XLRowIterator::XLRowIterator(const XLRowRange& rowRange, XLIteratorLocation loc)
    : m_dataNode(std::make_unique<XMLNode>(*rowRange.m_dataNode)),
      m_firstRow(rowRange.m_firstRow),
      m_lastRow(rowRange.m_lastRow),
      m_currentRow(),
      m_sharedStrings(rowRange.m_sharedStrings)
{
    if (loc == XLIteratorLocation::End)
        m_currentRow = XLRow();
    else {
        m_currentRow = XLRow(getRowNode(*m_dataNode, m_firstRow), m_sharedStrings);
    }
}

/**
 * @details
 * @pre
 * @post
 */
XLRowIterator::~XLRowIterator() = default;

/**
 * @details
 * @pre
 * @post
 */
XLRowIterator::XLRowIterator(const XLRowIterator& other)
    : m_dataNode(std::make_unique<XMLNode>(*other.m_dataNode)),
      m_firstRow(other.m_firstRow),
      m_lastRow(other.m_lastRow),
      m_currentRow(other.m_currentRow),
      m_sharedStrings(other.m_sharedStrings)
{}

/**
 * @details
 * @pre
 * @post
 */
XLRowIterator::XLRowIterator(XLRowIterator&& other) noexcept = default;

/**
 * @details
 * @pre
 * @post
 */
XLRowIterator& XLRowIterator::operator=(const XLRowIterator& other)
{
    if (&other != this) {
        *m_dataNode     = *other.m_dataNode;
        m_firstRow       = other.m_firstRow;
        m_lastRow   = other.m_lastRow;
        m_currentRow   = other.m_currentRow;
        m_sharedStrings = other.m_sharedStrings;
    }

    return *this;
}

/**
 * @details
 * @pre
 * @post
 */
XLRowIterator& XLRowIterator::operator=(XLRowIterator&& other) noexcept = default;

/**
 * @details
 * @pre
 * @post
 */
XLRowIterator& XLRowIterator::operator++()
{
    auto rowNumber = m_currentRow.rowNumber() + 1;
    auto rowNode = m_currentRow.m_rowNode->next_sibling();

    if (rowNumber > m_lastRow)
        m_currentRow = XLRow();

    else if (!rowNode || rowNode.attribute("r").as_ullong() != rowNumber) {
        rowNode = m_dataNode->insert_child_after("row", *m_currentRow.m_rowNode);
        rowNode.append_attribute("r").set_value(rowNumber);
        m_currentRow = XLRow(rowNode, m_sharedStrings);
    }

    else
        m_currentRow = XLRow(rowNode, m_sharedStrings);
//        throw XLException("An internal error occured");

//    auto ref = m_currentCell.cellReference();
//
//    // ===== Determine the cell reference for the next cell.
//    if (ref.column() < m_bottomRight.column())
//        ref = XLCellReference(ref.row(), ref.column() + 1);
//    else if (ref.column() == m_bottomRight.column())
//        ref = XLCellReference(ref.row() + 1, m_topLeft.column());
//
//    if (ref > m_bottomRight)
//        m_currentCell = XLCell();
//    else if (ref > m_bottomRight || ref.row() == m_currentCell.cellReference().row()) {
//        auto node = m_currentCell.m_cellNode->next_sibling();
//        if (!node || XLCellReference(node.attribute("r").value()) != ref) {
//            node = m_currentCell.m_cellNode->parent().insert_child_after("c", *m_currentCell.m_cellNode);
//            node.append_attribute("r").set_value(ref.address().c_str());
//        }
//        m_currentCell = XLCell(node, m_sharedStrings);
//    }
//    else if (ref.row() > m_currentCell.cellReference().row()) {
//        auto rowNode = m_currentCell.m_cellNode->parent().next_sibling();
//        if (!rowNode || rowNode.attribute("r").as_ullong() != ref.row()) {
//            rowNode = m_currentCell.m_cellNode->parent().parent().insert_child_after("row", m_currentCell.m_cellNode->parent());
//            rowNode.append_attribute("r").set_value(ref.row());
//            // getRowNode(*m_dataNode, ref.row());
//        }
//
//        m_currentCell = XLCell(getCellNode(rowNode, ref.column()), m_sharedStrings);
//    }
//    else
//        throw XLException("An internal error occured");

    return *this;
}

/**
 * @details
 * @pre
 * @post
 */
XLRowIterator XLRowIterator::operator++(int) //NOLINT
{
    auto oldIter(*this);
    ++(*this);
    return oldIter;
}

/**
 * @details
 * @pre
 * @post
 */
XLRow& XLRowIterator::operator*()
{
    return m_currentRow;
}

/**
 * @details
 * @pre
 * @post
 */
XLRowIterator::pointer XLRowIterator::operator->()
{
    return &m_currentRow;
}

/**
 * @details
 * @pre
 * @post
 */
bool XLRowIterator::operator==(const XLRowIterator& rhs)
{
    return m_currentRow == rhs.m_currentRow;
}

/**
 * @details
 * @pre
 * @post
 */
bool XLRowIterator::operator!=(const XLRowIterator& rhs)
{
    return !(m_currentRow == rhs.m_currentRow);
}
