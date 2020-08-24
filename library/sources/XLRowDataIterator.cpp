//
// Created by Kenneth Balslev on 24/08/2020.
//

#include <pugixml.hpp>

#include "XLCellReference.hpp"
#include "XLRowDataIterator.hpp"
#include "XLRowDataRange.hpp"

using namespace OpenXLSX;

namespace OpenXLSX
{
    XMLNode getCellNode(XMLNode rowNode, uint16_t columnNumber);
}

/**
 * @details
 * @pre
 * @post
 */
XLRowDataIterator::XLRowDataIterator(const XLRowDataRange& rowDataRange, XLIteratorLocation loc)
    : m_rowNode(std::make_unique<XMLNode>(*rowDataRange.m_rowNode)),
      m_firstCol(rowDataRange.m_firstCol),
      m_lastCol(rowDataRange.m_lastCol),
      m_currentCell(),
      m_sharedStrings(rowDataRange.m_sharedStrings)
{
    if (loc == XLIteratorLocation::End)
        m_currentCell = XLCell();
    else {
        m_currentCell = XLCell(getCellNode(*m_rowNode, m_firstCol), m_sharedStrings);
    }
}

/**
 * @details
 * @pre
 * @post
 */
XLRowDataIterator::~XLRowDataIterator() = default;

/**
 * @details
 * @pre
 * @post
 */
XLRowDataIterator::XLRowDataIterator(const XLRowDataIterator& other)
    : m_rowNode(std::make_unique<XMLNode>(*other.m_rowNode)),
      m_firstCol(other.m_firstCol),
      m_lastCol(other.m_lastCol),
      m_currentCell(other.m_currentCell),
      m_sharedStrings(other.m_sharedStrings)
{}

/**
 * @details
 * @pre
 * @post
 */
XLRowDataIterator::XLRowDataIterator(XLRowDataIterator&& other) noexcept = default;

/**
 * @details
 * @pre
 * @post
 */
XLRowDataIterator& XLRowDataIterator::operator=(const XLRowDataIterator& other)
{
    if (&other != this) {
        *m_rowNode     = *other.m_rowNode;
        m_firstCol      = other.m_firstCol;
        m_lastCol       = other.m_lastCol;
        m_currentCell    = other.m_currentCell;
        m_sharedStrings = other.m_sharedStrings;
    }

    return *this;
}

/**
 * @details
 * @pre
 * @post
 */
XLRowDataIterator& XLRowDataIterator::operator=(XLRowDataIterator&& other) noexcept = default;

/**
 * @details
 * @pre
 * @post
 */
XLRowDataIterator& XLRowDataIterator::operator++()
{
    auto cellNumber = m_currentCell.cellReference().column() + 1;
    auto cellNode = m_currentCell.m_cellNode->next_sibling();

    if (cellNumber > m_lastCol)
        m_currentCell = XLCell();

    else if (!cellNode || XLCellReference(cellNode.attribute("r").value()).column() != cellNumber) {
        cellNode = m_rowNode->insert_child_after("c", *m_currentCell.m_cellNode);
        cellNode.append_attribute("r").set_value(XLCellReference(m_rowNode->attribute("r").as_ullong(), cellNumber).address().c_str());
        m_currentCell = XLCell(cellNode, m_sharedStrings);
    }

    else
        m_currentCell = XLCell(cellNode, m_sharedStrings);

    return *this;
}

/**
 * @details
 * @pre
 * @post
 */
XLRowDataIterator XLRowDataIterator::operator++(int)
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
XLCell& XLRowDataIterator::operator*()
{
    return m_currentCell;
}

/**
 * @details
 * @pre
 * @post
 */
XLRowDataIterator::pointer XLRowDataIterator::operator->()
{
    return &m_currentCell;
}

/**
 * @details
 * @pre
 * @post
 */
bool XLRowDataIterator::operator==(const XLRowDataIterator& rhs)
{
    return m_currentCell == rhs.m_currentCell;
}

/**
 * @details
 * @pre
 * @post
 */
bool XLRowDataIterator::operator!=(const XLRowDataIterator& rhs)
{
    return !(m_currentCell == rhs.m_currentCell);
}
