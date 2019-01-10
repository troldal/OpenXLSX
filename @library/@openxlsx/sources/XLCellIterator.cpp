//
// Created by KBA012 on 22/09/2016.
//


#include "XLCellIterator.h"
#include "XLCellRange.h"

using namespace OpenXLSX::Impl;

/*
 * =============================================================================================================
 * XLCellIterator::XLCellIterator
 * =============================================================================================================
 */

/**
 * @details Construction of an iterator from a cell range. The iterator will point to the top left cell in the grid.
 */
XLCellIterator::XLCellIterator(XLCellRange &range,
                               XLCell *cell)
    : m_cell(cell),
      m_range(&range),
      m_currentRow(1),
      m_currentColumn(1),
      m_numRows(1),
      m_numColumns(1)
{
    m_currentRow = 1;
    m_currentColumn = 1;

    m_numRows = m_range->NumRows();
    m_numColumns = m_range->NumColumns();
}

/**
 * @details The increment method will increment the iterator, first through each columns in a row, then over each row.
 */
const XLCellIterator &XLCellIterator::operator++()
{

    if (m_currentColumn < m_numColumns)
        ++m_currentColumn;
    else if (m_currentColumn == m_numColumns && m_currentRow < m_numRows) {
        m_currentColumn = 1;
        ++m_currentRow;
    }
    else {
        m_cell = nullptr;
        return *this;
    }

    m_cell = m_range->Cell(m_currentRow, m_currentColumn);

    return *this;
}

/**
 * @details Checks whether two iterator objects are pointing to the same object.
 */
bool XLCellIterator::operator!=(const XLCellIterator &other) const
{
    return (this->m_cell != other.m_cell);
}

/**
 * @details Dereferences the XLCell object to get the value.
 */
XLCell &XLCellIterator::operator*() const
{
    return *m_cell;
}

/**
 * @details
 */
XLCellIteratorConst::XLCellIteratorConst(const XLCellRange &range,
                                         const XLCell *cell)
    : m_cell(cell),
      m_range(&range),
      m_currentRow(1),
      m_currentColumn(1),
      m_numRows(1),
      m_numColumns(1)
{
    m_currentRow = 1;
    m_currentColumn = 1;

    m_numRows = m_range->NumRows();
    m_numColumns = m_range->NumColumns();
}

/**
 * @details
 */
const XLCellIteratorConst &XLCellIteratorConst::operator++()
{

    if (m_currentColumn < m_numColumns)
        ++m_currentColumn;
    else if (m_currentColumn == m_numColumns && m_currentRow < m_numRows) {
        m_currentColumn = 1;
        ++m_currentRow;
    }
    else {
        m_cell = nullptr;
        return *this;
    }

    m_cell = m_range->Cell(m_currentRow, m_currentColumn);

    return *this;
}

/**
 * @details
 */
bool XLCellIteratorConst::operator!=(const XLCellIteratorConst &other) const
{
    return (this->m_cell != other.m_cell);
}

/**
 * @details
 */
const XLCell &XLCellIteratorConst::operator*() const
{
    return *m_cell;
}
