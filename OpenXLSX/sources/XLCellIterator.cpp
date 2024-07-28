/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.
 8b    d8   MM.  ,M9 YM    d9 MM    MM  d'    `MM.   MM    / L    ,M9  d'    `MM.
  YMMMM9    MMYMMM9   YMMMM9 _MM_  _MM_M(_    _)MM_ _MMMMMMM MYMMMM9 _M(_    _)MM_
            MM
            MM
           _MM_

  Copyright (c) 2018, Kenneth Troldal Balslev

  All rights reserved.

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  - Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.
  - Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.
  - Neither the name of the author nor the
    names of any contributors may be used to endorse or promote products
    derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

// ===== External Includes ===== //
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCellIterator.hpp"
#include "XLCellRange.hpp"
#include "XLCellReference.hpp"
#include "XLException.hpp"
#include "utilities/XLUtilities.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
XLCellIterator::XLCellIterator(const XLCellRange& cellRange, XLIteratorLocation loc)
    : m_dataNode(std::make_unique<XMLNode>(*cellRange.m_dataNode)),
      m_topLeft(cellRange.m_topLeft),
      m_bottomRight(cellRange.m_bottomRight),
      m_sharedStrings(cellRange.m_sharedStrings),
      m_endReached(false)
{
    if (loc == XLIteratorLocation::End) {
        m_currentCell = XLCell();
        m_endReached = true;
    }
    else {
        m_currentCell = XLCell(getCellNode(getRowNode(*m_dataNode, m_topLeft.row()), m_topLeft.column()), m_sharedStrings);
    }
}

/**
 * @details
 */
XLCellIterator::~XLCellIterator() = default;

/**
 * @details
 */
XLCellIterator::XLCellIterator(const XLCellIterator& other)
    : m_dataNode(std::make_unique<XMLNode>(*other.m_dataNode)),
      m_topLeft(other.m_topLeft),
      m_bottomRight(other.m_bottomRight),
      m_currentCell(other.m_currentCell),
      m_sharedStrings(other.m_sharedStrings),
      m_endReached(other.m_endReached)
{}

/**
 * @details
 */
XLCellIterator::XLCellIterator(XLCellIterator&& other) noexcept = default;

/**
 * @details
 */
XLCellIterator& XLCellIterator::operator=(const XLCellIterator& other)
{
    if (&other != this) {
        *m_dataNode     = *other.m_dataNode;
        m_topLeft       = other.m_topLeft;
        m_bottomRight   = other.m_bottomRight;
        m_currentCell   = other.m_currentCell;
        m_sharedStrings = other.m_sharedStrings;
        m_endReached    = other.m_endReached;
    }

    return *this;
}

/**
 * @details
 */
XLCellIterator& XLCellIterator::operator=(XLCellIterator&& other) noexcept = default;

/**
 * @details
 */
XLCellIterator& XLCellIterator::operator++()
{
    if (m_endReached)
        throw XLInputError("XLCellIterator: tried to increment beyond end operator");

    auto ref = m_currentCell.cellReference();

    // ===== Determine the cell reference for the next cell.
    if (ref.column() < m_bottomRight.column())
        ref = XLCellReference(ref.row(), ref.column() + 1);
    else if (ref == m_bottomRight)
        m_endReached = true;
    else
        ref = XLCellReference(ref.row() + 1, m_topLeft.column());

    // 2024-06-03 TBD TODO: why ref > m_bottomRight - that shouldn't be possible? --> added exception to test for this
    if (ref > m_bottomRight)
        throw XLInternalError("XLCellIterator became > m_bottomRight - this should not happen!");

    if (m_endReached)
        m_currentCell = XLCell();
    else if (ref > m_bottomRight || ref.row() == m_currentCell.cellReference().row()) {    // TBD: remove ref > m_bottomRight condition unless I overlooked something
        auto node = m_currentCell.m_cellNode->next_sibling_of_type(pugi::node_element);
        if (node.empty() || XLCellReference(node.attribute("r").value()) != ref) {
            node = m_currentCell.m_cellNode->parent().insert_child_after("c", *m_currentCell.m_cellNode);
            node.append_attribute("r").set_value(ref.address().c_str());
        }
        m_currentCell = XLCell(node, m_sharedStrings);
    }
    else if (ref.row() > m_currentCell.cellReference().row()) {
        auto rowNode = m_currentCell.m_cellNode->parent().next_sibling_of_type(pugi::node_element);
        if (rowNode.empty() || rowNode.attribute("r").as_ullong() != ref.row()) {
            rowNode = m_currentCell.m_cellNode->parent().parent().insert_child_after("row", m_currentCell.m_cellNode->parent());
            rowNode.append_attribute("r").set_value(ref.row());
        }
        // ===== Pass the already known ref.row() to getCellNode so that it does not have to be fetched again
        m_currentCell = XLCell(getCellNode(rowNode, ref.column(), ref.row()), m_sharedStrings);
    }
    else
        throw XLInternalError("An internal error occured");

    return *this;
}

/**
 * @details
 */
XLCellIterator XLCellIterator::operator++(int)    // NOLINT
{
    auto oldIter(*this);
    ++(*this);
    return oldIter;
}

/**
 * @details
 */
XLCell& XLCellIterator::operator*() { return m_currentCell; }

/**
 * @details
 */
XLCellIterator::pointer XLCellIterator::operator->() { return &m_currentCell; }

/**
 * @details
 */
bool XLCellIterator::operator==(const XLCellIterator& rhs) const
{
    if (m_currentCell && !rhs.m_currentCell) return false;
    if (!m_currentCell && !rhs.m_currentCell) return true;
    return m_currentCell == rhs.m_currentCell;
}

/**
 * @details
 */
bool XLCellIterator::operator!=(const XLCellIterator& rhs) const { return !(*this == rhs); }

/**
 * @details
 * @note 2024-06-03: implemented a calculated distance based on m_currentCell, m_topLeft and m_bottomRight (if m_endReached)
 *                   accordingly, implemented defined setting of m_endReached at all times
 */
uint64_t XLCellIterator::distance(const XLCellIterator& last)
{
    // ===== Determine rows and columns, taking into account beyond-the-end iterators
    uint32_t row = (m_endReached ? m_bottomRight.row() : m_currentCell.cellReference().row());
    uint16_t col = (m_endReached ? m_bottomRight.column() + 1 : m_currentCell.cellReference().column());
    uint32_t lastRow = (last.m_endReached ? last.m_bottomRight.row() : last.m_currentCell.cellReference().row());
    // ===== lastCol can store +1 for beyond-the-end iterator without overflow because MAX_COLS is less than max uint16_t
    uint16_t lastCol = (last.m_endReached ? last.m_bottomRight.column() + 1 : last.m_currentCell.cellReference().column());

    uint16_t rowWidth = m_bottomRight.column() - m_topLeft.column() + 1;    // amount of cells in a row of the iterator range
    int64_t distance =  ((int64_t)(lastRow) - row) * rowWidth    //   row distance * rowWidth
                       + (int64_t)(lastCol) - col;               // + column distance (may be negative)
    if (distance < 0)
        throw XLInputError("XLCellIterator::distance is negative");

    return static_cast<uint64_t>(distance);    // after excluding negative result: cast back to positive value

    /* OBSOLETE CODE:
    // uint64_t result = 0;
    // while (*this != last) {
    //     ++result;
    //     ++(*this);
    // }
    // return result;
    */
}

/**
 * @details
 */
std::string XLCellIterator::address() const
{
    uint32_t row = (m_endReached ? m_bottomRight.row() : m_currentCell.cellReference().row());
    uint16_t col = (m_endReached ? m_bottomRight.column() + 1 : m_currentCell.cellReference().column());
    return (m_endReached ? "END(" : "") + XLCellReference(row, col).address() + (m_endReached ? ")" : "");
}
