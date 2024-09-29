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
XLCellIterator::XLCellIterator(const XLCellRange& cellRange, XLIteratorLocation loc, std::vector<XLStyleIndex> const * colStyles)
    : m_dataNode(std::make_unique<XMLNode>(*cellRange.m_dataNode)),
      m_topLeft(cellRange.m_topLeft),
      m_bottomRight(cellRange.m_bottomRight),
      m_sharedStrings(cellRange.m_sharedStrings),
      m_endReached(false),
      m_hintNode(),
      m_hintRow(0),
      m_currentCell(),
      m_currentCellStatus(XLNotLoaded),
      m_currentRow(0),
      m_currentColumn(0),
      m_colStyles(colStyles)
{
    if (loc == XLIteratorLocation::End)
        m_endReached = true;
    else {
        m_currentRow    = m_topLeft.row();
        m_currentColumn = m_topLeft.column();
    }
    if (m_colStyles == nullptr)
		 throw XLInternalError("XLCellIterator constructor parameter colStyles must not be nullptr");
// std::cout << "XLCellIterator constructed with topLeft " << m_topLeft.address() << " and bottomRight " << m_bottomRight.address() << std::endl;
// std::cout << "XLCellIterator m_endReached is " << ( m_endReached ? "true" : "false" ) << std::endl;
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
      m_topLeft      (other.m_topLeft),
      m_bottomRight  (other.m_bottomRight),
      m_sharedStrings(other.m_sharedStrings),
      m_endReached   (other.m_endReached),
      m_hintNode     (other.m_hintNode),
      m_hintRow      (other.m_hintRow),
      m_currentCell  (other.m_currentCell),
      m_currentCellStatus(other.m_currentCellStatus),
      m_currentRow   (other.m_currentRow),
      m_currentColumn(other.m_currentColumn),
      m_colStyles    (other.m_colStyles)
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
        m_topLeft       =  other.m_topLeft;
        m_bottomRight   =  other.m_bottomRight;
        m_sharedStrings =  other.m_sharedStrings;
        m_endReached    =  other.m_endReached;
        m_hintNode      =  other.m_hintNode;
        m_hintRow       =  other.m_currentRow;
        m_currentCell   =  other.m_currentCell;
        m_currentCellStatus = other.m_currentCellStatus;
        m_currentRow    =  other.m_currentRow;
        m_currentColumn =  other.m_currentColumn;
        m_colStyles     =  other.m_colStyles;
    }

    return *this;
}

/**
 * @details
 */
XLCellIterator& XLCellIterator::operator=(XLCellIterator&& other) noexcept = default;

namespace { // anonymous namespace for local functions findRowNode and findCellNode
    /**
     * @brief locate the XML row node within sheetDataNode for the row at rowNumber
     * @param sheetDataNode the XML sheetData node to search in
     * @param rowNumber the number of the row to locate
     * @return the XMLNode pointing to the row, or an empty XMLNode if the row does not exist
     */
    XMLNode findRowNode(XMLNode sheetDataNode, uint32_t rowNumber)
    {
        if (rowNumber < 1 || rowNumber > OpenXLSX::MAX_ROWS) {
            using namespace std::literals::string_literals;
            throw XLCellAddressError("rowNumber "s + std::to_string( rowNumber ) + " is outside valid range [1;"s + std::to_string(OpenXLSX::MAX_ROWS) + "]"s);
        }

        // ===== Get the last child of sheetDataNode that is of type node_element.
        XMLNode rowNode = sheetDataNode.last_child_of_type(pugi::node_element);

        // ===== If there are now rows in the worksheet, or the requested row is beyond the current max row, return an empty node
        if (rowNode.empty() || (rowNumber > rowNode.attribute("r").as_ullong()))
            return XMLNode{};

        // ===== If the requested node is closest to the end, start from the end and search backwards.
        if (rowNode.attribute("r").as_ullong() - rowNumber < rowNumber) {
            while (not rowNode.empty() && (rowNode.attribute("r").as_ullong() > rowNumber)) rowNode = rowNode.previous_sibling_of_type(pugi::node_element);
            if (rowNode.empty() || (rowNode.attribute("r").as_ullong() != rowNumber))
                return XMLNode{};
        }
        // ===== Otherwise, start from the beginning
        else {
            // ===== At this point, it is guaranteed that there is at least one node_element in the row that is not empty.
            rowNode = sheetDataNode.first_child_of_type(pugi::node_element);

            // ===== It has been verified above that the requested rowNumber is <= the row number of the last node_element, therefore this loop will halt.
            while (rowNode.attribute("r").as_ullong() < rowNumber) rowNode = rowNode.next_sibling_of_type(pugi::node_element);
            if (rowNode.attribute("r").as_ullong() > rowNumber)
                return XMLNode{};
        }

        return rowNode;
    }

    /**
     * @brief locate the XML cell node within rownode for the cell at columnNumber
     * @param rowNode the XML node of the row to search in
     * @param columnNumber the column number of the cell to locate
     * @return the XMLNode pointing to the cell, or an empty XMLNode if the cell does not exist
     */
    XMLNode findCellNode(XMLNode rowNode, uint16_t columnNumber)
    {
        if (columnNumber < 1 || columnNumber > OpenXLSX::MAX_COLS) {
            using namespace std::literals::string_literals;
            throw XLException("XLWorksheet::column: columnNumber "s + std::to_string(columnNumber) + " is outside allowed range [1;"s + std::to_string(MAX_COLS) + "]"s);
        }
        if (rowNode.empty()) return XMLNode{};

        XMLNode cellNode = rowNode.last_child_of_type(pugi::node_element);

        // ===== If there are no cells in the current row, or the requested cell is beyond the last cell in the row...
        if (cellNode.empty() || (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber))
            return XMLNode{};

        // ===== If the requested node is closest to the end, start from the end and search backwards...
        if (XLCellReference(cellNode.attribute("r").value()).column() - columnNumber < columnNumber) {
            while (not cellNode.empty() && (XLCellReference(cellNode.attribute("r").value()).column() > columnNumber))
                cellNode = cellNode.previous_sibling_of_type(pugi::node_element);
            if (cellNode.empty() || (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber))
                return XMLNode{};
        }
        // ===== Otherwise, start from the beginning
        else {
            // ===== At this point, it is guaranteed that there is at least one node_element in the row that is not empty.
            cellNode = rowNode.first_child_of_type(pugi::node_element);

            // ===== It has been verified above that the requested columnNumber is <= the column number of the last node_element, therefore this loop will halt:
            while (XLCellReference(cellNode.attribute("r").value()).column() < columnNumber)
                cellNode = cellNode.next_sibling_of_type(pugi::node_element);
            if (XLCellReference(cellNode.attribute("r").value()).column() > columnNumber)
                return XMLNode{};
        }
        return cellNode;
    }
}    // anonymous namespace

/**
 * @brief update m_currentCell by fetching (or inserting) a cell at m_currentRow, m_currentColumn
 */
void XLCellIterator::updateCurrentCell(bool createIfMissing)
{
    if (m_endReached)
        throw XLInputError("XLCellIterator updateCurrentCell: iterator should not be dereferenced when endReached() == true");

    if (m_currentCellStatus == XLLoaded) return;                         // nothing to do, cell is already loaded
    if (!createIfMissing && m_currentCellStatus == XLNoSuchCell) return; // nothing to do, cell has already been determined as missing

    // At this stage, m_currentCellStatus is XLUnloaded or XLNoSuchCell and createIfMissing == true

    // ===== Cell needs to be updated

    if (m_hintNode.empty()) { // no hint has been established: fetch first cell node the "tedious" way
        if (createIfMissing)      // getCellNode / getRowNode create missing cells
            m_currentCell = XLCell(getCellNode(getRowNode(*m_dataNode, m_currentRow), m_currentColumn, 0, *m_colStyles), m_sharedStrings);
        else                      // findCellNode / findRowNode return an empty cell for missing cells
            m_currentCell = XLCell(findCellNode(findRowNode(*m_dataNode, m_currentRow), m_currentColumn), m_sharedStrings);
    }
    else {
        // ===== Find or create, and fetch an XLCell at m_currentRow, m_currentColumn
        if (m_currentRow == m_hintRow) { // new cell is within the same row
            // ===== Start from m_hintNode and search forwards...
            XMLNode cellNode = m_hintNode.next_sibling_of_type(pugi::node_element);
            uint16_t colNo = 0;
            while (not cellNode.empty()) {
                colNo = XLCellReference(cellNode.attribute("r").value()).column();
                if(colNo >= m_currentColumn) break; // if desired cell was reached / passed, break before incrementing cellNode
                cellNode = cellNode.next_sibling_of_type(pugi::node_element);
            }
            if (colNo != m_currentColumn) cellNode = XMLNode{}; // if a higher column number was found, set empty node (means: "missing")
            // ===== Create missing cell node if createIfMissing == true
            if (createIfMissing && cellNode.empty()) {
                cellNode = m_hintNode.parent().insert_child_after("c", m_hintNode);
                setDefaultCellAttributes(cellNode, XLCellReference(m_currentRow, m_currentColumn).address(), m_hintNode.parent(),
                /**/                      m_currentColumn, *m_colStyles);
            }
            m_currentCell = XLCell(cellNode, m_sharedStrings); // cellNode.empty() can be true if createIfMissing == false and cell is not found
        }
        else if (m_currentRow > m_hintRow) {
            // ===== Start from m_hintNode parent row and search forwards...
            XMLNode rowNode = m_hintNode.parent().next_sibling_of_type(pugi::node_element);
            uint32_t rowNo = 0;
            while (not rowNode.empty()) {
                rowNo = rowNode.attribute("r").as_ullong();
                if (rowNo >= m_currentRow) break; // if desired row was reached / passed, break before incrementing rowNode
                rowNode = rowNode.next_sibling_of_type(pugi::node_element);
            }
            if (rowNo != m_currentRow) rowNode = XMLNode{}; // if a higher row number was found, set empty node (means: "missing")
            // ===== Create missing row node if createIfMissing == true
            if (createIfMissing && rowNode.empty()) {
                rowNode = m_dataNode->insert_child_after("row", m_hintNode.parent());
                rowNode.append_attribute("r").set_value(m_currentRow);
            }
            if (rowNode.empty())    // if row could not be found / created
                m_currentCell = XLCell{}; // make sure m_currentCell is set to an empty cell
            else {                  // else: row found
                if (createIfMissing) {
                    // ===== Pass the already known m_currentRow to getCellNode so that it does not have to be fetched again
                    m_currentCell = XLCell(getCellNode (rowNode, m_currentColumn, m_currentRow, *m_colStyles), m_sharedStrings);
                }
                else // ===== Do a "soft find" if a missing cell shall not be created
                    m_currentCell = XLCell(findCellNode(rowNode, m_currentColumn), m_sharedStrings);
            }
        }
        else
            throw XLInternalError("XLCellIterator::updateCurrentCell: an internal error occured (m_currentRow < m_hintRow)");
    }

    if (m_currentCell.empty())    // if cell is confirmed missing
        m_currentCellStatus = XLNoSuchCell; // mark this status for further calls to updateCurrentCell()
    else {
        // ===== If the current cell exists, update the hints
        m_hintNode   = *m_currentCell.m_cellNode;    // 2024-08-11: don't store a full XLCell, just the XMLNode, for better performance
        m_hintRow    = m_currentRow;
        m_currentCellStatus = XLLoaded; // mark cell status for further calls to updateCurrentCell()
    }
}

/**
 * @details
 */
XLCellIterator& XLCellIterator::operator++()
{
// std::cout << "XLCellIterator operator++, m_currentColumn " << m_currentColumn << " and m_currentRow " << m_currentRow << std::endl;
    if (m_endReached)
        throw XLInputError("XLCellIterator: tried to increment beyond end operator");

    if (m_currentColumn < m_bottomRight.column())
        ++m_currentColumn;
    else if(m_currentRow < m_bottomRight.row()) {
        ++m_currentRow;
        m_currentColumn = m_topLeft.column();
    }
    else
        m_endReached = true;

    m_currentCellStatus = XLNotLoaded; // trigger a new attempt to locate / create the cell via updateCurrentCell

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
XLCell& XLCellIterator::operator*()
{
// std::cout << "XLCellIterator dereference operator* invoked" << std::endl;
    updateCurrentCell(XLCreateIfMissing);
    return m_currentCell;
}

/**
 * @details
 */
XLCellIterator::pointer XLCellIterator::operator->()
{
// std::cout << "XLCellIterator dereference operator-> invoked" << std::endl;
    updateCurrentCell(XLCreateIfMissing);
    return &m_currentCell;
}

/**
 * @details
 */
bool XLCellIterator::operator==(const XLCellIterator& rhs) const
{
    // BUGFIX 2024-08-10: there was no test for (!m_currentCell && rhs.m_currentCell),
    //     leading to a potential dereference of a nullptr in m_currentCell::m_cellNode

    if (m_endReached && rhs.m_endReached) return true;    // If both iterators are end iterators

    if ((m_currentColumn != rhs.m_currentColumn)          // if iterators point to a different column or row
      ||(m_currentRow    != rhs.m_currentRow))
        return false;                                         // that means no match

    // CAUTION: for performance reasons, disabled all checks whether this and rhs are iterators on the same worksheet & range
    return true;

    // if (*m_dataNode != *rhs.m_dataNode) return false;     // TBD: iterators over different worksheets may never match
    // TBD if iterators shall be considered not equal if they were created on different XLCellRanges
    // this would require checking the topLeft and bottomRight references, potentially costing CPU time

    // return m_currentCell == rhs.m_currentCell;   // match only if cell nodes are equal
    // CAUTION: in the current code, that means iterators that point to the same column & row in different worksheets,
    // and cells that do not exist in both sheets, will be considered equal
}

/**
 * @details
 */
bool XLCellIterator::operator!=(const XLCellIterator& rhs) const { return !(*this == rhs); }

/**
 * @details
 */
bool XLCellIterator::cellExists()
{
    // ===== Update m_currentCell once so that cellExists will always test the correct cell (an empty cell if current cell doesn't exist)
    updateCurrentCell(XLDoNotCreateIfMissing);
    return not m_currentCell.empty();
}

/**
 * @details
 * @note 2024-06-03: implemented a calculated distance based on m_currentCell, m_topLeft and m_bottomRight (if m_endReached)
 *                   accordingly, implemented defined setting of m_endReached at all times
 */
uint64_t XLCellIterator::distance(const XLCellIterator& last)
{
    // ===== Determine rows and columns, taking into account beyond-the-end iterators
    uint32_t row = (m_endReached ? m_bottomRight.row() : m_currentRow);
    uint16_t col = (m_endReached ? m_bottomRight.column() + 1 : m_currentColumn);
    uint32_t lastRow = (last.m_endReached ? last.m_bottomRight.row() : last.m_currentRow);
    // ===== lastCol can store +1 for beyond-the-end iterator without overflow because MAX_COLS is less than max uint16_t
    uint16_t lastCol = (last.m_endReached ? last.m_bottomRight.column() + 1 : last.m_currentColumn);

    uint16_t rowWidth = m_bottomRight.column() - m_topLeft.column() + 1;    // amount of cells in a row of the iterator range
    int64_t distance =  ((int64_t)(lastRow) - row) * rowWidth    //   row distance * rowWidth
                       + (int64_t)(lastCol) - col;               // + column distance (may be negative)
    if (distance < 0)
        throw XLInputError("XLCellIterator::distance is negative");

    return static_cast<uint64_t>(distance);    // after excluding negative result: cast back to positive value
}

/**
 * @details
 */
std::string XLCellIterator::address() const
{
    uint32_t row = (m_endReached ? m_bottomRight.row() : m_currentRow);
    uint16_t col = (m_endReached ? m_bottomRight.column() + 1 : m_currentColumn);
    return (m_endReached ? "END(" : "") + XLCellReference(row, col).address() + (m_endReached ? ")" : "");
}
