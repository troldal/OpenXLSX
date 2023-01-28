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

using namespace OpenXLSX;

namespace OpenXLSX
{
    XMLNode getRowNode(XMLNode sheetDataNode, uint32_t rowNumber);
    XMLNode getCellNode(XMLNode rowNode, uint16_t columnNumber);
}    // namespace OpenXLSX

/**
 * @details
 */
XLCellIterator::XLCellIterator(const XLCellRange& cellRange, XLIteratorLocation loc)
    : m_dataNode(std::make_unique<XMLNode>(*cellRange.m_dataNode)),
      m_topLeft(cellRange.m_topLeft),
      m_bottomRight(cellRange.m_bottomRight),
      m_worksheet(cellRange.m_worksheet)
{
    if (loc == XLIteratorLocation::End)
        m_currentCell = XLCell();
    else {
        m_currentCell = XLCell(getCellNode(getRowNode(*m_dataNode, m_topLeft.row()), 
                                                m_topLeft.column()), m_worksheet);
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
      m_worksheet(other.m_worksheet)
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
        m_worksheet     = other.m_worksheet;
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
    auto ref = m_currentCell.cellReference();

    // ===== Determine the cell reference for the next cell.
    if (ref.column() < m_bottomRight.column())
        ref = XLCellReference(ref.row(), ref.column() + 1);
    else if (ref == m_bottomRight)
        m_endReached = true;
    else
        ref = XLCellReference(ref.row() + 1, m_topLeft.column());

    if (m_endReached)
        m_currentCell = XLCell();
    else if (ref > m_bottomRight || ref.row() == m_currentCell.cellReference().row()) {
        auto node = m_currentCell.m_cellNode->next_sibling();
        if (!node || XLCellReference(node.attribute("r").value()) != ref) {
            node = m_currentCell.m_cellNode->parent().insert_child_after("c", *m_currentCell.m_cellNode);
            node.append_attribute("r").set_value(ref.address().c_str());
        }
        m_currentCell = XLCell(node, m_worksheet);
    }
    else if (ref.row() > m_currentCell.cellReference().row()) {
        auto rowNode = m_currentCell.m_cellNode->parent().next_sibling();
        if (!rowNode || rowNode.attribute("r").as_ullong() != ref.row()) {
            rowNode = m_currentCell.m_cellNode->parent().parent().insert_child_after("row", m_currentCell.m_cellNode->parent());
            rowNode.append_attribute("r").set_value(ref.row());
            // getRowNode(*m_dataNode, ref.row());
        }

        m_currentCell = XLCell(getCellNode(rowNode, ref.column()), m_worksheet);
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
XLCell& XLCellIterator::operator*()
{
    return m_currentCell;
}

/**
 * @details
 */
XLCellIterator::pointer XLCellIterator::operator->()
{
    return &m_currentCell;
}

/**
 * @details
 */
bool XLCellIterator::operator==(const XLCellIterator& rhs) const
{
    if (m_currentCell && !rhs.m_currentCell)
        return false;
    if (!m_currentCell && !rhs.m_currentCell)
        return true;
    return m_currentCell == rhs.m_currentCell;
}

/**
 * @details
 */
bool XLCellIterator::operator!=(const XLCellIterator& rhs) const
{
    return !(*this == rhs);
}

/**
 * @details
 * @todo This implementation is rather ineffecient. Consider an alternative implementation.
 */
uint64_t XLCellIterator::distance(const XLCellIterator& last)
{
    uint64_t result = 0;
    while (*this != last) {
        ++result;
        ++(*this);
    }
    return result;
}
