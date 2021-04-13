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
#include <algorithm>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCell.hpp"
#include "XLRow.hpp"
#include "XLRowData.hpp"
#include "utilities/XLUtilities.hpp"

// ========== XLRowDataIterator  ============================================ //
namespace OpenXLSX
{
    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator::XLRowDataIterator(const XLRowDataRange& rowDataRange, XLIteratorLocation loc)
        : m_dataRange(std::make_unique<XLRowDataRange>(rowDataRange)),
          m_currentCell(loc == XLIteratorLocation::End
                            ? XLCell()
                            : XLCell(getCellNode(*m_dataRange->m_rowNode, m_dataRange->m_firstCol), m_dataRange->m_sharedStrings)),
          m_currentCol(loc == XLIteratorLocation::End ? m_dataRange->m_lastCol : m_dataRange->m_firstCol),
          m_cellNode(std::make_unique<XMLNode>(nullptr))
    {}

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
        : m_dataRange(std::make_unique<XLRowDataRange>(*other.m_dataRange)),
          m_currentCell(other.m_currentCell),
          m_currentCol(other.m_currentCol),
          m_cellNode(std::make_unique<XMLNode>(*other.m_cellNode))
    {}

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator::XLRowDataIterator(XLRowDataIterator&& other) noexcept = default;    // NOLINT

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator& XLRowDataIterator::operator=(const XLRowDataIterator& other)
    {
        if (&other != this) {
            *m_dataRange  = *other.m_dataRange;
            m_currentCell = other.m_currentCell;
            m_currentCol  = other.m_currentCol;
            *m_cellNode   = *other.m_cellNode;
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
        ++m_currentCol;

        // ===== Compute cellreference
        auto cellRef = XLCellReference(m_dataRange->m_rowNode->attribute("r").as_ullong(), m_currentCol);

        // ===== If the current m_cellNode reference is empty, find the next cell node in the row that is greater than
        // ===== or equal to the cell refecence.
        if (!m_cellNode) {
            *m_cellNode = m_dataRange->m_rowNode->find_child(
                [&](const XMLNode& node) { return XLCellReference(node.attribute("r").value()) >= cellRef; });

            // TODO: Check the logic here!
            if (*m_cellNode && XLCellReference(m_cellNode->attribute("r").value()) > cellRef)
                *m_cellNode = m_cellNode->previous_sibling();
        }


        else {
            // TODO: Check the logic here!
            if (XLCellReference(m_cellNode->next_sibling().attribute("r").value()) == cellRef)
                *m_cellNode = m_cellNode->next_sibling();
        }

        //        auto cellNumber = m_currentCell.cellReference().column() + 1;
        //        auto cellNode   = m_currentCell.m_cellNode->next_sibling();
        //
        //        if (cellNumber > m_dataRange->m_lastCol)
        //            m_currentCell = XLCell();
        //
        //        else if (!cellNode || XLCellReference(cellNode.attribute("r").value()).column() != cellNumber) {
        //            cellNode = m_dataRange->m_rowNode->insert_child_after("c", *m_currentCell.m_cellNode);
        //            cellNode.append_attribute("r").set_value(
        //                XLCellReference(m_dataRange->m_rowNode->attribute("r").as_ullong(), cellNumber).address().c_str());
        //            m_currentCell = XLCell(cellNode, m_dataRange->m_sharedStrings);
        //        }
        //
        //        else
        //            m_currentCell = XLCell(cellNode, m_dataRange->m_sharedStrings);

        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator XLRowDataIterator::operator++(int)    // NOLINT
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
        if (m_currentCol >= m_dataRange->m_lastCol)
            m_currentCell = XLCell();
//        else {
//            if (!m_cellNode) }

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

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator::operator bool() const
    {
        return false;
    }

}    // namespace OpenXLSX

// ========== XLRowDataRange  =============================================== //
namespace OpenXLSX
{
    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataRange::XLRowDataRange(const XMLNode& rowNode, uint16_t firstColumn, uint16_t lastColumn, XLSharedStrings* sharedStrings)
        : m_rowNode(std::make_unique<XMLNode>(rowNode)),
          m_firstCol(firstColumn),
          m_lastCol(lastColumn),
          m_sharedStrings(sharedStrings)
    {}

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataRange::XLRowDataRange(const XLRowDataRange& other)
        : m_rowNode(std::make_unique<XMLNode>(*other.m_rowNode)),
          m_firstCol(other.m_firstCol),
          m_lastCol(other.m_lastCol),
          m_sharedStrings(other.m_sharedStrings)

    {}

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataRange::XLRowDataRange(XLRowDataRange&& other) noexcept = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataRange::~XLRowDataRange() = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataRange& XLRowDataRange::operator=(const XLRowDataRange& other)
    {
        if (&other != this) {
            *m_rowNode      = *other.m_rowNode;
            m_firstCol      = other.m_firstCol;
            m_lastCol       = other.m_lastCol;
            m_sharedStrings = other.m_sharedStrings;
        }

        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataRange& XLRowDataRange::operator=(XLRowDataRange&& other) noexcept
    {
        if (&other != this) {
            *m_rowNode      = *other.m_rowNode;
            m_firstCol      = other.m_firstCol;
            m_lastCol       = other.m_lastCol;
            m_sharedStrings = other.m_sharedStrings;
        }

        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    uint16_t XLRowDataRange::cellCount() const
    {
        return m_lastCol - m_firstCol + 1;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator XLRowDataRange::begin()
    {
        return XLRowDataIterator(*this, XLIteratorLocation::Begin);
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator XLRowDataRange::end()
    {
        return XLRowDataIterator(*this, XLIteratorLocation::End);
    }

    /**
     * @details
     * @pre
     * @post
     * @todo To be implemented
     */
    void XLRowDataRange::clear() {}
}    // namespace OpenXLSX

// ========== XLRowDataProxy  =============================================== //
namespace OpenXLSX
{
    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy::~XLRowDataProxy() = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy& XLRowDataProxy::operator=(const XLRowDataProxy& other)
    {
        if (&other != this) {
            *this = other.getValues();
        }

        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy::XLRowDataProxy(XLRow* row, XMLNode* rowNode) : m_row(row), m_rowNode(rowNode) {}

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy::XLRowDataProxy(const XLRowDataProxy& other) = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy::XLRowDataProxy(XLRowDataProxy&& other) noexcept = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy& XLRowDataProxy::operator=(XLRowDataProxy&& other) noexcept = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy& XLRowDataProxy::operator=(const std::vector<XLCellValue>& values)
    {
        // ===== Find or create the first cell node
        auto curNode = m_rowNode->first_child();
        if (!curNode || XLCellReference(curNode.attribute("r").value()).column() != 1) {
            curNode = m_rowNode->append_child("c"); // TODO: Should it be prepended?
            curNode.append_attribute("r").set_value(XLCellReference(m_row->rowNumber(), 1).address().c_str());
        }

        auto     prevNode = XMLNode();
        uint16_t col      = 1;
        for (const auto& value : values) {
            if (!curNode || XLCellReference(curNode.attribute("r").value()).column() != col) {
                curNode = m_rowNode->insert_child_after("c", prevNode);
                curNode.append_attribute("r").set_value(XLCellReference(m_row->rowNumber(), col).address().c_str());
            }

            XLCell(curNode, m_row->m_sharedStrings).value() = value;

            prevNode = curNode;
            curNode  = curNode.next_sibling();
            ++col;
        }

        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy::operator std::vector<XLCellValue>()
    {
        return getValues();
    }

    /**
     * @details
     * @pre
     * @post
     */
    std::vector<XLCellValue> XLRowDataProxy::getValues() const
    {
        auto                     numCells = XLCellReference(m_rowNode->last_child().attribute("r").value()).column();
        std::vector<XLCellValue> result(numCells);

        for (auto& node : m_rowNode->children()) {
            if (XLCellReference(node.attribute("r").value()).column() > numCells) break;
            result[XLCellReference(node.attribute("r").value()).column() - 1] = XLCell(node, m_row->m_sharedStrings).value();
        }

        return result;
    }

}    // namespace OpenXLSX