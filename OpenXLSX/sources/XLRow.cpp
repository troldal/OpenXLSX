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
#include "XLCell.hpp"
#include "XLCellReference.hpp"
#include "XLRow.hpp"

#include "utilities/XLUtilities.hpp"

// ========== XLRow  ======================================================== //
namespace OpenXLSX
{
    /**
     * @details
     * @pre
     * @post
     */
    XLRow::XLRow() : m_rowNode(nullptr),
                     m_rowDataProxy(this, m_rowNode.get()) {}

    /**
     * @details Constructs a new XLRow object from information in the underlying XML file. A pointer to the corresponding
     * node in the underlying XML file must be provided.
     * @pre
     * @post
     */
    XLRow::XLRow(const XMLNode& rowNode, const XLSharedStrings& sharedStrings)
        : m_rowNode(std::make_unique<XMLNode>(rowNode)),
          m_sharedStrings(sharedStrings),
          m_rowDataProxy(this, m_rowNode.get())
    {}

    /**
     * @details
     * @pre
     * @post
     */
    XLRow::XLRow(const XLRow& other)
        : m_rowNode(other.m_rowNode ? std::make_unique<XMLNode>(*other.m_rowNode) : nullptr),
          m_sharedStrings(other.m_sharedStrings),
          m_rowDataProxy(this, m_rowNode.get())
    {}

    /**
     * @details Because the m_rowDataProxy variable is tied to an exact XLRow object, the move operation is
     * not a 'pure' move, as a new XLRowDataProxy has to be constructed.
     * @pre
     * @post
     */
    XLRow::XLRow(XLRow&& other) noexcept
        : m_rowNode(std::move(other.m_rowNode)),
          m_sharedStrings(std::move(other.m_sharedStrings)),
          m_rowDataProxy(this, m_rowNode.get())
    {}

    /**
     * @details
     * @pre
     * @post
     */
    XLRow::~XLRow() = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRow& XLRow::operator=(const XLRow& other)
    {
        if (&other != this) {
            auto temp = XLRow(other);
            std::swap(*this, temp);
        }
        return *this;
    }

    /**
     * @details Because the m_rowDataProxy variable is tied to an exact XLRow object, the move operation is
     * not a 'pure' move, as a new XLRowDataProxy has to be constructed.
     * @pre
     * @post
     */
    XLRow& XLRow::operator=(XLRow&& other) noexcept
    {
        if (&other != this) {
            m_rowNode       = std::move(other.m_rowNode);
            m_sharedStrings = other.m_sharedStrings;
            m_rowDataProxy  = XLRowDataProxy(this, m_rowNode.get());
        }
        return *this;
    }

    /**
     * @details Returns the m_height member by getValue.
     * @pre
     * @post
     */
    double XLRow::height() const
    {
        return m_rowNode->attribute("ht").as_double(15.0); // NOLINT
    }

    /**
     * @details Set the height of the row. This is done by setting the getValue of the 'ht' attribute and setting the
     * 'customHeight' attribute to true.
     * @pre
     * @post
     */
    void XLRow::setHeight(float height)
    {
        // Set the 'ht' attribute for the Cell. If it does not exist, create it.
        if (!m_rowNode->attribute("ht"))
            m_rowNode->append_attribute("ht") = height;
        else
            m_rowNode->attribute("ht").set_value(height);

        // Set the 'customHeight' attribute. If it does not exist, create it.
        if (!m_rowNode->attribute("customHeight"))
            m_rowNode->append_attribute("customHeight") = 1;
        else
            m_rowNode->attribute("customHeight").set_value(1);
    }

    /**
     * @details Return the m_descent member by getValue.
     * @pre
     * @post
     */
    float XLRow::descent() const
    {
        return m_rowNode->attribute("x14ac:dyDescent").as_float(0.25); // NOLINT
    }

    /**
     * @details Set the descent by setting the 'x14ac:dyDescent' attribute in the XML file
     * @pre
     * @post
     */
    void XLRow::setDescent(float descent)
    {
        // Set the 'x14ac:dyDescent' attribute. If it does not exist, create it.
        if (!m_rowNode->attribute("x14ac:dyDescent"))
            m_rowNode->append_attribute("x14ac:dyDescent") = descent;
        else
            m_rowNode->attribute("x14ac:dyDescent") = descent;
    }

    /**
     * @details Determine if the row is hidden or not.
     * @pre
     * @post
     */
    bool XLRow::isHidden() const
    {
        return m_rowNode->attribute("hidden").as_bool(false);
    }

    /**
     * @details Set the hidden state by setting the 'hidden' attribute to true or false.
     * @pre
     * @post
     */
    void XLRow::setHidden(bool state)
    {
        // Set the 'hidden' attribute. If it does not exist, create it.
        if (!m_rowNode->attribute("hidden"))
            m_rowNode->append_attribute("hidden") = static_cast<int>(state);
        else
            m_rowNode->attribute("hidden").set_value(static_cast<int>(state));
    }

    /**
     * @details
     * @pre
     * @post
     */
    uint64_t XLRow::rowNumber() const
    {
        return m_rowNode->attribute("r").as_ullong();
    }

    /**
     * @details Get the number of cells in the row, by returning the size of the m_cells vector.
     * @pre
     * @post
     */
    unsigned int XLRow::cellCount() const
    {
        if (!m_rowNode->last_child())
            return 0;
        return XLCellReference(m_rowNode->last_child().attribute("r").value()).column();
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy& XLRow::values()
    {
        return m_rowDataProxy;
    }

    /**
     * @details
     * @pre
     * @post
     */
    const XLRowDataProxy& XLRow::values() const
    {
        return m_rowDataProxy;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataRange XLRow::cells() const
    {
        return XLRowDataRange(*m_rowNode, 1, XLCellReference(m_rowNode->last_child().attribute("r").value()).column(), m_sharedStrings);
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataRange XLRow::cells(uint16_t cellCount) const
    {
        return XLRowDataRange(*m_rowNode, 1, cellCount, m_sharedStrings);
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataRange XLRow::cells(uint16_t firstCell, uint16_t lastCell) const
    {
        return XLRowDataRange(*m_rowNode, firstCell, lastCell, m_sharedStrings);
    }

    bool XLRow::isEqual(const XLRow& lhs, const XLRow& rhs)
    {
        if (lhs.m_rowNode && !rhs.m_rowNode)
            return false;
        if (!lhs.m_rowNode && !rhs.m_rowNode)
            return true;
        return *lhs.m_rowNode == *rhs.m_rowNode;
    }

    bool XLRow::isLessThan(const XLRow& lhs, const XLRow& rhs)
    {
        return *lhs.m_rowNode < *rhs.m_rowNode;
    }

}    // namespace OpenXLSX

// ========== XLRowIterator  ================================================ //
namespace OpenXLSX
{
    /**
     * @details
     * @pre
     * @post
     */
    XLRowIterator::XLRowIterator(const XLRowRange& rowRange, XLIteratorLocation loc)
        : m_dataNode(std::make_unique<XMLNode>(*rowRange.m_dataNode)),
          m_firstRow(rowRange.m_firstRow),
          m_lastRow(rowRange.m_lastRow),
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
            auto temp = XLRowIterator(other);
            std::swap(*this, temp);
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
        auto rowNode   = m_currentRow.m_rowNode->next_sibling();

        if (rowNumber > m_lastRow)
            m_currentRow = XLRow();

        else if (!rowNode || rowNode.attribute("r").as_ullong() != rowNumber) {
            rowNode = m_dataNode->insert_child_after("row", *m_currentRow.m_rowNode);
            rowNode.append_attribute("r").set_value(rowNumber);
            m_currentRow = XLRow(rowNode, m_sharedStrings);
        }

        else
            m_currentRow = XLRow(rowNode, m_sharedStrings);

        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowIterator XLRowIterator::operator++(int)    // NOLINT
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
    bool XLRowIterator::operator==(const XLRowIterator& rhs) const
    {
        return m_currentRow == rhs.m_currentRow;
    }

    /**
     * @details
     * @pre
     * @post
     */
    bool XLRowIterator::operator!=(const XLRowIterator& rhs) const
    {
        return !(m_currentRow == rhs.m_currentRow);
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowIterator::operator bool() const
    {
        return false;
    }

}    // namespace OpenXLSX

// ========== XLRowRange  =================================================== //
namespace OpenXLSX
{
    /**
     * @details
     * @pre
     * @post
     */
    XLRowRange::XLRowRange(const XMLNode& dataNode, uint32_t first, uint32_t last, const OpenXLSX::XLSharedStrings& sharedStrings)
        : m_dataNode(std::make_unique<XMLNode>(dataNode)),
          m_firstRow(first),
          m_lastRow(last),
          m_sharedStrings(sharedStrings)
    {}

    /**
     * @details
     * @pre
     * @post
     */
    XLRowRange::XLRowRange(const XLRowRange& other)
        : m_dataNode(std::make_unique<XMLNode>(*other.m_dataNode)),
          m_firstRow(other.m_firstRow),
          m_lastRow(other.m_lastRow),
          m_sharedStrings(other.m_sharedStrings)
    {}

    /**
     * @details
     * @pre
     * @post
     */
    XLRowRange::XLRowRange(XLRowRange&& other) noexcept = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowRange::~XLRowRange() = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowRange& XLRowRange::operator=(const XLRowRange& other)
    {
        if (&other != this) {
            auto temp = XLRowRange(other);
            std::swap(*this, temp);
        }

        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowRange& XLRowRange::operator=(XLRowRange&& other) noexcept = default;

    /**
     * @details
     * @pre
     * @post
     */
    uint32_t XLRowRange::rowCount() const
    {
        return m_lastRow - m_firstRow + 1;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowIterator XLRowRange::begin()
    {
        return XLRowIterator(*this, XLIteratorLocation::Begin);
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowIterator XLRowRange::end()
    {
        return XLRowIterator(*this, XLIteratorLocation::End);
    }

}    // namespace OpenXLSX
