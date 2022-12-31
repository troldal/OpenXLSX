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

  Written by Akira SHIMAHARA

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
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLTableRows.hpp"

using namespace OpenXLSX;



// ========== XLTableRowIterator  ================================================ //


 XLTableRowIterator:: XLTableRowIterator(const XLTableRows& tableRows, XLIteratorLocation loc):
                m_range (XLCellRange(  *tableRows.m_dataNode,
                        XLCellReference (tableRows.m_firstRow, tableRows.m_firstCol),
                        XLCellReference (tableRows.m_firstRow, tableRows.m_lastCol),
                        tableRows.m_worksheet)),
                m_firstIterRow(tableRows.m_firstRow),
                m_lastIterRow(tableRows.m_lastRow)
{
    if (loc == XLIteratorLocation::End)
        m_range = XLCellRange(  *tableRows.m_dataNode,
                            XLCellReference (tableRows.m_lastRow+1, tableRows.m_firstCol),
                            XLCellReference (tableRows.m_lastRow+1, tableRows.m_lastCol),
                            tableRows.m_worksheet);
}

 XLTableRowIterator::~ XLTableRowIterator() = default;

 XLTableRowIterator::XLTableRowIterator(const XLTableRowIterator& other)
        : m_range(other.m_range)
{ }

XLTableRowIterator::XLTableRowIterator(XLTableRowIterator&& other) noexcept = default;


XLTableRowIterator& XLTableRowIterator::operator=(const XLTableRowIterator& other)
{
    if (&other != this) {
        auto temp = XLTableRowIterator(other);
        std::swap(*this, temp);
    }
    return *this;
}

XLTableRowIterator& XLTableRowIterator::operator=(XLTableRowIterator&& other) noexcept = default;



XLTableRowIterator& XLTableRowIterator::operator++()
{
    m_range.offset(1);/*
    if(m_range.rangeCoordinates().first.first > m_lastIterRow)
        m_range = XLCellRange();
    */

    return *this;
}

/**
 * @details
 * @pre
 * @post
 */
XLTableRowIterator XLTableRowIterator::operator++(int)    // NOLINT
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
XLTableRowIterator::reference XLTableRowIterator::operator*()
{
    return m_range;
}

/**
 * @details
 * @pre
 * @post
 */
XLTableRowIterator::pointer XLTableRowIterator::operator->()
{
    return &m_range;
}


/**
 * @details
 * @pre
 * @post
 */
bool XLTableRowIterator::operator==(const XLTableRowIterator& rhs) const
{
    return  m_range == rhs.m_range;
}

/**
 * @details
 * @pre
 * @post
 */
bool XLTableRowIterator::operator!=(const XLTableRowIterator& rhs) const
{
    return !(*this == rhs);
}

/**
 * @details
 * @pre
 * @post
 */
XLTableRowIterator::operator bool() const
{
    return false;
}


// ========== XLTableRows  =================================================== //

XLTableRows::XLTableRows(const XMLNode& dataNode, 
                        uint32_t firstRow, 
                        uint32_t lastRow,
                        uint16_t firstCol,
                        uint16_t lastCol,
                        const XLWorksheet* wks):
        XLRowRange(dataNode, firstRow, lastRow, wks),
        m_firstCol(firstCol),
        m_lastCol(lastCol)
{}

/**
 * @brief Copy Constructor
 */
XLTableRows::XLTableRows(const XLTableRows& other):
        XLRowRange(other),
        m_firstCol(other.m_firstCol),
        m_lastCol(other.m_lastCol)
{}

/**
 * @brief Move Constructor
 */
XLTableRows::XLTableRows(XLTableRows&& other) noexcept = default;

XLTableRows::~XLTableRows() = default;

/**
 * @brief Copy assignment operator.
 */
XLTableRows& XLTableRows::operator=(const XLTableRows& other)
{
    if (&other != this) {
        auto temp = XLTableRows(other);
        std::swap(*this, temp);
    }
    return *this;
}

/**
 * @brief Move assignment operator.
 */
XLTableRows& XLTableRows::operator=(XLTableRows&& other) noexcept = default;


XLCellRange XLTableRows::operator[](uint32_t index) const
{
    uint32_t row = m_firstRow + index;

    if (row > m_lastRow)
        return XLCellRange(*m_dataNode,
                        XLCellReference (m_firstRow, m_firstCol),
                        XLCellReference (m_lastRow, m_lastCol),
                        m_worksheet);
    
    return XLCellRange(*m_dataNode,
                        XLCellReference (row, m_firstCol),
                        XLCellReference (row, m_lastCol),
                        m_worksheet);
}


/**
 * @details
 * @pre
 * @post
 */
XLTableRowIterator XLTableRows::begin()
{
    return XLTableRowIterator(*this, XLIteratorLocation::Begin);
}

/**
 * @details
 * @pre
 * @post
 */
XLTableRowIterator XLTableRows::end()
{
    return XLTableRowIterator(*this, XLIteratorLocation::End);
}

/**
 * @details
 * @pre
 * @post
 */
uint32_t XLTableRows::numRows() const
{
    return m_firstRow + 1 - m_lastRow;
}