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
#include <stdexcept>

// ===== OpenXLSX Includes ===== //
#include "XLCellRange.hpp"
#include "XLSheet.hpp"

using namespace OpenXLSX;

namespace OpenXLSX
{
    XMLNode getRowNode(XMLNode sheetDataNode, uint32_t rowNumber);
    XMLNode getCellNode(XMLNode rowNode, uint16_t columnNumber);
}   

/**
 * @details From the two XLCellReference objects, the constructor calculates the dimensions of the range.
 * If the range exceeds the current bounds of the spreadsheet, the spreadsheet is resized to fit.
 * @pre
 * @post
 */
XLCellRange::XLCellRange(const XMLNode&         dataNode,
                         const XLCellReference& topLeft,
                         const XLCellReference& bottomRight,
                         const XLWorksheet*     wks)
    : m_dataNode(std::make_unique<XMLNode>(dataNode)),
      m_topLeft(topLeft),
      m_bottomRight(bottomRight),
      m_worksheet(wks)
{}

/**
 * @details
 * @pre
 * @post
 */
XLCellRange::XLCellRange(const XLCellRange& other)
    : m_dataNode(std::make_unique<XMLNode>(*other.m_dataNode)),
      m_topLeft(other.m_topLeft),
      m_bottomRight(other.m_bottomRight),
      m_worksheet(other.m_worksheet)
{}

/**
 * @details
 * @pre
 * @post
 */
XLCellRange::XLCellRange(XLCellRange&& other) noexcept = default;

/**
 * @details
 * @pre
 * @post
 */
XLCellRange::~XLCellRange() = default;

/**
 * @details
 * @pre
 * @post
 */
XLCellRange& XLCellRange::operator=(const XLCellRange& other)
{
    if (&other != this) {
        *m_dataNode     = *other.m_dataNode;
        m_topLeft       = other.m_topLeft;
        m_bottomRight   = other.m_bottomRight;
        m_worksheet = other.m_worksheet;
    }

    return *this;
}

/**
 * @details
 * @pre
 * @post
 */
XLCellRange& XLCellRange::operator=(XLCellRange&& other) noexcept
{
    if (&other != this) {
        *m_dataNode     = *other.m_dataNode;
        m_topLeft       = other.m_topLeft;
        m_bottomRight   = other.m_bottomRight;
       m_worksheet = other.m_worksheet;
    }

    return *this;
}

/**
 * @details
 * @pre
 * @post
 */
bool XLCellRange::operator==(const XLCellRange& rhs) const
{
    return ((m_topLeft == rhs.m_topLeft)&&(m_bottomRight == rhs.m_bottomRight));
}

/**
 * @details
 * @pre
 * @post
 */
bool XLCellRange::operator!=(const XLCellRange& rhs) const
{
    return !(*this == rhs);
}

XLCell XLCellRange::operator[](uint32_t index) const
{
    uint16_t nCol = numColumns();
    uint32_t row = m_topLeft.row() + index / nCol;
    uint16_t col = m_topLeft.column() + index % nCol;

    if(row > m_bottomRight.row())
        return XLCell();
    
    return XLCell(getCellNode(getRowNode(*m_dataNode, row), col), m_worksheet);   
}

/**
 * @details
 * @pre
 * @post
 */
uint32_t XLCellRange::numRows() const
{
    return m_bottomRight.row() + 1 - m_topLeft.row();
}

/**
 * @details
 * @pre
 * @post
 */
uint16_t XLCellRange::numColumns() const
{
    return m_bottomRight.column() + 1 - m_topLeft.column();
}

/**
 * @details
 * @pre
 * @post
 */
XLCellIterator XLCellRange::begin() const
{
    return XLCellIterator(*this, XLIteratorLocation::Begin);
}

/**
 * @details
 * @pre
 * @post
 */
XLCellIterator XLCellRange::end() const
{
    return XLCellIterator(*this, XLIteratorLocation::End);
}

/**
 * @details
 * @pre
 * @post
 */
void XLCellRange::offset(int row, int col)
{
    m_topLeft.offset(row,col);
    m_bottomRight.offset(row, col);
}



/**
 * @details
 * @pre
 * @post
 */
void XLCellRange::clear()
{
    for(auto& cell: *this) cell.value().clear();
}

std::pair<std::string,std::string> XLCellRange::topLeftBottomRight(const std::string& ref)
{
    std::string::size_type n = ref.find(':');
    if(n>ref.size())
        throw XLInputError("Invalid reference \"" + ref + "\" for a range");
    
    return std::make_pair(ref.substr(0, n),ref.substr(n+1));
}
