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
#include "XLCellRange.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
XLCellRange::XLCellRange()
    : m_dataNode(std::make_unique<XMLNode>(XMLNode{})),
      m_topLeft(XLCellReference("A1")),
      m_bottomRight(XLCellReference("A1")),
      m_sharedStrings{},
      m_columnStyles{}
{}

/**
 * @details From the two XLCellReference objects, the constructor calculates the dimensions of the range.
 * If the range exceeds the current bounds of the spreadsheet, the spreadsheet is resized to fit.
 * @pre
 * @post
 */
XLCellRange::XLCellRange(const XMLNode&         dataNode,
                         const XLCellReference& topLeft,
                         const XLCellReference& bottomRight,
                         const XLSharedStrings& sharedStrings)
    : m_dataNode(std::make_unique<XMLNode>(dataNode)),
      m_topLeft(topLeft),
      m_bottomRight(bottomRight),
      m_sharedStrings(sharedStrings),
      m_columnStyles{}
{
    if (m_topLeft.row() > m_bottomRight.row() || m_topLeft.column() > m_bottomRight.column()) {
        using namespace std::literals::string_literals;
        throw XLInputError("XLCellRange constructor: topLeft ("s + topLeft.address() + ")"s
        /**/             + " does not point to a lower or equal row and column than bottomRight ("s + bottomRight.address() + ")"s);
    }
}

/**
 * @details
 * @pre
 * @post
 */
XLCellRange::XLCellRange(const XLCellRange& other)
    : m_dataNode(std::make_unique<XMLNode>(*other.m_dataNode)),
      m_topLeft(other.m_topLeft),
      m_bottomRight(other.m_bottomRight),
      m_sharedStrings(other.m_sharedStrings),
      m_columnStyles(other.m_columnStyles)
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
        m_sharedStrings = other.m_sharedStrings;
        m_columnStyles  = other.m_columnStyles;
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
        m_sharedStrings = other.m_sharedStrings;
        m_columnStyles  = other.m_columnStyles;
    }

    return *this;
}

/**
 * @details Predetermine all defined column styles & gather them in a vector for performant access when XLCellIterator creates new cells
 */
void XLCellRange::fetchColumnStyles()
{
    XMLNode cols = m_dataNode->parent().child("cols");

    uint16_t vecPos = 0;
    XMLNode col = cols.first_child_of_type(pugi::node_element);
    while (not col.empty()) {
        uint16_t minCol = col.attribute("min").as_int(0);
        uint16_t maxCol = col.attribute("max").as_int(0);
        if (minCol > maxCol || !minCol || !maxCol) {
            using namespace std::literals::string_literals;
            throw XLInputError(
                "column attributes min (\""s + col.attribute("min").value() + "\") and max (\""s + col.attribute("min").value() + "\")"s
                " must be set and min must not be larger than max"s
            );
        }
        if (maxCol > m_columnStyles.size()) m_columnStyles.resize(maxCol);                        // resize m_columnStyles if necessary
        for( ; vecPos + 1 < minCol; ++vecPos ) m_columnStyles[ vecPos ] = XLDefaultCellFormat;    // set all non-defined columns to default
        XLStyleIndex colStyle = col.attribute("style").as_uint(XLDefaultCellFormat);              // acquire column style attribute
        for( ; vecPos < maxCol; ++vecPos ) m_columnStyles[ vecPos ] = colStyle;               // set all covered columns to defined style
        col = col.next_sibling_of_type(pugi::node_element);    // advance to next <col> entry, if any
    }
}

/**
 * @details
 */
const XLCellReference XLCellRange::topLeft() const { return m_topLeft; }

/**
 * @details
 */
const XLCellReference XLCellRange::bottomRight() const { return m_bottomRight; }

/**
 * @details Evaluate the top left and bottom right cells as string references and
 *          concatenate them with a colon ':'
 */
std::string XLCellRange::address() const { return m_topLeft.address() + ":" + m_bottomRight.address(); }

/**
 * @details
 * @pre
 * @post
 */
uint32_t XLCellRange::numRows() const { return m_bottomRight.row() + 1 - m_topLeft.row(); }

/**
 * @details
 * @pre
 * @post
 */
uint16_t XLCellRange::numColumns() const { return m_bottomRight.column() + 1 - m_topLeft.column(); }

/**
 * @details
 * @pre
 * @post
 */
XLCellIterator XLCellRange::begin() const { return XLCellIterator(*this, XLIteratorLocation::Begin, &m_columnStyles); }

/**
 * @details
 * @pre
 * @post
 */
XLCellIterator XLCellRange::end() const { return XLCellIterator(*this, XLIteratorLocation::End, &m_columnStyles); }

/**
 * @details
 * @pre
 * @post
 */
void XLCellRange::clear()
{
    for (auto& cell : *this) cell.value().clear();
}

/**
 * @details Iterate over the full range and set format for each existing cell
 */
bool XLCellRange::setFormat(size_t cellFormatIndex)
{
    // ===== Iterate over all cells in the range
    for (auto it = begin(); it != end(); ++it)
        if (!it->setCellFormat(cellFormatIndex))    // attempt to set cell format
            return false;                               // fail if any setCellFormat failed
    return true; // success if loop finished nominally
}
