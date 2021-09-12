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
#include "XLCellRange.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
XLCell::XLCell() : m_cellNode(nullptr),
                   m_sharedStrings(nullptr),
                   m_valueProxy(XLCellValueProxy(this, m_cellNode.get())),
                   m_formulaProxy(XLFormulaProxy(this, m_cellNode.get())){}

/**
 * @details This constructor creates a XLCell object based on the cell XMLNode input parameter, and is
 * intended for use when the corresponding cell XMLNode already exist.
 * If a cell XMLNode does not exist (i.e., the cell is empty), use the relevant constructor to create an XLCell
 * from a XLCellReference parameter.
 */
XLCell::XLCell(const XMLNode& cellNode, XLSharedStrings* sharedStrings)
    : m_cellNode(std::make_unique<XMLNode>(cellNode)),
      m_sharedStrings(sharedStrings),
      m_valueProxy(XLCellValueProxy(this, m_cellNode.get())),
      m_formulaProxy(XLFormulaProxy(this, m_cellNode.get()))
{}

/**
 * @details
 */
XLCell::XLCell(const XLCell& other)
    : m_cellNode(other.m_cellNode ? std::make_unique<XMLNode>(*other.m_cellNode) : nullptr),
      m_sharedStrings(other.m_sharedStrings),
      m_valueProxy(XLCellValueProxy(this, m_cellNode.get())),
      m_formulaProxy(XLFormulaProxy(this, m_cellNode.get()))
{}

/**
 * @details
 */
XLCell::XLCell(XLCell&& other) noexcept
    : m_cellNode(std::move(other.m_cellNode)),
      m_sharedStrings(other.m_sharedStrings),
      m_valueProxy(XLCellValueProxy(this, m_cellNode.get())),
      m_formulaProxy(XLFormulaProxy(this, m_cellNode.get()))
{}

/**
 * @details
 */
XLCell::~XLCell() = default;

/**
 * @details
 */
XLCell& XLCell::operator=(const XLCell& other)
{
    if (&other != this) {
    XLCell temp = other;
    std::swap(*this, temp);

    }

    return *this;
}

/**
 * @details
 */
XLCell& XLCell::operator=(XLCell&& other) noexcept
{
    if (&other != this) {
        m_cellNode      = std::move(other.m_cellNode);
        m_sharedStrings = other.m_sharedStrings;
        m_valueProxy    = XLCellValueProxy(this, m_cellNode.get());
    }

    return *this;
}

/**
 * @details This methods copies a range into a new location, with the top left cell being located in the target cell.
 * The copying is done in the following way:
 * - Define a range with the top left cell being the target cell.
 * - Calculate the size of the range from the originating range and set the new range to the same size.
 * - Copy the contents of the original range to the new range.
 * - Return a reference to the first cell in the new range.
 * @pre
 * @post
 * @todo Consider what happens if the target range extends beyond the maximum spreadsheet limits
 */
XLCell& XLCell::operator=(const XLCellRange& range)
{
    auto            first = cellReference();
    XLCellReference last(first.row() + range.numRows() - 1, first.column() + range.numColumns() - 1);
    XLCellRange     rng(m_cellNode->parent().parent(), first, last, nullptr);
    rng = range;

    return *this;
}

/**
 * @details
 */
XLCell::operator bool() const
{
    return m_cellNode && *m_cellNode;
}

/**
 * @details This function returns a const reference to the cellReference property.
 */
XLCellReference XLCell::cellReference() const
{
    if (!*this) throw XLInternalError("XLCell object has not been properly initiated.");
    return XLCellReference(m_cellNode->attribute("r").value());
}

/**
 * @details
 */
bool XLCell::hasFormula() const
{
    if (!*this) return false;
    return m_cellNode->child("f") != nullptr;
}

XLFormulaProxy& XLCell::formula()
{
    if (!*this) throw XLInternalError("XLCell object has not been properly initiated.");
    return m_formulaProxy;
}

const XLFormulaProxy& XLCell::formula() const
{
    if (!*this) throw XLInternalError("XLCell object has not been properly initiated.");
    return m_formulaProxy;
}

/**
 * @details
 */
//std::string XLCell::formula() const
//{
//    if (!*this) throw XLInternalError("XLCell object has not been properly initiated.");
//    return m_cellNode->child("f").text().get();
//}

/**
 * @details
 * @pre
 * @post
 */
//void XLCell::setFormula(const std::string& newFormula)
//{
//    if (!*this) throw XLInternalError("XLCell object has not been properly initiated.");
//    m_cellNode->child("f").text().set(newFormula.c_str());
//}

/**
 * @details
 * @pre
 * @post
 */
XLCellValueProxy& XLCell::value()
{
    if (!*this) throw XLInternalError("XLCell object has not been properly initiated.");
    return m_valueProxy;
}

/**
 * @details
 * @pre
 * @post
 */
const XLCellValueProxy& XLCell::value() const
{
    if (!*this) throw XLInternalError("XLCell object has not been properly initiated.");
    return m_valueProxy;
}