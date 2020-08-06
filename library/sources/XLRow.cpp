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
#include "XLCellReference.hpp"
#include "XLRow.hpp"

using namespace OpenXLSX;

/**
 * @details Constructs a new XLRow object from information in the underlying XML file. A pointer to the corresponding
 * node in the underlying XML file must be provided.
 */
XLRow::XLRow(const XMLNode& rowNode) : m_rowNode(std::make_unique<XMLNode>(rowNode)) {}

XLRow::XLRow(const XLRow& other) : m_rowNode(std::make_unique<XMLNode>(*other.m_rowNode)) {}

XLRow::~XLRow() = default;

XLRow& XLRow::operator=(const XLRow& other)
{
    if (&other != this) *m_rowNode = *other.m_rowNode;
    return *this;
}

/**
 * @details Returns the m_height member by value.
 */
double XLRow::height() const
{
    return m_rowNode->attribute("ht").as_double(15.0);
}

/**
 * @details Set the height of the row. This is done by setting the value of the 'ht' attribute and setting the
 * 'customHeight' attribute to true.
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
 * @details Return the m_descent member by value.
 */
float XLRow::descent() const
{
    return m_rowNode->attribute("x14ac:dyDescent").as_float(0.25);
}

/**
 * @details Set the descent by setting the 'x14ac:dyDescent' attribute in the XML file
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
 */
bool XLRow::isHidden() const
{
    return m_rowNode->attribute("hidden").as_bool(false);
}

/**
 * @details Set the hidden state by setting the 'hidden' attribute to true or false.
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
 */
int64_t XLRow::rowNumber() const
{
    return static_cast<int64_t>(m_rowNode->attribute("r").as_ullong());
}

/**
 * @details Get the number of cells in the row, by returning the size of the m_cells vector.
 */
unsigned int XLRow::cellCount() const
{
    return XLCellReference(m_rowNode->last_child().attribute("r").value()).column();
}
