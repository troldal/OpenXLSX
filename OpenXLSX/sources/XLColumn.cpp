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
#include "XLColumn.hpp"

using namespace OpenXLSX;

/**
 * @details Assumes each node only has data for one column.
 */
XLColumn::XLColumn(const XMLNode& columnNode) : m_columnNode(std::make_unique<XMLNode>(columnNode)) {}

XLColumn::XLColumn(const XLColumn& other) : m_columnNode(std::make_unique<XMLNode>(*other.m_columnNode)) {}

XLColumn::XLColumn(XLColumn&& other) noexcept = default;

XLColumn::~XLColumn() = default;

XLColumn& XLColumn::operator=(const XLColumn& other)
{
    if (&other != this) *m_columnNode = *other.m_columnNode;
    return *this;
}

/**
 * @details
 */
float XLColumn::width() const
{
    return columnNode().attribute("width").as_float();
}

/**
 * @details
 */
void XLColumn::setWidth(float width) // NOLINT
{
    // Set the 'Width' attribute for the Cell. If it does not exist, create it.
    auto widthAtt = columnNode().attribute("width");
    if (!widthAtt) widthAtt = columnNode().append_attribute("width");

    widthAtt.set_value(width);

    // Set the 'customWidth' attribute for the Cell. If it does not exist, create it.
    auto customAtt = columnNode().attribute("customWidth");
    if (!customAtt) customAtt = columnNode().append_attribute("customWidth");

    customAtt.set_value("1");
}

/**
 * @details
 */
bool XLColumn::isHidden() const
{
    return columnNode().attribute("hidden").as_bool();
}

/**
 * @details
 */
void XLColumn::setHidden(bool state) // NOLINT
{
    auto hiddenAtt = columnNode().attribute("hidden");
    if (!hiddenAtt) hiddenAtt = columnNode().append_attribute("hidden");

    if (state)
        hiddenAtt.set_value("1");
    else
        hiddenAtt.set_value("0");
}

/**
 * @details
 */
XMLNode& XLColumn::columnNode() const
{
    return *m_columnNode;
}
