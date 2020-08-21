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
XLCell::XLCellValueProxy::XLCellValueProxy(XLCell* cell) : m_cell(cell) {}
XLCell::XLCellValueProxy::~XLCellValueProxy() = default;
XLCell::XLCellValueProxy::XLCellValueProxy(const XLCell::XLCellValueProxy& other) = default;
XLCell::XLCellValueProxy::XLCellValueProxy(XLCell::XLCellValueProxy&& other) noexcept = default;
XLCell::XLCellValueProxy& XLCell::XLCellValueProxy::operator=(const XLCell::XLCellValueProxy& other) = default;
XLCell::XLCellValueProxy& XLCell::XLCellValueProxy::operator=(XLCell::XLCellValueProxy&& other) noexcept = default;

XLCell::XLCellValueProxy::operator XLCellValue()
{
    return m_cell->getValue();
}

/**
 * @details
 */
XLCell::XLCell() : m_cellNode(nullptr), m_sharedStrings(nullptr), m_valueProxy(XLCellValueProxy(this)) {}

/**
 * @details This constructor creates a XLCell object based on the cell XMLNode input parameter, and is
 * intended for use when the corresponding cell XMLNode already exist.
 * If a cell XMLNode does not exist (i.e., the cell is empty), use the relevant constructor to create an XLCell
 * from a XLCellReference parameter.
 * The initialization algorithm is as follows:
 *  -# Is the parameter a nullptr? If yes, throw a invalid_argument exception.
 *  -# Read the formula and getValue child nodes. These may be nullptr, if the cell is empty.
 *  -# Read the address, type and style attributes to the cellNode. The type attribute may be nullptr
 *  -# set the m_cellReference property using the getValue of the m_addressAttribute.
 *  -# Set the cell type, using the previous information:
 *      -# If there is no type attribute and no getValue node, the cell is empty.
 *      -# If there is no type attribute but there is a value node, the cell has a number getValue.
 *      -# Otherwise, determine the celltype based on the type attribute.
 */
XLCell::XLCell(const XMLNode& cellNode, XLSharedStrings* sharedStrings)
    : m_cellNode(std::make_unique<XMLNode>(cellNode)),
      m_sharedStrings(sharedStrings),
      m_valueProxy(XLCellValueProxy(this))
{}

/**
 * @details
 */
XLCell::XLCell(const XLCell& other)
    : m_cellNode(other.m_cellNode ? std::make_unique<XMLNode>(*other.m_cellNode) : nullptr),
      m_sharedStrings(other.m_sharedStrings),
      m_valueProxy(XLCellValueProxy(this))
{}

/**
 * @details
 */
XLCell::XLCell(XLCell&& other) noexcept
    : m_cellNode(std::move(other.m_cellNode)),
      m_sharedStrings(other.m_sharedStrings),
      m_valueProxy(XLCellValueProxy(this))
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
        m_cellNode      = other.m_cellNode ? std::make_unique<XMLNode>(*other.m_cellNode) : nullptr;
        m_sharedStrings = other.m_sharedStrings;
        m_valueProxy    = XLCellValueProxy(this);
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
        m_valueProxy    = XLCellValueProxy(this);
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
 * @details
 */
XLValueType XLCell::valueType() const
{
    // ===== If neither a Type attribute or a getValue node is present, the cell is empty.
    if (!m_cellNode->attribute("t") && !m_cellNode->child("v")) return XLValueType::Empty;

    // ===== If a Type attribute is not present, but a getValue node is, the cell contains a number.
    else if ((!m_cellNode->attribute("t") || (strcmp(m_cellNode->attribute("t").value(), "n") == 0 && m_cellNode->child("v") != nullptr))) {
        std::string numberString = m_cellNode->child("v").text().get();
        if (numberString.find('.') != std::string::npos || numberString.find("E-") != std::string::npos ||
            numberString.find("e-") != std::string::npos)
            return XLValueType::Float;
        else
            return XLValueType::Integer;
    }

    // ===== If the cell is of type "s", the cell contains a shared string.
    else if (m_cellNode->attribute("t") != nullptr && strcmp(m_cellNode->attribute("t").value(), "s") == 0)
        return XLValueType::String;    // NOLINT

    // ===== If the cell is of type "inlineStr", the cell contains an inline string.
    else if (m_cellNode->attribute("t") != nullptr && strcmp(m_cellNode->attribute("t").value(), "inlineStr") == 0)
        return XLValueType::String;

    // ===== If the cell is of type "str", the cell contains an ordinary string.
    else if (m_cellNode->attribute("t") != nullptr && strcmp(m_cellNode->attribute("t").value(), "str") == 0)
        return XLValueType::String;

    // ===== If the cell is of type "b", the cell contains a boolean.
    else if (m_cellNode->attribute("t") != nullptr && strcmp(m_cellNode->attribute("t").value(), "b") == 0)
        return XLValueType::Boolean;

    // ===== Otherwise, the cell contains an error.
    else
        return XLValueType::Error;    // the m_typeAttribute has the ValueAsString "e"
}

/**
 * @details
 */
XLCell::XLCellValueProxy& XLCell::value()
{
    return m_valueProxy;
}

const XLCell::XLCellValueProxy& XLCell::value() const
{
    return m_valueProxy;
}

/**
 * @details
 */
XLCellValue XLCell::getValue() const
{
    switch (valueType()) {
        case XLValueType::Empty:
            return XLCellValue().clear();

        case XLValueType::Float:
            return XLCellValue(m_cellNode->child("v").text().as_double());

        case XLValueType::Integer:
            return XLCellValue(m_cellNode->child("v").text().as_llong());

        case XLValueType::String:
            if (strcmp(m_cellNode->attribute("t").value(), "s") == 0)
                return XLCellValue(m_sharedStrings->getString(static_cast<uint32_t>(m_cellNode->child("v").text().as_ullong())));
            else if (strcmp(m_cellNode->attribute("t").value(), "str") == 0)
                return XLCellValue(m_cellNode->child("v").text().get());
            else
                throw XLException("Unknown string type");

        case XLValueType::Boolean:
            return XLCellValue(m_cellNode->child("v").text().as_bool());

        default:
            return XLCellValue().setError();
    }
}

/**
 * @details This function returns a const reference to the cellReference property.
 */
XLCellReference XLCell::cellReference() const
{
    return XLCellReference(m_cellNode->attribute("r").value());
}

/**
 * @details
 */
bool XLCell::hasFormula() const
{
    return m_cellNode->child("f") != nullptr;
}

/**
 * @details
 */
std::string XLCell::formula() const
{
    return m_cellNode->child("f").text().get();
}

/**
 * @details
 * @pre
 * @post
 */


void XLCell::setFormula(const std::string& newFormula)
{
    m_cellNode->child("f").text().set(newFormula.c_str());
}

/**
 * @details
 * @pre
 * @post
 */
void XLCell::setInteger(int64_t numberValue)
{
    if (!m_cellNode->child("v")) m_cellNode->append_child("v");

    m_cellNode->remove_attribute("t");
    m_cellNode->child("v").text().set(numberValue);
    m_cellNode->child("v").remove_attribute(m_cellNode->child("v").attribute("xml:space"));
}

/**
 * @details
 * @pre
 * @post
 */
void XLCell::setBoolean(bool numberValue)
{
    if (!m_cellNode->attribute("t")) m_cellNode->append_attribute("t");
    if (!m_cellNode->child("v")) m_cellNode->append_child("v");

    m_cellNode->attribute("t").set_value("b");
    m_cellNode->child("v").text().set(numberValue ? 1 : 0);
    m_cellNode->child("v").remove_attribute(m_cellNode->child("v").attribute("xml:space"));
}

/**
 * @details
 * @pre
 * @post
 */
void XLCell::setFloat(double numberValue)
{
    if (!m_cellNode->child("v")) m_cellNode->append_child("v");

    m_cellNode->remove_attribute("t");
    m_cellNode->child("v").text().set(numberValue);
    m_cellNode->child("v").remove_attribute(m_cellNode->child("v").attribute("xml:space"));
}

/**
 * @details
 * @pre
 * @post
 */
void XLCell::setString(const char* stringValue)
{
    if (!m_cellNode->attribute("t")) m_cellNode->append_attribute("t");
    if (!m_cellNode->child("v")) m_cellNode->append_child("v");

    m_cellNode->attribute("t").set_value("str");
    m_cellNode->child("v").text().set(stringValue);

    auto s = std::string_view(stringValue);
    if (s.front() == ' ' || s.back() == ' ') {
        if (!m_cellNode->attribute("xml:space")) m_cellNode->append_attribute("xml:space");
        m_cellNode->attribute("xml:space").set_value("preserve");
    }
}


