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
#include <cstring>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCellValue.hpp"
#include "XLException.hpp"
#include "XLSharedStrings.hpp"

using namespace OpenXLSX;

namespace
{
    /**
     * @brief The XLCellType class is an enumeration of the possible cell types, as recognized by Excel.
     */
    enum class XLCellType { Empty, Boolean, Number, Error, String };

    /**
     * @details Determine the cell type, based on the contents of the underlying XML file, and returns the corresponding
     * XLCellType object.
     * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
     * @post The current object, and any associated objects, are unchanged.
     */
    XLCellType getCellType(XMLNode cellNode)
    {
        // ===== If neither a Type attribute or a value node is present, the cell is empty.
        if (!cellNode.attribute("t") && !cellNode.child("v")) {
            return XLCellType::Empty;
        }

        // ===== If a Type attribute is not present, but a value node is, the cell contains a number.
        if ((!cellNode.attribute("t") || (strcmp(cellNode.attribute("t").value(), "n") == 0 && cellNode.child("v") != nullptr))) {
            return XLCellType::Number;
        }

        // ===== If the cell is of type "s", the cell contains a shared string.
        if (cellNode.attribute("t") != nullptr && strcmp(cellNode.attribute("t").value(), "s") == 0) {
            return XLCellType::String;
        }

        // ===== If the cell is of type "inlineStr", the cell contains an inline string.
        if (cellNode.attribute("t") != nullptr && strcmp(cellNode.attribute("t").value(), "inlineStr") == 0) {
            return XLCellType::String;
        }

        // ===== If the cell is of type "str", the cell contains an ordinary string.
        if (cellNode.attribute("t") != nullptr && strcmp(cellNode.attribute("t").value(), "str") == 0) {
            return XLCellType::String;
        }

        // ===== If the cell is of type "b", the cell contains a boolean.
        if (cellNode.attribute("t") != nullptr && strcmp(cellNode.attribute("t").value(), "b") == 0) {
            return XLCellType::Boolean;
        }

        // ===== Otherwise, the cell contains an error.
        return XLCellType::Error;    // the m_typeAttribute has the ValueAsString "e"
    }

    /**
     * @brief
     */
    enum class XLNumberType { Integer, Float };

    /**
     * @details The number type (integer or floating point) is determined simply by identifying whether or not a decimal
     * point is present in the input string. If present, the number type is floating point.
     */
    XLNumberType DetermineNumberType(const std::string& numberString)
    {
        if (numberString.find('.') != std::string::npos || numberString.find("E-") != std::string::npos ||
            numberString.find("e-") != std::string::npos)
            return XLNumberType::Float;

        return XLNumberType::Integer;
    }
}    // namespace

/**
 * @details The constructor sets the m_parentCell to the value of the input parameter. The m_parentCell member variable
 * is a reference to the parent XLCell object, in order to avoid null objects.
 * @pre The parent input parameter is a valid XLCell object.
 * @post A valid XLCellValue object has been constructed.
 */
XLCellValue::XLCellValue(const XMLNode& cellNode, XLSharedStrings* sharedStrings) noexcept
    : m_cellNode(std::make_unique<XMLNode>(cellNode)),
      m_sharedStrings(sharedStrings)
{}

/**
 * @details
 */
XLCellValue::~XLCellValue() = default;

/**
 * @details
 */
XLCellValue::XLCellValue(const XLCellValue& other)
    : m_cellNode(std::make_unique<XMLNode>(*other.m_cellNode)),
      m_sharedStrings(other.m_sharedStrings)
{}

/**
 * @details The copy constructor and copy assignment operator works differently for XLCellValue objects.
 * The copy constructor creates an exact copy of the object, with the same parent XLCell object. The copy
 * assignment operator only copies the underlying cell value and type attribute to the target object.
 * @pre Both the lhs and rhs are valid objects.
 * @post Successful assignment to the target object.
 */
XLCellValue& XLCellValue::operator=(const XLCellValue& other)
{
    if (&other != this) {
        for (auto& attribute : m_cellNode->attributes())
            if (std::string_view(attribute.name()) != "r") m_cellNode->remove_attribute(attribute);

        for (auto& attribute : other.m_cellNode->attributes())
            if (std::string_view(attribute.name()) != "r") m_cellNode->append_copy(attribute);

        m_cellNode->remove_child(m_cellNode->child("v"));
        m_cellNode->append_copy(other.m_cellNode->child("v"));
    }

    return *this;
}

/**
 * @brief
 * @param other
 * @return
 * @todo Currently, the move constructor is identical to the copy constructor. Ensure that this is the correct
 * behaviour.
 */
XLCellValue& XLCellValue::operator=(XLCellValue&& other) noexcept
{
    if (&other != this) {
        for (auto& attribute : m_cellNode->attributes())
            if (std::string_view(attribute.name()) != "r") m_cellNode->remove_attribute(attribute);

        for (auto& attribute : other.m_cellNode->attributes())
            if (std::string_view(attribute.name()) != "r") m_cellNode->append_copy(attribute);

        m_cellNode->remove_child(m_cellNode->child("v"));
        m_cellNode->append_copy(other.m_cellNode->child("v"));
    }

    return *this;
}

/**
 * @details The assignment operator taking a std::string object as a parameter, calls the corresponding Set method
 * and returns the resulting object.
 * @pre The XLCellValue object and the stringValue parameter are both valid.
 * @post The underlying cell xml data has been set to the value of the input parameter and the type attribute has been
 * set to 'str'
 */
XLCellValue& XLCellValue::operator=(const std::string& stringValue)
{
    set(stringValue);
    return *this;
}

/**
 * @details The assignment operator taking a char* string as a parameter, calls the corresponding Set method
 * and returns the resulting object.
 * @pre The XLCellValue object and the stringValue parameter are both valid.
 * @post The underlying cell xml data has been set to the value of the input parameter and the type attribute has been
 * set to 'str'
 */
XLCellValue& XLCellValue::operator=(const char* stringValue)
{
    set(stringValue);
    return *this;
}

/**
 * @details If the value type is already a String, the value will be set to the new value. Otherwise, the m_value
 * member variable will be set to an XLString object with the given value.
 * @pre The XLCellValue object and the stringValue parameter are both valid.
 * @post The underlying cell xml data has been set to the value of the input parameter and the type attribute has been
 * set to 'str'
 */
void XLCellValue::set(const std::string& stringValue)
{
    set(stringValue.c_str());
}

/**
 * @details Converts the char* parameter to a std::string and calls the corresponding Set method.
 * @pre The XLCellValue object and the stringValue parameter are both valid.
 * @post The underlying cell xml data has been set to the value of the input parameter and the type attribute has been
 * set to 'str'
 */
void XLCellValue::set(const char* stringValue)
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

/**
 * @details Deletes the value node and type attribute if they exists.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The value node and the type attribute in the underlying xml data has been deleted.
 */
void XLCellValue::clear()
{
    m_cellNode->remove_child(m_cellNode->child("v"));
    m_cellNode->remove_attribute(m_cellNode->attribute("t"));
}

/**
 * @details Return the cell value as a string, by calling the AsString method of the m_value member variable.
 * @pre The m_value member variable is valid.
 * @post The current object, and any associated objects, are unchanged.
 */
std::string XLCellValue::asString() const
{
    if (strcmp(m_cellNode->attribute("t").value(), "b") == 0) {
        if (strcmp(m_cellNode->child("v").text().get(), "0") == 0) return "FALSE";

        return "TRUE";
    }

    if (strcmp(m_cellNode->attribute("t").value(), "s") == 0)
        return m_sharedStrings->getString(static_cast<uint32_t>(m_cellNode->child("v").text().as_ullong()));

    return m_cellNode->child("v").text().get();
}

/**
 * @details Determine the value type, based on the cell type, and return the corresponding XLValueType object.
 * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
 * @post The current object, and any associated objects, are unchanged.
 */
XLValueType XLCellValue::valueType() const
{
    switch (getCellType(*m_cellNode)) {
        case XLCellType::Empty:
            return XLValueType::Empty;

        case XLCellType::Error:
            return XLValueType::Error;

        case XLCellType::Number: {
            if (DetermineNumberType(m_cellNode->child("v").text().get()) == XLNumberType::Integer) return XLValueType::Integer;

            return XLValueType::Float;
        }

        case XLCellType::Boolean:
            return XLValueType::Boolean;

        default:
            return XLValueType::String;
    }
}

/**
 * @details
 */
void XLCellValue::setInteger(int64_t numberValue)
{
    if (!m_cellNode->child("v")) m_cellNode->append_child("v");

    m_cellNode->remove_attribute("t");
    m_cellNode->child("v").text().set(numberValue);
    m_cellNode->child("v").remove_attribute(m_cellNode->child("v").attribute("xml:space"));
    //    m_cellNode->child("v").attribute("xml:space").set_value("default");
}

/**
 * @details
 */
void XLCellValue::setBoolean(bool numberValue)
{
    if (!m_cellNode->attribute("t")) m_cellNode->append_attribute("t");
    if (!m_cellNode->child("v")) m_cellNode->append_child("v");

    m_cellNode->attribute("t").set_value("b");
    m_cellNode->child("v").text().set(numberValue ? 1 : 0);
    m_cellNode->child("v").remove_attribute(m_cellNode->child("v").attribute("xml:space"));
    //    m_cellNode->child("v").attribute("xml:space").set_value("default");
}

/**
 * @details
 */
void XLCellValue::setFloat(double numberValue)
{
    if (!m_cellNode->child("v")) m_cellNode->append_child("v");

    m_cellNode->remove_attribute("t");
    m_cellNode->child("v").text().set(numberValue);
    m_cellNode->child("v").remove_attribute(m_cellNode->child("v").attribute("xml:space"));
}

/**
 * @details
 */
int64_t XLCellValue::getInteger() const
{
    if (valueType() != XLValueType::Integer) throw XLException("Cell value is not Integer");
    return m_cellNode->child("v").text().as_llong();
}

/**
 * @details
 */
bool XLCellValue::getBoolean() const
{
    if (valueType() != XLValueType::Boolean) throw XLException("Cell value is not Boolean");
    return m_cellNode->child("v").text().as_bool();
}

/**
 * @details
 */
long double XLCellValue::getFloat() const
{
    if (valueType() != XLValueType::Float) throw XLException("Cell value is not Float");
    return static_cast<long double>(m_cellNode->child("v").text().as_double());
}

/**
 * @details
 */
const char* XLCellValue::getString() const
{
    if (valueType() != XLValueType::String) throw XLException("Cell value is not String");
    if (strcmp(m_cellNode->attribute("t").value(), "str") == 0)    // ordinary string
        return m_cellNode->child("v").text().get();
    if (strcmp(m_cellNode->attribute("t").value(), "s") == 0)    // shared string
        return m_sharedStrings->getString(static_cast<uint32_t>(m_cellNode->child("v").text().as_ullong()));

    throw XLException("Unknown string type");
}
