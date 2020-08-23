//
// Created by Kenneth Balslev on 23/08/2020.
//

// ===== External Includes ===== //
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCell.hpp"
#include "XLCellValueProxy.hpp"
#include "XLException.hpp"
#include "XLSharedStrings.hpp"

using namespace OpenXLSX;

/**
 * @details
 * @pre
 * @post
 */
XLCellValueProxy::XLCellValueProxy(XLCell* cell, XMLNode* cellNode) : m_cell(cell), m_cellNode(cellNode) {}

/**
 * @details
 * @pre
 * @post
 */
XLCellValueProxy::~XLCellValueProxy() = default;

/**
 * @details
 * @pre
 * @post
 */
XLCellValueProxy::XLCellValueProxy(const XLCellValueProxy& other) = default;

/**
 * @details
 * @pre
 * @post
 */
XLCellValueProxy::XLCellValueProxy(XLCellValueProxy&& other) noexcept = default;

/**
 * @details
 * @pre
 * @post
 */
XLCellValueProxy& XLCellValueProxy::operator=(const XLCellValueProxy& other)
{
    if (&other != this) {
        *this = other.getValue();
    }

    return *this;
}

/**
 * @details
 * @pre
 * @post
 */
XLCellValueProxy& XLCellValueProxy::operator=(XLCellValueProxy&& other) noexcept = default;

/**
 * @details
 * @pre
 * @post
 */
XLCellValueProxy::operator XLCellValue()
{
    return getValue();
}

/**
 * @details
 * @pre
 * @post
 */
XLCellValueProxy& XLCellValueProxy::clear()
{
    m_cellNode->remove_attribute("t");
    m_cellNode->remove_attribute(" xml:space");
    m_cellNode->remove_child("v");
    return *this;
}
/**
 * @details
 * @pre
 * @post
 */
XLCellValueProxy& XLCellValueProxy::setError()
{
    if(!m_cellNode->attribute("t")) m_cellNode->append_attribute("t");
    m_cellNode->attribute("t").set_value("e");
    m_cellNode->remove_attribute(" xml:space");
    m_cellNode->remove_child("v");
    return *this;
}

/**
 * @details
 * @pre
 * @post
 */
XLValueType XLCellValueProxy::type() const
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
 * @pre
 * @post
 */
std::string XLCellValueProxy::typeAsString() const
{
    switch (type()) {
        case XLValueType::Empty:
            return "empty";
        case XLValueType::Boolean:
            return "boolean";
        case XLValueType::Integer:
            return "integer";
        case XLValueType::Float:
            return "float";
        case XLValueType::String:
            return "string";
        default:
            return "error";
    }
}

/**
 * @details
 * @pre
 * @post
 */
void XLCellValueProxy::setInteger(int64_t numberValue)
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
void XLCellValueProxy::setBoolean(bool numberValue)
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
void XLCellValueProxy::setFloat(double numberValue)
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
void XLCellValueProxy::setString(const char* stringValue)
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
 * @details
 * @pre
 * @post
 */
XLCellValue XLCellValueProxy::getValue() const
{
    switch (type()) {
        case XLValueType::Empty:
            return XLCellValue().clear();

        case XLValueType::Float:
            return XLCellValue(m_cellNode->child("v").text().as_double());

        case XLValueType::Integer:
            return XLCellValue(m_cellNode->child("v").text().as_llong());

        case XLValueType::String:
            if (strcmp(m_cellNode->attribute("t").value(), "s") == 0)
                return XLCellValue(m_cell->m_sharedStrings->getString(static_cast<uint32_t>(m_cellNode->child("v").text().as_ullong())));
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
