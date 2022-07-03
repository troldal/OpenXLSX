//
// Created by Kenneth Balslev on 19/08/2020.
//

// ===== External Includes ===== //
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCell.hpp"
#include "XLCellValue.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

/**
 * @details Constructor. Default implementation has been used.
 * @pre
 * @post
 */
XLCellValue::XLCellValue() = default;

/**
 * @details Copy constructor. The default implementation will be used.
 * @pre The object to be copied must be valid.
 * @post A valid copy is constructed.
 */
XLCellValue::XLCellValue(const OpenXLSX::XLCellValue& other) = default;

/**
 * @details Move constructor. The default implementation will be used.
 * @pre The object to be copied must be valid.
 * @post A valid copy is constructed.
 */
XLCellValue::XLCellValue(OpenXLSX::XLCellValue&& other) noexcept = default;

/**
 * @details Destructor. The default implementation will be used
 * @pre None.
 * @post The object is destructed.
 */
XLCellValue::~XLCellValue() = default;

/**
 * @details Copy assignment operator. The default implementation will be used.
 * @pre The object to be copied must be a valid object.
 * @post A the copied-to object is valid.
 */
XLCellValue& OpenXLSX::XLCellValue::operator=(const OpenXLSX::XLCellValue& other) = default;

/**
 * @details Move assignment operator. The default implementation will be used.
 * @pre The object to be moved must be a valid object.
 * @post The moved-to object is valid.
 */
XLCellValue& OpenXLSX::XLCellValue::operator=(OpenXLSX::XLCellValue&& other) noexcept = default;

/**
 * @details Clears the contents of the XLCellValue object. Setting the value to an empty string is not sufficient
 * (as an empty string is still a valid string). The m_type variable must also be set to XLValueType::Empty.
 * @pre
 * @post
 */
XLCellValue& XLCellValue::clear()
{
    m_type  = XLValueType::Empty;
    m_value = std::string("");
    return *this;
}

/**
 * @details Sets the value type to XLValueType::Error. The value will be set to an empty string.
 * @pre
 * @post
 */
XLCellValue& XLCellValue::setError(const std::string &error)
{
    m_type  = XLValueType::Error;
    m_value = error;
    return *this;
}

/**
 * @details Get the value type of the current object.
 * @pre
 * @post
 */
XLValueType XLCellValue::type() const
{
    return m_type;
}

/**
 * @details Get the value type of the current object, as a string representation
 * @pre
 * @post
 */
std::string XLCellValue::typeAsString() const
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
 * @details Constructor
 * @pre The cell and cellNode pointers must not be nullptr and must point to valid objects.
 * @post A valid XLCellValueProxy has been created.
 */
XLCellValueProxy::XLCellValueProxy(XLCell* cell, XMLNode* cellNode) : m_cell(cell), m_cellNode(cellNode)
{
    assert(cell);                  // NOLINT
//    assert(cellNode);              // NOLINT
//    assert(!cellNode->empty());    // NOLINT
}

/**
 * @details Destructor. Default implementation has been used.
 * @pre
 * @post
 */
XLCellValueProxy::~XLCellValueProxy() = default;

/**
 * @details Copy constructor. Default implementation has been used.
 * @pre
 * @post
 */
XLCellValueProxy::XLCellValueProxy(const XLCellValueProxy& other) = default;

/**
 * @details Move constructor. Default implementation has been used.
 * @pre
 * @post
 */
XLCellValueProxy::XLCellValueProxy(XLCellValueProxy&& other) noexcept = default;

/**
 * @details Copy assignment operator. The function is implemented in terms of the templated
 * value assignment operators, i.e. it is the XLCellValue that is that is copied,
 * not the object itself.
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
 * @details Move assignment operator. Default implementation has been used.
 * @pre
 * @post
 */
XLCellValueProxy& XLCellValueProxy::operator=(XLCellValueProxy&& other) noexcept = default;

/**
 * @details Implicitly convert the XLCellValueProxy object to a XLCellValue object.
 * @pre
 * @post
 */
XLCellValueProxy::operator XLCellValue()
{
    return getValue();
}

/**
 * @details Clear the contents of the cell. This removes all children of the cell node.
 * @pre The m_cellNode must not be null, and must point to a valid XML cell node object.
 * @post The cell node must be valid, but empty.
 */
XLCellValueProxy& XLCellValueProxy::clear()
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode);              // NOLINT
    assert(!m_cellNode->empty());    // NOLINT

    // ===== Remove the type attribute
    m_cellNode->remove_attribute("t");

    // ===== Disable space preservation (only relevant for strings).
    m_cellNode->remove_attribute(" xml:space");

    // ===== Remove the value node.
    m_cellNode->remove_child("v");
    return *this;
}
/**
 * @details Set the cell value to a error state. This will remove all children and attributes, except
 * the type attribute, which is set to "e"
 * @pre The m_cellNode must not be null, and must point to a valid XML cell node object.
 * @post The cell node must be valid.
 */
XLCellValueProxy& XLCellValueProxy::setError(const std::string &error)
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode);              // NOLINT
    assert(!m_cellNode->empty());    // NOLINT

    // ===== If the cell node doesn't have a type attribute, create it.
    if (!m_cellNode->attribute("t")) m_cellNode->append_attribute("t");

    // ===== Set the type to "e", i.e. error
    m_cellNode->attribute("t").set_value("e");

    // ===== If the cell node doesn't have a value child node, create it.
    if (!m_cellNode->child("v")) m_cellNode->append_child("v");

    // ===== Set the child value to the error
    m_cellNode->child("v").text().set(error.c_str());

    // ===== Disable space preservation (only relevant for strings).
    m_cellNode->remove_attribute(" xml:space");

    return *this;
}

/**
 * @details Get the value type for the cell.
 * @pre The m_cellNode must not be null, and must point to a valid XML cell node object.
 * @post No change should be made.
 */
XLValueType XLCellValueProxy::type() const
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode);              // NOLINT
    assert(!m_cellNode->empty());    // NOLINT

    // ===== If neither a Type attribute or a getValue node is present, the cell is empty.
    if (!m_cellNode->attribute("t") && !m_cellNode->child("v")) return XLValueType::Empty;

    // ===== If a Type attribute is not present, but a value node is, the cell contains a number.
    if ((!m_cellNode->attribute("t") || (strcmp(m_cellNode->attribute("t").value(), "n") == 0 && m_cellNode->child("v") != nullptr))) {
        std::string numberString = m_cellNode->child("v").text().get();
        if (numberString.find('.') != std::string::npos || numberString.find("E-") != std::string::npos ||
            numberString.find("e-") != std::string::npos)
            return XLValueType::Float;
        return XLValueType::Integer;
    }

    // ===== If the cell is of type "s", the cell contains a shared string.
    if (m_cellNode->attribute("t") != nullptr && strcmp(m_cellNode->attribute("t").value(), "s") == 0)
        return XLValueType::String;    // NOLINT

    // ===== If the cell is of type "inlineStr", the cell contains an inline string.
    if (m_cellNode->attribute("t") != nullptr && strcmp(m_cellNode->attribute("t").value(), "inlineStr") == 0)
        return XLValueType::String;

    // ===== If the cell is of type "str", the cell contains an ordinary string.
    if (m_cellNode->attribute("t") != nullptr && strcmp(m_cellNode->attribute("t").value(), "str") == 0)
        return XLValueType::String;

    // ===== If the cell is of type "b", the cell contains a boolean.
    if (m_cellNode->attribute("t") != nullptr && strcmp(m_cellNode->attribute("t").value(), "b") == 0)
        return XLValueType::Boolean;

    // ===== Otherwise, the cell contains an error.
    return XLValueType::Error;    // the m_typeAttribute has the ValueAsString "e"
}

/**
 * @details Get the value type of the current object, as a string representation.
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
 * @details Set cell to an integer value. This is private helper function for setting the cell value
 * directly in the underlying XML file.
 * @pre The m_cellNode must not be null, and must point to a valid XMLNode object.
 * @post The underlying XMLNode has been updated correctly, representing an integer value.
 */
void XLCellValueProxy::setInteger(int64_t numberValue)
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode);              // NOLINT
    assert(!m_cellNode->empty());    // NOLINT

    // ===== If the cell node doesn't have a value child node, create it.
    if (!m_cellNode->child("v")) m_cellNode->append_child("v");

    // ===== The type ("t") attribute is not required for number values.
    m_cellNode->remove_attribute("t");

    // ===== Set the text of the value node.
    m_cellNode->child("v").text().set(numberValue);

    // ===== Disable space preservation (only relevant for strings).
    m_cellNode->child("v").remove_attribute(m_cellNode->child("v").attribute("xml:space"));
}

/**
 * @details Set the cell to a bool value. This is private helper function for setting the cell value
 * directly in the underlying XML file.
 * @pre The m_cellNode must not be null, and must point to a valid XMLNode object.
 * @post The underlying XMLNode has been updated correctly, representing an bool value.
 */
void XLCellValueProxy::setBoolean(bool numberValue)
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode);              // NOLINT
    assert(!m_cellNode->empty());    // NOLINT

    // ===== If the cell node doesn't have a type child node, create it.
    if (!m_cellNode->attribute("t")) m_cellNode->append_attribute("t");

    // ===== If the cell node doesn't have a value child node, create it.
    if (!m_cellNode->child("v")) m_cellNode->append_child("v");

    // ===== Set the type attribute.
    m_cellNode->attribute("t").set_value("b");

    // ===== Set the text of the value node.
    m_cellNode->child("v").text().set(numberValue ? 1 : 0);

    // ===== Disable space preservation (only relevant for strings).
    m_cellNode->child("v").remove_attribute(m_cellNode->child("v").attribute("xml:space"));
}

/**
 * @details Set the cell to a floating point value. This is private helper function for setting the cell value
 * directly in the underlying XML file.
 * @pre The m_cellNode must not be null, and must point to a valid XMLNode object.
 * @post The underlying XMLNode has been updated correctly, representing a floating point value.
 */
void XLCellValueProxy::setFloat(double numberValue)
{
    // check for nan / inf
    if (std::isfinite(numberValue)) {
        // ===== Check that the m_cellNode is valid.
        assert(m_cellNode);              // NOLINT
        assert(!m_cellNode->empty());    // NOLINT

        // ===== If the cell node doesn't have a value child node, create it.
        if (!m_cellNode->child("v")) m_cellNode->append_child("v");

        // ===== The type ("t") attribute is not required for number values.
        m_cellNode->remove_attribute("t");

        // ===== Set the text of the value node.
        m_cellNode->child("v").text().set(numberValue);

        // ===== Disable space preservation (only relevant for strings).
        m_cellNode->child("v").remove_attribute(m_cellNode->child("v").attribute("xml:space"));
    }
    else {
        setError("#NUM!");
        return;
    }
}

/**
 * @details Set the cell to a string value. This is private helper function for setting the cell value
 * directly in the underlying XML file.
 * @pre The m_cellNode must not be null, and must point to a valid XMLNode object.
 * @post The underlying XMLNode has been updated correctly, representing a string value.
 */
void XLCellValueProxy::setString(const char* stringValue)
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode);              // NOLINT
    assert(!m_cellNode->empty());    // NOLINT

    // ===== If the cell node doesn't have a type child node, create it.
    if (!m_cellNode->attribute("t")) m_cellNode->append_attribute("t");

    // ===== If the cell node doesn't have a value child node, create it.
    if (!m_cellNode->child("v")) m_cellNode->append_child("v");

    // ===== Set the type attribute.
    m_cellNode->attribute("t").set_value("s");

    // ===== Get or create the index in the XLSharedStrings object.
    auto index = (m_cell->m_sharedStrings.stringExists(stringValue) ? m_cell->m_sharedStrings.getStringIndex(stringValue)
                                                                     : m_cell->m_sharedStrings.appendString(stringValue));

    // ===== Set the text of the value node.
    m_cellNode->child("v").text().set(index);

    // IMPLEMENTATION FOR EMBEDDED STRINGS:
    //    m_cellNode->attribute("t").set_value("str");
    //    m_cellNode->child("v").text().set(stringValue);
    //
    //    auto s = std::string_view(stringValue);
    //    if (s.front() == ' ' || s.back() == ' ') {
    //        if (!m_cellNode->attribute("xml:space")) m_cellNode->append_attribute("xml:space");
    //        m_cellNode->attribute("xml:space").set_value("preserve");
    //    }
}

/**
 * @details Get a copy of the XLCellValue object for the cell. This is private helper function for returning an
 * XLCellValue object corresponding to the cell value.
 * @pre The m_cellNode must not be null, and must point to a valid XMLNode object.
 * @post No changes should be made.
 */
XLCellValue XLCellValueProxy::getValue() const
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode);              // NOLINT
    assert(!m_cellNode->empty());    // NOLINT

    switch (type()) {
        case XLValueType::Empty:
            return XLCellValue().clear();

        case XLValueType::Float:
            return XLCellValue { m_cellNode->child("v").text().as_double() };

        case XLValueType::Integer:
            return XLCellValue { m_cellNode->child("v").text().as_llong() };

        case XLValueType::String:
            if (strcmp(m_cellNode->attribute("t").value(), "s") == 0)
                return XLCellValue { m_cell->m_sharedStrings.getString(static_cast<uint32_t>(m_cellNode->child("v").text().as_ullong())) };
            else if (strcmp(m_cellNode->attribute("t").value(), "str") == 0)
                return XLCellValue { m_cellNode->child("v").text().get() };
            else if (strcmp(m_cellNode->attribute("t").value(), "inlineStr") == 0)
                return XLCellValue { m_cellNode->child("is").child("t").text().get() };
            else
                throw XLInternalError("Unknown string type");

        case XLValueType::Boolean:
            return XLCellValue { m_cellNode->child("v").text().as_bool() };

        case XLValueType::Error:
            return XLCellValue().setError(m_cellNode->child("v").text().as_string());
            
        default:
            return XLCellValue().setError("");
    }
}
