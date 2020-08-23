//
// Created by Kenneth Balslev on 19/08/2020.
//

// ===== External Includes ===== //
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLCellValue.hpp"
#include "XLCellValueProxy.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

/**
 * @details
 * @pre
 * @post
 */
XLCellValue::XLCellValue() = default;

/**
 * @details
 * @pre
 * @post
 */
XLCellValue::XLCellValue(const OpenXLSX::XLCellValue& other) = default;

/**
 * @details
 * @pre
 * @post
 */
XLCellValue::XLCellValue(OpenXLSX::XLCellValue&& other) noexcept = default;

/**
 * @details
 * @pre
 * @post
 */
XLCellValue::~XLCellValue() = default;

/**
 * @details
 * @pre
 * @post
 */
XLCellValue& OpenXLSX::XLCellValue::operator=(const OpenXLSX::XLCellValue& other) = default;

/**
 * @details
 * @pre
 * @post
 */
XLCellValue& OpenXLSX::XLCellValue::operator=(OpenXLSX::XLCellValue&& other) noexcept = default;

/**
 * @details
 * @pre
 * @post
 */
XLCellValue::XLCellValue(const XLCellValueProxy& proxy) : XLCellValue(proxy.getValue()) {}

/**
 * @details
 * @pre
 * @post
 */
XLCellValue& XLCellValue::clear()
{
    m_type  = XLValueType::Empty;
    m_value = "";
    return *this;
}

/**
 * @details
 * @pre
 * @post
 */
XLCellValue& XLCellValue::setError()
{
    m_type  = XLValueType::Error;
    m_value = "";
    return *this;
}

/**
 * @details
 * @pre
 * @post
 */
XLValueType XLCellValue::type() const
{
    return m_type;
}

/**
 * @details
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
