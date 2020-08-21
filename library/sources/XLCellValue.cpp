//
// Created by Kenneth Balslev on 19/08/2020.
//

#include "XLCellValue.hpp"
#include "XLException.hpp"
#include <pugixml.hpp>

OpenXLSX::XLCellValue::XLCellValue() = default;
OpenXLSX::XLCellValue::XLCellValue(const OpenXLSX::XLCellValue& other) = default;
OpenXLSX::XLCellValue::XLCellValue(OpenXLSX::XLCellValue&& other) noexcept = default;
OpenXLSX::XLCellValue::~XLCellValue() = default;
OpenXLSX::XLCellValue& OpenXLSX::XLCellValue::operator=(const OpenXLSX::XLCellValue& other) = default;
OpenXLSX::XLCellValue& OpenXLSX::XLCellValue::operator=(OpenXLSX::XLCellValue&& other) noexcept = default;

OpenXLSX::XLCellValue& OpenXLSX::XLCellValue::clear()
{
    m_type  = XLValueType::Empty;
    m_value = "";
    return *this;
}

OpenXLSX::XLCellValue& OpenXLSX::XLCellValue::setError()
{
    m_type  = XLValueType::Error;
    m_value = "";
    return *this;
}

OpenXLSX::XLValueType OpenXLSX::XLCellValue::type() const
{
    return m_type;
}
