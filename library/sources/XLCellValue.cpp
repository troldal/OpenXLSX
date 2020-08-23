//
// Created by Kenneth Balslev on 19/08/2020.
//

#include "XLCellValue.hpp"
#include "XLCellValueProxy.hpp"
#include "XLException.hpp"
#include <pugixml.hpp>

using namespace OpenXLSX;

XLCellValue::XLCellValue()                                       = default;
XLCellValue::XLCellValue(const OpenXLSX::XLCellValue& other)     = default;
XLCellValue::XLCellValue(OpenXLSX::XLCellValue&& other) noexcept = default;
XLCellValue::~XLCellValue()                                      = default;
XLCellValue& OpenXLSX::XLCellValue::operator=(const OpenXLSX::XLCellValue& other) = default;
XLCellValue& OpenXLSX::XLCellValue::operator=(OpenXLSX::XLCellValue&& other) noexcept = default;

XLCellValue::XLCellValue(const XLCellValueProxy& proxy) : XLCellValue(proxy.getValue()) {}

XLCellValue& XLCellValue::clear()
{
    m_type  = XLValueType::Empty;
    m_value = "";
    return *this;
}

XLCellValue& XLCellValue::setError()
{
    m_type  = XLValueType::Error;
    m_value = "";
    return *this;
}

XLValueType XLCellValue::type() const
{
    return m_type;
}
