//
// Created by Troldal on 2019-01-21.
//

#include <XLCellValue.hpp>

#include "XLCellValue.hpp"
#include "XLCellValue_Impl.hpp"

using namespace OpenXLSX;

XLCellValue::XLCellValue(Impl::XLCellValue& value)
        : m_value(&value) {
}

XLCellValue& XLCellValue::operator=(const XLCellValue& other) {

    *m_value = *other.m_value;
    return *this;
}

XLCellValue& XLCellValue::operator=(const char* stringValue) {

    *m_value = stringValue;
    return *this;
}

XLCellValue& XLCellValue::operator=(const std::string& stringValue) {

    *m_value = stringValue;
    return *this;
}

std::string XLCellValue::AsString() const {

    return m_value->AsString();
}

void XLCellValue::Set(const char* stringValue) {

    m_value->Set(stringValue);
}

void XLCellValue::Set(const std::string& stringValue) {

    m_value->Set(stringValue);
}

void XLCellValue::Clear() {

    m_value->Clear();
}

XLValueType XLCellValue::ValueType() const {

    return m_value->ValueType();
}

void XLCellValue::SetInteger(long long int numberValue) {

    m_value->SetInteger(numberValue);
}

void XLCellValue::SetBoolean(bool numberValue) {

    m_value->SetBoolean(numberValue);
}

void XLCellValue::SetFloat(long double numberValue) {

    m_value->SetFloat(numberValue);
}

long long int XLCellValue::GetInteger() const {

    return m_value->GetInteger();
}

bool XLCellValue::GetBoolean() const {

    return m_value->GetBoolean();
}

long double XLCellValue::GetFloat() const {

    return m_value->GetFloat();
}

const char* XLCellValue::GetString() const {

    return m_value->GetString();
}

