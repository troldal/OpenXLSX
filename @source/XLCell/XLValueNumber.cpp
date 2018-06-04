//
// Created by Kenneth Balslev on 29/12/2017.
//

#include "XLValueNumber.h"
#include "XLCell.h"
#include "XLCellValue.h"
#include "../XLSheet/XLWorksheet.h"
#include "../XLWorkbook/XLException.h"

using namespace OpenXLSX;
using namespace std;

/**
 * @details
 */
XLValueNumber::XLValueNumber()
    : XLValue(),
      m_numberType(XLNumberType::Integer)
{
    Set(0LL);
}

/**
 * @details
 */
const string& XLValueNumber::Set(long long int numberValue)
{
    m_numberType = XLNumberType::Integer;
    SetValueString(to_string(numberValue));
    return ValueString();
}

/**
 * @details
 */
const string& XLValueNumber::Set(long double numberValue)
{
    m_numberType = XLNumberType::Float;
    SetValueString(to_string(numberValue));
    return ValueString();
}

/**
 * @details
 */
long long int XLValueNumber::Integer() const
{
    if (m_numberType != XLNumberType::Integer) throw XLException("Cell value is not Integer");
    return stoll(ValueString());
}

/**
 * @details
 */
long double XLValueNumber::Float() const
{
    if (m_numberType != XLNumberType::Float) throw XLException("Cell value is not Float");
    return stold(ValueString());
}

/**
 * @details
 */
long long int XLValueNumber::AsInteger() const
{
    return stoll(ValueString());
}

/**
 * @details
 */
long double XLValueNumber::AsFloat() const
{
    return stold(ValueString());
}

/**
 * @details
 */
XLNumberType XLValueNumber::NumberType() const
{
    return m_numberType;
}

XLValueType XLValueNumber::ValueType() const
{
    switch (m_numberType) {
        case XLNumberType::Integer:
            return XLValueType::Integer;

        default:
            return XLValueType::Float;
    }
}

/**
 * @details
 */
XLCellType XLValueNumber::CellType() const
{
    return XLCellType::Number;
}

string XLValueNumber::TypeString() const
{
    return "";
}
/**
 * @details
 */
XLNumberType XLValueNumber::DetermineNumberType(const std::string &numberString) const
{
    if (numberString.find('.') != string::npos) return XLNumberType::Float;
    else return XLNumberType::Integer;
}

/**
 * @details
 */
string XLValueNumber::AsString() const
{
    return ValueString();
}
