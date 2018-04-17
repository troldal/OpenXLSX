//
// Created by Kenneth Balslev on 29/12/2017.
//

#include "XLValueNumber.h"
#include "XLCell.h"
#include "XLCellValue.h"
#include "XLWorksheet.h"
#include "XLException.h"

using namespace OpenXLSX;
using namespace std;

/**
 * @details
 */
XLValueNumber::XLValueNumber(XLCellValue &parent)
    : XLValue(parent),
      m_numberType(XLNumberType::Integer),
      m_integer(0LL),
      m_float(0.0)
{
    ParentCellValue()->SetValueNode("0");
    ParentCellValue()->SetTypeAttribute(TypeString());

    m_numberType = DetermineNumberType(ParentCellValue()->ValueNode()->Value());

    switch (m_numberType) {
        case XLNumberType::Integer:
            m_integer = stoll(ParentCellValue()->ValueNode()->Value());
            break;

        case XLNumberType::Float:
            m_float = stold(ParentCellValue()->ValueNode()->Value());
            break;
    }
}

/**
 * @details
 */
XLValueNumber::XLValueNumber(long long int numberValue, XLCellValue &parent)
    : XLValue(parent),
      m_numberType(XLNumberType::Integer),
      m_integer(0LL),
      m_float(0.0)
{
    Set(numberValue);
}

/**
 * @details
 */
XLValueNumber::XLValueNumber(long double numberValue, XLCellValue &parent)
    : XLValue(parent),
      m_numberType(XLNumberType::Float),
      m_integer(0LL),
      m_float(0.0)
{
    Set(numberValue);
}

/**
 * @details
 */
XLValueNumber &XLValueNumber::operator=(const XLValueNumber &other)
{
    m_numberType = other.m_numberType;
    m_integer = other.m_integer;
    m_float = other.m_float;
    ParentCellValue()->SetValueNode(AsString());
    ParentCellValue()->ParentCell()->SetTypeAttribute(TypeString());
    ParentCellValue()->ParentCell()->SetModified();
    return *this;
}

/**
 * @details
 */
XLValueNumber &XLValueNumber::operator=(XLValueNumber &&other) noexcept
{
    m_numberType = other.m_numberType;
    m_integer = other.m_integer;
    m_float = other.m_float;
    ParentCellValue()->ValueNode()->SetValue(AsString());
    ParentCellValue()->ParentCell()->SetTypeAttribute(TypeString());
    ParentCellValue()->ParentCell()->SetModified();
    return *this;
}

/**
 * @details
 */
XLValueNumber &XLValueNumber::operator=(long long int numberValue)
{
    Set(numberValue);
    return *this;
}

/**
 * @details
 */
XLValueNumber &XLValueNumber::operator=(long double numberValue)
{
    Set(numberValue);
    return *this;
}

/**
 * @details
 */
void XLValueNumber::Set(long long int numberValue)
{
    m_numberType = XLNumberType::Integer;
    m_integer = numberValue;
    m_float = 0.0;
    ParentCellValue()->SetValueNode(AsString());
    ParentCellValue()->SetTypeAttribute(TypeString());
    ParentCellValue()->ParentCell()->SetModified();

}

/**
 * @details
 */
void XLValueNumber::Set(long double numberValue)
{
    m_numberType = XLNumberType::Float;
    m_integer = 0;
    m_float = numberValue;
    ParentCellValue()->SetValueNode(AsString());
    ParentCellValue()->SetTypeAttribute(TypeString());
    ParentCellValue()->ParentCell()->SetModified();
}

/**
 * @details
 */
long long int XLValueNumber::Integer() const
{
    if (m_numberType != XLNumberType::Integer) throw XLException("Cell value is not Integer");
    return m_integer;
}

/**
 * @details
 */
long double XLValueNumber::Float() const
{
    if (m_numberType != XLNumberType::Float) throw XLException("Cell value is not Float");
    return m_float;
}

/**
 * @details
 */
long long int XLValueNumber::AsInteger() const
{
    switch (m_numberType) {
        case XLNumberType::Float:
            return static_cast<long long int>(m_float);

        case XLNumberType::Integer:
            return m_integer;
    }

    return 0; // Non-reachable code; included only to silence compiler warning.
}

/**
 * @details
 */
long double XLValueNumber::AsFloat() const
{
    switch (m_numberType) {
        case XLNumberType::Float:
            return m_float;

        case XLNumberType::Integer:
            return static_cast<long double>(m_integer);
    }

    return 0.0; // Non-reachable code; included only to silence compiler warning.
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
    bool hasDecimalPoint = (numberString.find('.') != 0);

    if (hasDecimalPoint) return XLNumberType::Float;
    else return XLNumberType::Integer;
}

/**
 * @details
 */
string XLValueNumber::AsString() const
{
    switch (NumberType()) {
        case XLNumberType::Integer:
            return to_string(m_integer);

        case XLNumberType::Float:
            return to_string(m_float);
    }

    return ""; // Non-reachable code; included only to silence compiler warning.
}
unique_ptr<XLValue> XLValueNumber::Clone(XLCell &parent)
{
    unique_ptr<XLValue> result(new XLValueNumber(*ParentCellValue()));
    *result = *this;

    return result;
}
