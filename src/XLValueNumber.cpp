//
// Created by Kenneth Balslev on 29/12/2017.
//

#include "XLValueNumber.h"
#include "XLCell.h"
#include "XLCellValue.h"
#include "XLWorksheet.h"
#include "XLException.h"

using namespace RapidXLSX;
using namespace std;
using namespace boost;

/**
 * @details
 */
XLValueNumber::XLValueNumber(XLCellValue &parent)
    : XLValue(parent),
      m_number(0LL),
      m_numberType(XLNumberType::Integer)
{
    ParentCellValue().SetValueNode("0");
    ParentCellValue().SetTypeAttribute(TypeString());

    m_numberType = DetermineNumberType(ParentCellValue().ValueNode()->value());

    switch (m_numberType) {
        case XLNumberType::Integer:
            m_number = stoll(ParentCellValue().ValueNode()->value());
            break;

        case XLNumberType::Float:
            m_number = stold(ParentCellValue().ValueNode()->value());
            break;
    }
}

/**
 * @details
 */
XLValueNumber::XLValueNumber(long long int numberValue, XLCellValue &parent)
    : XLValue(parent),
      m_number(numberValue),
      m_numberType(XLNumberType::Integer)
{
    Set(numberValue);
}

/**
 * @details
 */
XLValueNumber::XLValueNumber(long double numberValue, XLCellValue &parent)
    : XLValue(parent),
      m_number(numberValue),
      m_numberType(XLNumberType::Float)
{
    Set(numberValue);
}

/**
 * @details
 */
XLValueNumber &XLValueNumber::operator=(const XLValueNumber &other)
{
    m_number = other.m_number;
    m_numberType = other.m_numberType;
    ParentCellValue().SetValueNode(AsString());
    ParentCellValue().ParentCell().SetTypeAttribute(TypeString());
    ParentCellValue().ParentCell().SetModified();
    return *this;
}

/**
 * @details
 */
XLValueNumber &XLValueNumber::operator=(XLValueNumber &&other) noexcept
{
    m_number = std::move(other.m_number);
    m_numberType = other.m_numberType;
    ParentCellValue().ValueNode()->setValue(AsString());
    ParentCellValue().ParentCell().SetTypeAttribute(TypeString());
    ParentCellValue().ParentCell().SetModified();
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
    m_number = numberValue;
    m_numberType = XLNumberType::Integer;
    ParentCellValue().SetValueNode(AsString());
    ParentCellValue().SetTypeAttribute(TypeString());
    ParentCellValue().ParentCell().SetModified();

}

/**
 * @details
 */
void XLValueNumber::Set(long double numberValue)
{
    m_number = numberValue;
    m_numberType = XLNumberType::Float;
    ParentCellValue().SetValueNode(AsString());
    ParentCellValue().SetTypeAttribute(TypeString());
    ParentCellValue().ParentCell().SetModified();
}

/**
 * @details
 */
long long int XLValueNumber::Integer() const
{
    if (m_numberType != XLNumberType::Integer) throw XLException("Cell value is not Integer");
    return get<long long int>(m_number);
}

/**
 * @details
 */
long double XLValueNumber::Float() const
{
    if (m_numberType != XLNumberType::Float) throw XLException("Cell value is not Float");
    return get<long double>(m_number);
}

/**
 * @details
 */
long long int XLValueNumber::AsInteger() const
{
    switch (m_numberType) {
        case XLNumberType::Float:
            return static_cast<long long int>(get<long double>(m_number));

        case XLNumberType::Integer:
            return get<long long int>(m_number);
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
            return get<long double>(m_number);

        case XLNumberType::Integer:
            return static_cast<long double>(get<long long int>(m_number));
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
            return to_string(get<long long int>(m_number));

        case XLNumberType::Float:
            return to_string(get<long double>(m_number));
    }

    return ""; // Non-reachable code; included only to silence compiler warning.
}
unique_ptr<XLValue> XLValueNumber::Clone(XLCell &parent)
{
    unique_ptr<XLValue> result(new XLValueNumber(ParentCellValue()));
    *result = *this;

    return result;
}
