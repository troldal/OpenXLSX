//
// Created by Kenneth Balslev on 31/12/2017.
//

#include "XLValue.h"
#include "XLCellValue.h"

using namespace OpenXLSX;
using namespace std;

/**
 * @details The constructor initializes the m_parentCellValue member variable with the value of the parent parameter.
 */
XLValue::XLValue() : m_valueString()
{
}

std::string XLValue::ValueString() const
{
    return m_valueString;
}

void XLValue::SetValueString(const std::string &value)
{
    m_valueString = value;
}
