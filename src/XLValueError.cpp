//
// Created by Kenneth Balslev on 04/01/2018.
//

#include "XLValueError.h"

using namespace OpenXLSX;
using namespace std;

/**
 * @details
 */
XLValueError::XLValueError(XLCellValue &parent)
    : XLValue(parent)
{

}

/**
 * @details
 */
std::unique_ptr<XLValue> XLValueError::Clone(XLCell &parent)
{
    unique_ptr<XLValue> result(new XLValueError(ParentCellValue()));
    *result = *this;

    return result;
}

/**
 * @details
 */
XLValueType XLValueError::ValueType() const
{
    return XLValueType::Error;
}

/**
 * @details
 */
XLCellType XLValueError::CellType() const
{
    return XLCellType::Error;
}

/**
 * @details
 */
std::string XLValueError::TypeString() const
{
    return "e";
}

/**
 * @details
 */
std::string XLValueError::AsString() const
{
    return "Error";
}
