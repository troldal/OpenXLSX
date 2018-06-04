//
// Created by Kenneth Balslev on 31/12/2017.
//

#include "XLValueBoolean.h"
#include "XLCell.h"
#include "XLCellValue.h"

using namespace OpenXLSX;
using namespace std;

/**
 * @details The 'default' constructor, creates an XLBoolean object and initializes it to a value of False.
 */
XLValueBoolean::XLValueBoolean()
    : XLValue()
{
    Set(false);
}

/**
 * @details This function sets the cell node in the underlying XML file to the right value, based on the input. A value
 * node and a type attribute is created (if they don't exists already) and set to the appropriate values.
 */
const std::string& XLValueBoolean::Set(bool boolValue)
{
    if (boolValue) SetValueString("1");
    else SetValueString("0");
    return ValueString();
}

/**
 * @details This function returns an XLBool object, based on the value in the underlying XML file.
 */
bool XLValueBoolean::Get() const
{
    if (ValueString() == "1")
        return true;
    else
        return false;
}

/**
 * @details This function returns XLValueType::Boolean.
 */
XLValueType XLValueBoolean::ValueType() const
{
    return XLValueType::Boolean;
}

/**
 * @details This function returns XLCellType::Boolean.
 */
XLCellType XLValueBoolean::CellType() const
{
    return XLCellType::Boolean;
}

/**
 * @details A std::string "b" is returned, which is the value of the type attribute for boolean values.
 */
string XLValueBoolean::TypeString() const
{
    return "b";
}

/**
 * @details This function returns "TRUE" if the value is "1"; otherwise it returns "FALSE".
 */
std::string XLValueBoolean::AsString() const
{
    if (ValueString() == "1")
        return "TRUE";
    else
        return "FALSE";
}

