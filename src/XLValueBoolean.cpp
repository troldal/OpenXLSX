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
XLValueBoolean::XLValueBoolean(XLCellValue &parent)
    : XLValue(parent)
{
    if (!ParentCellValue()->ValueNode()) {
        ParentCellValue()->SetValueNode("0");
        ParentCellValue()->SetTypeAttribute(TypeString());
    }

    if (string(ParentCellValue()->ValueNode().text().get()) == "0") Set(XLBool::False);
    else Set(XLBool::True);
}

/**
 * @details An XLBoolean object is constructed, based on the XLBool input to the constructor.
 */
XLValueBoolean::XLValueBoolean(XLBool boolValue, XLCellValue &parent)
    : XLValue(parent)
{
    Set(boolValue);
}

/**
 * @details An XLBoolean object is constructed, based on the int input to the constructor; False if input == 0,
 * otherwise True.
 */
XLValueBoolean::XLValueBoolean(unsigned int boolValue, XLCellValue &parent)
    : XLValue(parent)
{
    if (boolValue == 1)
        Set(XLBool::True);
    else
        Set(XLBool::False);
}

/**
 * @details An XLBoolean object is constructed, based on the bool input to the constructor.
 */
XLValueBoolean::XLValueBoolean(bool boolValue, XLCellValue &parent)
    : XLValue(parent)
{
    if (boolValue)
        Set(XLBool::True);
    else
        Set(XLBool::False);
}

/**
 * @details This function sets the cell node in the underlying XML file to the right value, based on the input. A value
 * node and a type attribute is created (if they don't exists already) and set to the appropriate values.
 */
void XLValueBoolean::Set(XLBool boolValue)
{
    if (!ParentCellValue()->HasValueNode()) ParentCellValue()->CreateValueNode();
    if (!ParentCellValue()->HasTypeAttribute()) ParentCellValue()->CreateTypeAttribute();

    ParentCellValue()->SetTypeAttribute(TypeString());
    ParentCellValue()->ParentCell()->SetModified();

    switch (boolValue) {
        case XLBool::True:
            ParentCellValue()->ValueNode().text().set("1");
            break;

        case XLBool::False:
            ParentCellValue()->ValueNode().text().set("0");
            break;
    }
}

/**
 * @details This function returns an XLBool object, based on the value in the underlying XML file.
 */
XLBool XLValueBoolean::Boolean() const
{
    if (string(ParentCellValue()->ValueNode().text().get()) == "1")
        return XLBool::True;
    else
        return XLBool::False;
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
 * @details This function retruns "True" if the value is XLBool::True; otherwise it returns "False".
 */
std::string XLValueBoolean::AsString() const
{
    if (Boolean() == XLBool::True)
        return "True";
    else
        return "False";
}

/**
 * @details A new unique_ptr with a 'default' constructed XLBoolean object is created. Subsequently, the value of the
 * current object (*this) is copy assigned to the new unique_ptr object. The result is returned.
 * @todo Is this the right way to do this? Is there a more elegant way?
 */
std::unique_ptr<XLValue> XLValueBoolean::Clone(XLCell &parent)
{
    unique_ptr<XLValue> result(new XLValueBoolean(*ParentCellValue()));
    *result = *this;

    return result;
}
