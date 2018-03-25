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
XLValue::XLValue(XLCellValue &parent)
    : m_parentCellValue(parent)
{

}

/**
 * @details The copy constructor doesn't actually copy anything, as there are no member variables to copy (except for
 * the parent XLCellValue object, which should not be copied. Therefore, this function simply returns a reference to
 * itself.
 * @pre The object has been properly initialized.
 * @post The object and any associated objects are left unchanged.
 */
XLValue &XLValue::operator=(const XLValue &other)
{
    return *this;
}

/**
 * @details The move constructor doesn't actually move anything, as there are no member variables to move (except for
 * the parent XLCellValue object, which should not be moved. Therefore, this function simply returns a reference to
 * itself.
 * @pre The object has been properly initialized.
 * @post The object and any associated objects are left unchanged.
 */
XLValue &XLValue::operator=(XLValue &&other) noexcept
{
    return *this;
}

/**
 * @details Returns the m_parentCellValue member variable.
 * @pre The object has been properly initialized with a valid parent XLCellValue object.
 * @post The object and any associated objects are left unchanged.
 */
XLCellValue &XLValue::ParentCellValue()
{
    return m_parentCellValue;
}

/**
 * @details Returns the m_parentCellValue member variable.
 * @pre The object has been properly initialized with a valid parent XLCellValue object.
 * @post The object and any associated objects are left unchanged.
 */
const XLCellValue &XLValue::ParentCellValue() const
{
    return m_parentCellValue;
}
