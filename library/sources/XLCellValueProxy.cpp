//
// Created by Kenneth Balslev on 23/08/2020.
//

#include "XLCellValueProxy.hpp"
#include "XLCell.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
XLCellValueProxy::XLCellValueProxy(XLCell* cell) : m_cell(cell) {}
XLCellValueProxy::~XLCellValueProxy()                                 = default;
XLCellValueProxy::XLCellValueProxy(const XLCellValueProxy& other)     = default;
XLCellValueProxy::XLCellValueProxy(XLCellValueProxy&& other) noexcept = default;
XLCellValueProxy& XLCellValueProxy::operator                          =(const XLCellValueProxy& other)
{
    if (&other != this) {
        *this = other.m_cell->getValue();
    }

    return *this;
}
XLCellValueProxy& XLCellValueProxy::operator=(XLCellValueProxy&& other) noexcept = default;

XLCellValueProxy::operator XLCellValue()
{
    return m_cell->getValue();
}

void XLCellValueProxy::setInteger(int64_t numberValue)
{
    m_cell->setInteger(numberValue);
}
void XLCellValueProxy::setBoolean(bool numberValue)
{
    m_cell->setBoolean(numberValue);
}
void XLCellValueProxy::setFloat(double numberValue)
{
    m_cell->setFloat(numberValue);
}
void XLCellValueProxy::setString(const char* stringValue)
{
    m_cell->setString(stringValue);
}
XLCellValue XLCellValueProxy::getValue() const
{
    return m_cell->getValue();
}

/**
 * @details
 */
XLCellValueProxy& XLCell::value()
{
    return m_valueProxy;
}

const XLCellValueProxy& XLCell::value() const
{
    return m_valueProxy;
}