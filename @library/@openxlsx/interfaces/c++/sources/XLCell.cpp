//
// Created by Troldal on 2019-01-21.
//

#include <XLCell.h>

#include "XLCell.h"
#include "XLCell_Impl.h"

using namespace OpenXLSX;

XLCell::XLCell(Impl::XLCell& cell)
        : m_cell(&cell) {
}

XLValueType XLCell::ValueType() const {

    return m_cell->ValueType();
}

const XLCellReference XLCell::CellReference() const {

    return XLCellReference(*m_cell->CellReference());
}

XLCellValue XLCell::Value() {

    return XLCellValue(m_cell->Value());
}

const XLCellValue XLCell::Value() const {

    return XLCellValue(m_cell->Value());
}
