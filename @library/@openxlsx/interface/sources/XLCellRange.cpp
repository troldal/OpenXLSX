//
// Created by Troldal on 2019-01-21.
//

#include <XLCellRange.h>

#include "XLCellRange.h"
#include "XLCellRange_Impl.h"

using namespace OpenXLSX;

XLCellRange::XLCellRange(Impl::XLCellRange range)
        : m_cellrange(std::make_unique<Impl::XLCellRange>(range)){

}

XLCellRange::~XLCellRange() {

}

XLCell XLCellRange::Cell(unsigned long row, unsigned int column) {

    return XLCell(*m_cellrange->Cell(row, column));
}

const XLCell XLCellRange::Cell(unsigned long row, unsigned int column) const {

    return XLCell(*m_cellrange->Cell(row, column));
}

unsigned long XLCellRange::NumRows() const {

    return m_cellrange->NumRows();
}

unsigned int XLCellRange::NumColumns() const {

    return m_cellrange->NumColumns();
}

void XLCellRange::Transpose(bool state) const {

    m_cellrange->Transpose(state);
}

XLCellIterator XLCellRange::begin() {

    return XLCellIterator(m_cellrange->begin());
}

XLCellIterator XLCellRange::end() {

    return XLCellIterator(m_cellrange->end());
}

void XLCellRange::Clear() {

    m_cellrange->Clear();
}
