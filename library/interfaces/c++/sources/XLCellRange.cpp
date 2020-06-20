//
// Created by Troldal on 2019-01-21.
//

#include <XLCellRange.hpp>

#include "XLCellRange.hpp"
#include "XLCellRange_Impl.hpp"

using namespace OpenXLSX;

XLCellRange::XLCellRange(Impl::XLCellRange range)
        : m_cellrange(std::make_unique<Impl::XLCellRange>(range)),
          m_cells(nullptr) {

}

XLCellRange::XLCellRange(const XLCellRange& other)
        : m_cellrange(std::make_unique<Impl::XLCellRange>(*other.m_cellrange)),
          m_cells(nullptr) {

}

XLCellRange::XLCellRange(XLCellRange&& other) = default;

XLCellRange::~XLCellRange() = default;

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

    m_cells = nullptr;
    m_cellrange->Transpose(state);
}

XLCellIterator XLCellRange::begin() {

    if (!m_cells)
        InitCells();
    return m_cells->begin();
}

XLCellIteratorConst XLCellRange::begin() const {

    if (!m_cells)
        InitCells();
    return m_cells->begin();
}

XLCellIterator XLCellRange::end() {

    if (!m_cells)
        InitCells();
    return m_cells->end();
}

XLCellIteratorConst XLCellRange::end() const {

    if (!m_cells)
        InitCells();
    return m_cells->end();
}

void XLCellRange::Clear() {

    m_cellrange->Clear();
}

void XLCellRange::InitCells() const {

    m_cells = std::make_unique<std::vector<XLCell>>();

    for (unsigned long row = 1; row <= m_cellrange->NumRows(); row++) {
        for (unsigned int column = 1; column <= m_cellrange->NumColumns(); column++) {
            m_cells->push_back(Cell(row, column));
        }
    }
}
