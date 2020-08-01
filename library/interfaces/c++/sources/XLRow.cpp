//
// Created by Troldal on 2019-01-23.
//

#include <XLRow.hpp>

#include "XLRow.hpp"
#include "XLRow_Impl.hpp"

using namespace OpenXLSX;

XLRow::XLRow(Impl::XLRow& row)
        : m_row(&row) {

}

float XLRow::Height() const {

    return m_row->Height();
}

void XLRow::SetHeight(float height) {

    m_row->SetHeight(height);
}

float XLRow::Descent() const {

    return m_row->Descent();
}

void XLRow::SetDescent(float descent) {

    m_row->SetDescent(descent);
}

bool XLRow::IsHidden() const {

    return m_row->IsHidden();
}

void XLRow::SetHidden(bool state) {

    m_row->SetHidden(state);
}

XLCell XLRow::Cell(unsigned int column) {

    return XLCell(*m_row->Cell(column));
}

const XLCell XLRow::Cell(unsigned int column) const {

    return XLCell(*m_row->Cell(column));
}

unsigned int XLRow::CellCount() const {

    return m_row->CellCount();
}
