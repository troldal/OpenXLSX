//
// Created by Troldal on 2019-01-25.
//

#include "XLColumn.hpp"
#include "XLColumn_Impl.hpp"

using namespace OpenXLSX;

OpenXLSX::XLColumn::XLColumn(Impl::XLColumn& column)
        : m_column(&column) {

}

float OpenXLSX::XLColumn::Width() const {

    return m_column->Width();
}

void OpenXLSX::XLColumn::SetWidth(float width) {

    m_column->SetWidth(width);
}

bool OpenXLSX::XLColumn::IsHidden() const {

    return m_column->IsHidden();
}

void OpenXLSX::XLColumn::SetHidden(bool state) {

    m_column->SetHidden(state);
}
