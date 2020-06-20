//
// Created by Troldal on 2019-01-16.
//

#include "XLSheet.hpp"
#include "XLSheet_Impl.hpp"

using namespace OpenXLSX;

XLSheet::XLSheet(Impl::XLSheet& sheet)
        : m_sheet(&sheet) {
}

std::string const XLSheet::Name() const {

    return m_sheet->Name();
}

void XLSheet::SetName(const std::string& name) {

    m_sheet->SetName(name);
}

const XLSheetState& XLSheet::State() const {

    return m_sheet->State();
}

void XLSheet::SetState(XLSheetState state) {

    m_sheet->SetState(state);
}

const XLSheetType& XLSheet::Type() const {

    return m_sheet->Type();
}

unsigned int XLSheet::Index() const {

    return m_sheet->Index();
}

void XLSheet::SetIndex(int index) {

    m_sheet->SetIndex();
}
