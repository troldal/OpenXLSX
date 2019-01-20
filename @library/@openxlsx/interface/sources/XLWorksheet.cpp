//
// Created by Troldal on 2019-01-16.
//

#include "XLWorksheet.h"
#include "XLWorksheet_Impl.h"

using namespace OpenXLSX;

XLWorksheet::XLWorksheet(Impl::XLSheet& sheet) : XLSheet(sheet) {
}

unsigned int XLWorksheet::ColumnCount() const noexcept {

    return dynamic_cast<Impl::XLWorksheet*>(m_sheet)->ColumnCount();
}

unsigned long XLWorksheet::RowCount() const noexcept {

    return dynamic_cast<Impl::XLWorksheet*>(m_sheet)->RowCount();
}

void XLWorksheet::Export(const std::string& fileName,
                         char decimal,
                         char delimiter) {

    dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Export(fileName,
                                                      decimal,
                                                      delimiter);

}

void XLWorksheet::Import(const std::string& fileName,
                         const std::string& delimiter) {

    dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Import(fileName,
                                                      delimiter);

}
