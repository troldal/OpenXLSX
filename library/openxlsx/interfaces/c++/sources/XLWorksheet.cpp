//
// Created by Troldal on 2019-01-16.
//

#include "XLWorksheet.hpp"
#include "XLWorksheet_Impl.hpp"
#include "XLCellRange_Impl.hpp"

using namespace OpenXLSX;

XLWorksheet::XLWorksheet(Impl::XLSheet& sheet)
        : XLSheet(sheet) {
}

XLCell XLWorksheet::Cell(const XLCellReference& ref) {

    return XLCell(*dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Cell(ref.Address()));
}

const XLCell XLWorksheet::Cell(const XLCellReference& ref) const {

    return XLCell(*dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Cell(ref.Address()));
}

XLCell XLWorksheet::Cell(const std::string& address) {

    return XLCell(*dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Cell(address));
}

const XLCell XLWorksheet::Cell(const std::string& address) const {

    return XLCell(*dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Cell(address));
}

XLCell XLWorksheet::Cell(unsigned long rowNumber, unsigned int columnNumber) {

    return XLCell(*dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Cell(rowNumber, columnNumber));
}

const XLCell XLWorksheet::Cell(unsigned long rowNumber, unsigned int columnNumber) const {

    return XLCell(*dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Cell(rowNumber, columnNumber));
}

XLCellReference XLWorksheet::FirstCell() const noexcept {

    return XLCellReference(dynamic_cast<Impl::XLWorksheet*>(m_sheet)->FirstCell());
}

XLCellReference XLWorksheet::LastCell() const noexcept {

    return XLCellReference(dynamic_cast<Impl::XLWorksheet*>(m_sheet)->LastCell());
}

XLCellRange XLWorksheet::Range() {

    return XLCellRange(dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Range());
}

const XLCellRange XLWorksheet::Range() const {

    return XLCellRange(dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Range());
}

XLCellRange XLWorksheet::Range(const XLCellReference& topLeft, const XLCellReference& bottomRight) {

    return XLCellRange(dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Range(Impl::XLCellReference(topLeft.Address()),
                                                                        Impl::XLCellReference(bottomRight.Address())));
}

const XLCellRange XLWorksheet::Range(const XLCellReference& topLeft, const XLCellReference& bottomRight) const {

    return XLCellRange(dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Range(Impl::XLCellReference(topLeft.Address()),
                                                                        Impl::XLCellReference(bottomRight.Address())));
}

XLRow XLWorksheet::Row(unsigned long rowNumber) {

    return XLRow(*dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Row(rowNumber));
}

const XLRow XLWorksheet::Row(unsigned long rowNumber) const {

    return XLRow(*dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Row(rowNumber));
}

XLColumn XLWorksheet::Column(unsigned int columnNumber) {

    return XLColumn(*dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Column(columnNumber));
}

const XLColumn XLWorksheet::Column(unsigned int columnNumber) const {

    return XLColumn(*dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Column(columnNumber));
}

unsigned int XLWorksheet::ColumnCount() const noexcept {

    return dynamic_cast<Impl::XLWorksheet*>(m_sheet)->ColumnCount();
}

unsigned long XLWorksheet::RowCount() const noexcept {

    return dynamic_cast<Impl::XLWorksheet*>(m_sheet)->RowCount();
}

void XLWorksheet::Export(const std::string& fileName, char decimal, char delimiter) {

    dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Export(fileName, decimal, delimiter);

}

void XLWorksheet::Import(const std::string& fileName, const std::string& delimiter) {

    dynamic_cast<Impl::XLWorksheet*>(m_sheet)->Import(fileName, delimiter);

}
