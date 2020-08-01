//
// Created by Troldal on 2019-01-20.
//

#include <XLCellReference.hpp>

#include "XLCellReference.hpp"
#include "XLCellReference_Impl.hpp"

using namespace OpenXLSX;
using namespace std;

XLCellReference::XLCellReference(const Impl::XLCellReference& sheet)
        : m_cellReference(make_unique<Impl::XLCellReference>(sheet)) {

}

XLCellReference::XLCellReference(const std::string& cellAddress)
        : m_cellReference(make_unique<Impl::XLCellReference>(Impl::XLCellReference(cellAddress))) {

}

XLCellReference::XLCellReference(unsigned long row, unsigned int column)
        : m_cellReference(make_unique<Impl::XLCellReference>(Impl::XLCellReference(row, column))) {

}

XLCellReference::XLCellReference(const XLCellReference& other)
        : m_cellReference(make_unique<Impl::XLCellReference>(*other.m_cellReference)) {

}

XLCellReference::XLCellReference(XLCellReference&& other) = default;

XLCellReference& XLCellReference::operator=(const XLCellReference& other) {

    *m_cellReference = *other.m_cellReference;
    return *this;
}

XLCellReference& XLCellReference::operator=(XLCellReference&& other) = default;

XLCellReference::~XLCellReference() = default;

unsigned long XLCellReference::Row() {

    return m_cellReference->Row();
}

void XLCellReference::SetRow(unsigned long row) {

    m_cellReference->SetRow(row);
}

unsigned int XLCellReference::Column() {

    return m_cellReference->Column();
}

void XLCellReference::SetColumn(unsigned int column) {

    m_cellReference->SetColumn(column);
}

void XLCellReference::SetRowAndColumn(unsigned long row, unsigned int column) {

    m_cellReference->SetRowAndColumn(row, column);
}

std::string XLCellReference::Address() const {

    return m_cellReference->Address();
}

void XLCellReference::SetAddress(const std::string& address) {

    m_cellReference->SetAddress(address);
}

