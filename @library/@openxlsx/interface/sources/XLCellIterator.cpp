//
// Created by Troldal on 2019-01-22.
//

#include "XLCellIterator.h"
#include "XLCellIterator_Impl.h"

using namespace OpenXLSX;

XLCellIterator::XLCellIterator(Impl::XLCellIterator iterator)
        : m_iterator(std::make_unique<Impl::XLCellIterator>(iterator)) {

}

const XLCellIterator& XLCellIterator::operator++() {

    auto a = m_iterator.get();
    auto b = *a;
    b.operator++();
    return *this;
}

bool XLCellIterator::operator!=(const XLCellIterator& other) const {

    return (*m_iterator != *other.m_iterator);
}

XLCell XLCellIterator::operator*() const {

    auto a = m_iterator.get();
    auto b = *a;
    return XLCell(b.operator*());
}

XLCellIterator::~XLCellIterator() {

}
