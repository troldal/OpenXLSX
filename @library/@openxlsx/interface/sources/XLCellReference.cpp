//
// Created by Troldal on 2019-01-20.
//

#include "XLCellReference.h"
#include "XLCellReference_Impl.h"

using namespace OpenXLSX;

XLCellReference::XLCellReference(const std::string& cellAddress) : m_cellReference(Impl::XLCellReference(cellAddress)) {

}

