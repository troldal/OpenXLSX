//
// Created by Troldal on 2019-01-21.
//

#ifndef OPENXLSX_XLCELLRANGE_H
#define OPENXLSX_XLCELLRANGE_H

#include "XLCell.h"
#include "XLCellIterator.h"

namespace OpenXLSX {
    namespace Impl {
        class XLCellRange;
    }

    class XLCellRange
    {
    public:
        explicit XLCellRange(Impl::XLCellRange range);

        XLCellRange(const XLCellRange& other) = default;

        XLCellRange(XLCellRange&& other) = default;

        virtual ~XLCellRange();

        XLCell Cell(unsigned long row,
                     unsigned int column);

        const XLCell Cell(unsigned long row,
                           unsigned int column) const;

        unsigned long NumRows() const;

        unsigned int NumColumns() const;

        void Transpose(bool state) const;

        XLCellIterator begin();

//        XLCellIteratorConst begin() const;

        XLCellIterator end();

//        XLCellIteratorConst end() const;

        void Clear();


    private:
        std::unique_ptr<Impl::XLCellRange> m_cellrange;
    };
}



#endif //OPENXLSX_XLCELLRANGE_H
