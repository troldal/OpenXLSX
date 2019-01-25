//
// Created by Troldal on 2019-01-22.
//

#ifndef OPENXLSX_XLCELLITERATOR_H
#define OPENXLSX_XLCELLITERATOR_H

#include "XLCell.h"

namespace OpenXLSX {
    namespace Impl {
        class XLCellIterator;
        class XLCellIteratorConst;
    }

    class XLCellIterator
    {
    public:
        explicit XLCellIterator(Impl::XLCellIterator iterator);

        virtual ~XLCellIterator();

        const XLCellIterator& operator++();

        bool operator!=(const XLCellIterator& other) const;

        XLCell operator*() const;

    private:
        std::unique_ptr<Impl::XLCellIterator> m_iterator;

    };

    class XLCellIteratorConst
    {
    public:
/*
        explicit XLCellIteratorConst(const Impl::XLCellIteratorConst iterator);

        virtual ~XLCellIteratorConst();

        const XLCellIteratorConst& operator++();

        bool operator!=(const XLCellIteratorConst& other) const;

        const XLCell operator*() const;
*/

    private:
        //const std::unique_ptr<Impl::XLCellIteratorConst> m_iterator;

    };
}



#endif //OPENXLSX_XLCELLITERATOR_H
