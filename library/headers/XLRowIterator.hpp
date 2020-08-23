//
// Created by Kenneth Balslev on 22/08/2020.
//

#ifndef OPENXLSX_XLROWITERATOR_HPP
#define OPENXLSX_XLROWITERATOR_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

#include <algorithm>

#include "OpenXLSX-Exports.hpp"
#include "XLIterator.hpp"
#include "XLRow.hpp"
//#include "XLRowRange.hpp"
//#include "XLCellReference.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    class XLRowRange;

    class OPENXLSX_EXPORT XLRowIterator
    {
    public:
        using iterator_category = std::forward_iterator_tag;
        using value_type        = XLRow;
        using difference_type   = int64_t;
        using pointer           = XLRow*;
        using reference         = XLRow&;

        /**
         * @brief
         * @param cellRange
         * @param loc
         */
        explicit XLRowIterator(const XLRowRange& rowRange, XLIteratorLocation loc);

        /**
         * @brief
         */
        ~XLRowIterator();

        /**
         * @brief
         * @param other
         */
        XLRowIterator(const XLRowIterator& other);

        /**
         * @brief
         * @param other
         */
        XLRowIterator(XLRowIterator&& other) noexcept;

        /**
         * @brief
         * @param other
         * @return
         */
        XLRowIterator& operator=(const XLRowIterator& other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLRowIterator& operator=(XLRowIterator&& other) noexcept;

        /**
         * @brief
         * @return
         */
        XLRowIterator& operator++();

        /**
         * @brief
         * @return
         */
        XLRowIterator operator++(int);    // NOLINT

        /**
         * @brief
         * @return
         */
        reference operator*();

        /**
         * @brief
         * @return
         */
        pointer operator->();

        /**
         * @brief
         * @param rhs
         * @return
         */
        bool operator==(const XLRowIterator& rhs);

        /**
         * @brief
         * @param rhs
         * @return
         */
        bool operator!=(const XLRowIterator& rhs);

    private:
        std::unique_ptr<XMLNode> m_dataNode;                  /**< */
        uint32_t                 m_firstRow { 1 };            /**< The cell reference of the first cell in the range */
        uint32_t                 m_lastRow { 1048576 };       /**< The cell reference of the last cell in the range */
        XLRow                    m_currentRow;                /**< */
        XLSharedStrings*         m_sharedStrings { nullptr }; /**< */
    };

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLROWITERATOR_HPP
