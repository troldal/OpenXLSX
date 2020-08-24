//
// Created by Kenneth Balslev on 24/08/2020.
//

#ifndef OPENXLSX_XLROWDATAITERATOR_HPP
#define OPENXLSX_XLROWDATAITERATOR_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <algorithm>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLCell.hpp"
#include "XLIterator.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    class XLRowDataRange;

    class OPENXLSX_EXPORT XLRowDataIterator
    {
    public:
        using iterator_category = std::forward_iterator_tag;
        using value_type        = XLCell;
        using difference_type   = int64_t;
        using pointer           = XLCell*;
        using reference         = XLCell&;

        /**
         * @brief
         * @param cellRange
         * @param loc
         */
        explicit XLRowDataIterator(const XLRowDataRange& rowDataRange, XLIteratorLocation loc);

        /**
         * @brief
         */
        ~XLRowDataIterator();

        /**
         * @brief
         * @param other
         */
        XLRowDataIterator(const XLRowDataIterator& other);

        /**
         * @brief
         * @param other
         */
        XLRowDataIterator(XLRowDataIterator&& other) noexcept;

        /**
         * @brief
         * @param other
         * @return
         */
        XLRowDataIterator& operator=(const XLRowDataIterator& other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLRowDataIterator& operator=(XLRowDataIterator&& other) noexcept;

        /**
         * @brief
         * @return
         */
        XLRowDataIterator& operator++();

        /**
         * @brief
         * @return
         */
        XLRowDataIterator operator++(int);    // NOLINT

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
        bool operator==(const XLRowDataIterator& rhs);

        /**
         * @brief
         * @param rhs
         * @return
         */
        bool operator!=(const XLRowDataIterator& rhs);

    private:
        std::unique_ptr<XMLNode> m_rowNode;                  /**< */
        uint16_t                 m_firstCol { 1 };            /**< The cell reference of the first cell in the range */
        uint16_t                 m_lastCol { 16384 };       /**< The cell reference of the last cell in the range */
        XLCell                   m_currentCell;                /**< */
        XLSharedStrings*         m_sharedStrings { nullptr }; /**< */
    };

}

#pragma warning(pop)
#endif    // OPENXLSX_XLROWDATAITERATOR_HPP
