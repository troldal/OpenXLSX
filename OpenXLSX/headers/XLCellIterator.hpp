/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.
 8b    d8   MM.  ,M9 YM    d9 MM    MM  d'    `MM.   MM    / L    ,M9  d'    `MM.
  YMMMM9    MMYMMM9   YMMMM9 _MM_  _MM_M(_    _)MM_ _MMMMMMM MYMMMM9 _M(_    _)MM_
            MM
            MM
           _MM_

  Copyright (c) 2018, Kenneth Troldal Balslev

  All rights reserved.

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  - Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.
  - Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.
  - Neither the name of the author nor the
    names of any contributors may be used to endorse or promote products
    derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

#ifndef OPENXLSX_XLCELLITERATOR_HPP
#define OPENXLSX_XLCELLITERATOR_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#   pragma warning(push)
#   pragma warning(disable : 4251)
#   pragma warning(disable : 4275)
#endif // _MSC_VER

#include <algorithm>

#include "OpenXLSX-Exports.hpp"
#include "XLCell.hpp"
#include "XLCellReference.hpp"
#include "XLIterator.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    class OPENXLSX_EXPORT XLCellIterator
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
         * @param colStyles is an optional vector with all column styles configured for the worksheet when the iterator is created.
         *                  This vector - if provided - will be used to get default cell styles for newly created cells, instead of performing
         *                   a potentially expensive lookup in <cols>
         */
        explicit XLCellIterator(const XLCellRange& cellRange, XLIteratorLocation loc, std::vector<XLStyleIndex> const * colStyles);

        /**
         * @brief
         */
        ~XLCellIterator();

        /**
         * @brief
         * @param other
         */
        XLCellIterator(const XLCellIterator& other);

        /**
         * @brief
         * @param other
         */
        [[maybe_unused]] XLCellIterator(XLCellIterator&& other) noexcept;

        /**
         * @brief
         * @param other
         * @return
         */
        XLCellIterator& operator=(const XLCellIterator& other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLCellIterator& operator=(XLCellIterator&& other) noexcept;

    private:    // ===== Switch to private method that is used by the XLIterator increment operator++ and the dereference operators * and ->
        static constexpr const bool XLCreateIfMissing      = true;     // code readability for updateCurrentCell parameter createIfMissing
        static constexpr const bool XLDoNotCreateIfMissing = false;    //   "
        /**
         * @brief update m_currentCell by fetching (or inserting) a cell at m_currentRow, m_currentColumn
         * @param createIfMissing m_currentCell will only be inserted if createIfMissing is true
         */
        void updateCurrentCell(bool createIfMissing);

    public:     // ===== Switch back to public methods
        /**
         * @brief
         * @return
         */
        XLCellIterator& operator++();

        /**
         * @brief
         * @return
         */
        XLCellIterator operator++(int);    // NOLINT

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
        bool operator==(const XLCellIterator& rhs) const;

        /**
         * @brief
         * @param rhs
         * @return
         */
        bool operator!=(const XLCellIterator& rhs) const;

        /**
         * @brief determine whether the cell that the iterator points to exists (cell at m_currentRow, m_currentColumn)
         * @return true if XML already has an entry for that cell, otherwise false
         */
        bool cellExists();

        /**
         * @brief determine whether iterator is at 1 beyond the last cell in range
         * @return
         * @note 28-07-2024: Removed const from return type (Troldal)
         */
        bool endReached() const { return m_endReached; }

        /**
         * @brief
         * @param last
         * @return
         */
        uint64_t distance(const XLCellIterator& last);

        /**
         * @brief get the XLCellReference::address corresponding to the current iterator position
         * @return an XLCellReference::address, with m_bottomRight.col() + 1 for the beyond-the-end iterator
         * @note 28-07-2024: Removed const from return type (Troldal)
         */
        std::string address() const;

    private:
        std::unique_ptr<XMLNode> m_dataNode;             /**< */
        XLCellReference          m_topLeft;              /**< The cell reference of the first cell in the range */
        XLCellReference          m_bottomRight;          /**< The cell reference of the last cell in the range */
        XLSharedStrings          m_sharedStrings;        /**< */
        bool                     m_endReached;           /**< */
        XMLNode                  m_hintNode;             /**< The cell node of the last existing cell found up to current iterator position */
        uint32_t                 m_hintRow;              /**<   the row number for m_hintCell */
        XLCell                   m_currentCell;          /**< The cell to which the iterator is currently pointing, if it exists, otherwise an empty XLCell */
        static constexpr const int XLNotLoaded  = 0;    // code readability for m_currentCellStatus
        static constexpr const int XLNoSuchCell = 1;    //   "
        static constexpr const int XLLoaded     = 2;    //   "
        int                      m_currentCellStatus;    /**< Status of m_currentCell: XLNotLoaded, XLNoSuchCell or XLLoaded */
        uint32_t                 m_currentRow;
        uint16_t                 m_currentColumn;
        std::vector<XLStyleIndex> const * m_colStyles;
    };

    /**
     * @brief      ostream output of XLIterator position as XLCellReference::address
     * @param os   the ostream destination
     * @param it    the XLIterator whose position to send to the stream
     * @return
     */
    inline std::ostream& operator<<(std::ostream& os, const XLCellIterator& it)
    {
        os << it.address();
        return os;
    }
}    // namespace OpenXLSX

// ===== Template specialization for std::distance.
namespace std    // NOLINT
{
    using OpenXLSX::XLCellIterator;
    template<>
    inline std::iterator_traits<XLCellIterator>::difference_type distance<XLCellIterator>(XLCellIterator first, XLCellIterator last)
    {
        return static_cast<std::iterator_traits<XLCellIterator>::difference_type>(first.distance(last));
    }
}    // namespace std

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#   pragma warning(pop)
#endif // _MSC_VER

#endif    // OPENXLSX_XLCELLITERATOR_HPP
