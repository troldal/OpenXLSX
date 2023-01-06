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

    Written by Akira SHIMAHARA

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

#ifndef OPENXLSX_XLTABLEROWS_HPP
#define OPENXLSX_XLTABLEROWS_HPP

#pragma warning(push)
//#pragma warning(disable : 4251)
//#pragma warning(disable : 4275)

// ===== External Includes ===== //

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLCellRange.hpp"
#include "XLRow.hpp"

namespace OpenXLSX
{
    class XLTableRows;

    /**
     * @brief The XLTableRowIterator class is used to iterate througt XLTableRows
     */
    class OPENXLSX_EXPORT XLTableRowIterator
    {
    public:
        using iterator_category = std::forward_iterator_tag;
        using value_type        = XLCellRange;
        using pointer           = XLCellRange*;
        using reference         = XLCellRange&;

        /**
         * @brief
         * @param rowRange
         * @param loc
         */
        explicit XLTableRowIterator(const XLTableRows& tableRows, 
                                    XLIteratorLocation loc);

        /**
         * @brief
         */
        ~XLTableRowIterator();

        /**
         * @brief
         * @param other
         */
        XLTableRowIterator(const XLTableRowIterator& other);

        /**
         * @brief
         * @param other
         */
        XLTableRowIterator(XLTableRowIterator&& other) noexcept;

        /**
         * @brief
         * @param other
         * @return
         */
        XLTableRowIterator& operator=(const XLTableRowIterator& other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLTableRowIterator& operator=(XLTableRowIterator&& other) noexcept;

        /**
         * @brief
         * @return
         */
        XLTableRowIterator& operator++();

        /**
         * @brief
         * @return
         */
        XLTableRowIterator operator++(int);    // NOLINT

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
        bool operator==(const XLTableRowIterator& rhs) const;

        /**
         * @brief
         * @param rhs
         * @return
         */
        bool operator!=(const XLTableRowIterator& rhs) const;

        /**
         * @brief
         * @return
         */
        explicit operator bool() const;


    private:
        XLCellRange m_range;               /**< The cell range reference of the first cell in the range */
        uint32_t m_firstIterRow;
        uint32_t m_lastIterRow;
           /**< */
    };

    //----------------------------------------------------------------------------------------------------------------------
    //            Class XLTableRows
    //----------------------------------------------------------------------------------------------------------------------

    /**
     * @brief Class XLTableRows derived from XLRowRange. 
     * This an iterable object within the databodyrange
     */
    class OPENXLSX_EXPORT XLTableRows : public XLRowRange
    {
        friend class XLTableRowIterator;

    public:
        /**
         * @brief The constructor. 
         * @param XMLNode from the table file
         */
        XLTableRows(const XMLNode& dataNode, 
                    uint32_t firstRow, 
                    uint32_t lastRow,
                    uint16_t firstCol,
                    uint16_t lastCol,
                    const XLWorksheet* wks);
        /**
         * @brief Copy Constructor
         * @param other
         */
        XLTableRows(const XLTableRows& other);

        /**
         * @brief Move Constructor
         * @param other
         */
        XLTableRows(XLTableRows&& other) noexcept;


        //XLTableRows(const XLTableColumn&) = delete;
        //XLTableRows& operator=(const XLTableColumn&) = delete;
        ~XLTableRows();

        /**
         * @brief Copy assignment operator.
         * @note The copy assignment operator .
         */
        XLTableRows& operator=(const XLTableRows& other);

        /**
         * @brief Move assignment operator.
         * @note The move assignment operator has been explicitly deleted.
         */
        XLTableRows& operator=(XLTableRows&& other) noexcept;

        /**
         * @brief Overloading [] operator to access elements in array style
         * @return a XLCellRange corresponding to the index row
         * @note if the index is out of range, the entire table will be returned
         */
        XLCellRange operator[](uint32_t index) const;

        /**
         * @brief
         * @return
         */
        XLTableRowIterator begin();

        /**
         * @brief
         * @return
         */
        XLTableRowIterator end();

        /**
         * @brief
         * @return the number of rows of the table
         */
        uint32_t numRows() const;


    private:
            uint16_t   m_firstCol;                  /**< The cell reference of the first cell in the range */
            uint16_t   m_lastCol;         
    };
}    // namespace OpenXLSX

//#pragma warning(pop)
#endif    // OPENXLSX_XLTABLEROWS_HPP
