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

#ifndef OPENXL_XLCELLGRIDITERATOR_H
#define OPENXL_XLCELLGRIDITERATOR_H

#include "XLCell.h"

namespace OpenXLSX::Impl
{

    class XLCellRange;

//======================================================================================================================
//========== XLCellIterator Class ======================================================================================
//======================================================================================================================

    /**
     * @brief Encapsulates an iterator for traversing all cells in a given range
     * @todo Unless this is used in range-based for loops, the IDE highlights errors (but compiles OK).
     */
    class XLCellIterator
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor. Takes a pointer to an XLCellRange object as an argument.
         * @param range The range to create an iterator for.
         * @param cell
         */
        explicit XLCellIterator(XLCellRange &range,
                                XLCell *cell);

        /**
         * @brief Method for incrementing the iterator.
         */
        const XLCellIterator &operator++();

        /**
         * @brief Method for checking equality.
         * @param other The XLCellGridIterator to compare with.
         * @return true if equal; otherwise false.
         */
        bool operator!=(const XLCellIterator &other) const;

        /**
         * @brief Method for dereferencing the underlying pointer object.
         * @return A reference to the value at the iterator.
         */
        XLCell &operator*() const;

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:

        XLCell *m_cell; /**< The cell at which the iterator is currently pointing. */
        XLCellRange *m_range; /**< The range the iterator is iterating over. */

        unsigned long m_currentRow; /**< The current row in the range. */
        unsigned int m_currentColumn; /**< The current column in the range. */

        unsigned long m_numRows; /**< The number of rows in the range. */
        unsigned int m_numColumns; /**< The number of columns in the range. */

    };


//======================================================================================================================
//========== XLCellIteratorConst Class =================================================================================
//======================================================================================================================

    class XLCellIteratorConst
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor. Takes a pointer to an XLCellRange object as an argument.
         * @param range The range to create an iterator for.
         * @param cell
         */
        explicit XLCellIteratorConst(const XLCellRange &range,
                                     const XLCell *cell);

        /**
         * @brief Method for incrementing the iterator.
         */
        const XLCellIteratorConst &operator++();

        /**
         * @brief Method for checking equality.
         * @param other The XLCellGridIterator to compare with.
         * @return true if equal; otherwise false.
         */
        bool operator!=(const XLCellIteratorConst &other) const;

        /**
         * @brief Method for dereferencing the underlying pointer object.
         * @return A reference to the value at the iterator.
         */
        const XLCell &operator*() const;

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:

        const XLCell *m_cell; /**< The cell at which the iterator is currently pointing. */
        const XLCellRange *m_range; /**< The range the iterator is iterating over. */

        unsigned long m_currentRow; /**< The current row in the range. */
        unsigned int m_currentColumn; /**< The current column in the range. */

        unsigned long m_numRows; /**< The number of rows in the range. */
        unsigned int m_numColumns; /**< The number of columns in the range. */

    };

}


#endif //OPENXL_XLCELLGRIDITERATOR_H
