//
// Created by KBA012 on 22/09/2016.
//

#ifndef OPENXL_XLCELLGRIDITERATOR_H
#define OPENXL_XLCELLGRIDITERATOR_H

#include "XLCell.h"

namespace RapidXLSX
{

    class XLCellRange;

//======================================================================================================================
//========== XLCellIterator Class ======================================================================================
//======================================================================================================================

    /**
     * @brief Encapsulates an iterator for traversing all cells in a given range, based on Boost.Iterator
     * @todo Unless this is used in range-based for loops, the IDE highlights errors (but compiles OK).
     * Consider to creater the iterator class independent of Boost.Iterator.
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
