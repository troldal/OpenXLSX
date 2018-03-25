//
// Created by KBA012 on 15/09/2016.
//

#ifndef OPENXL_XLCELLRANGE_H
#define OPENXL_XLCELLRANGE_H

#include "XLCellReference.h"
#include "XLCellIterator.h"

#include <string>

namespace RapidXLSX
{
    class XLWorksheet;

//======================================================================================================================
//========== XLWorksheet Class =========================================================================================
//======================================================================================================================

    /**
     * @brief This class encapsulates the concept of a cell range, i.e. a square area
     * (or subset) of cells in a spreadsheet.
     * @todo Consider specifying starting cell and direction of iterator.
     */
    class XLCellRange
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor
         * @param sheet A pointer to the parent spreadsheet, i.e. the sheet from which the range refers. Must not be nullptr.
         * @param topLeft The first (top left) cell in the range.
         * @param bottomRight The last (bottom right) cell in the range.
         */
        explicit XLCellRange(XLWorksheet &sheet,
                             const XLCellReference &topLeft,
                             const XLCellReference &bottomRight);

        /**
         * @brief
         * @param sheet
         * @param topLeft
         * @param bottomRight
         */
        explicit XLCellRange(const XLWorksheet &sheet,
                             const XLCellReference &topLeft,
                             const XLCellReference &bottomRight);

        /**
         * @brief Copy constructor [default].
         * @param other The range object to be copied.
         * @note This implements the default copy constructor, i.e. memberwise copying.
         */
        XLCellRange(const XLCellRange &other) = default;

        /**
         * @brief Move constructor [default].
         * @param other The range object to be moved.
         * @note This implements the default move constructor, i.e. memberwise move.
         */
        XLCellRange(XLCellRange &&other) = default;

        /**
         * @brief Destructor [default]
         * @note This implements the default destructor.
         */
        ~XLCellRange() = default;

        /**
         * @brief The copy assignment operator [default]
         * @param other The range object to be copied and assigned.
         * @return A reference to the new object.
         * @throws A std::range_error if the source range and destination range are of different size and shape.
         * @note This implements the default copy assignment operator.
         */
        XLCellRange &operator=(const XLCellRange &other);

        /**
         * @brief The move assignment operator [default].
         * @param other The range object to be moved and assigned.
         * @return A reference to the new object.
         * @note This implements the default move assignment operator.
         */
        XLCellRange &operator=(XLCellRange &&other) = default;

        /**
         * @brief Get a pointer to the cell at the given coordinates.
         * @param row The row number, relative to the first row of the range (index base 1).
         * @param column The column number, relative to the first column of the range (index base 1).
         * @return A pointer to the cell at the given range coordinates.
         */
        XLCell *Cell(unsigned long row,
                     unsigned int column);

        /**
         * @brief Get a const pointer to the cell at the given coordinates.
         * @param row The row number, relative to the first row of the range (index base 1).
         * @param column The column number, relative to the first column of the range (index base 1).
         * @return A const pointer to the cell at the given range coordinates.
         */
        const XLCell *Cell(unsigned long row,
                           unsigned int column) const;

        /**
         * @brief Get the number of rows in the range.
         * @return The number of rows.
         */
        unsigned long NumRows() const;

        /**
         * @brief Get the number of columns in the range.
         * @return The number of columns.
         */
        unsigned int NumColumns() const;

        /**
         * @brief Transpose the range.
         */
        void Transpose(bool state) const;

        /**
         * @brief Get an iterator pointing to the first cell in the range.
         * @return The iterator pointing to the first cell in the range.
         */
        XLCellIterator begin();

        /**
         * @brief
         * @return
         */
        XLCellIteratorConst begin() const;

        /**
         * @brief Get an iterator pointing to the end of the grid (nullptr)
         * @return The iterator pointing to the end of the range.
         */
        XLCellIterator end();

        /**
         * @brief
         * @return
         */
        XLCellIteratorConst end() const;

        /**
         * @brief
         */
        void Clear();

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:

        XLWorksheet *m_parentWorksheet; /**< A pointer to the parent spreadsheet */

        XLCellReference m_topLeft; /**< The cell reference of the first cell in the range */
        XLCellReference m_bottomRight; /**< The cell reference of the last cell in the range */
        unsigned long m_rowOffset; /**< The row offset, relative to the parent spreadsheet */
        unsigned int m_columnOffset; /**< The column offset, relative to the parent spreadsheet */
        unsigned long m_rows; /**< The number of rows in the range */
        unsigned int m_columns; /**< The number of columns in the range */

        mutable bool m_transpose; /**< */

    };

}

#endif //OPENXL_XLCELLRANGE_H