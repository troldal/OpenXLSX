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

#ifndef OPENXLSX_XLWORKSHEET_HPP
#define OPENXLSX_XLWORKSHEET_HPP

#include "openxlsx_export.h"
#include "XLSheet.hpp"
#include "XLCellReference.hpp"
#include "XLCell.hpp"
#include "XLCellRange.hpp"
#include "XLRow.hpp"
#include "XLColumn.hpp"

namespace OpenXLSX
{
    namespace Impl
    {
        class XLWorksheet;
    } // namespace Impl

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLWorksheet : public XLSheet
    {
    public:

        /**
         * @brief
         * @param sheet
         */
        explicit XLWorksheet(Impl::XLSheet& sheet);

        /**
         * @brief
         * @param other
         */
        XLWorksheet(const XLWorksheet& other) = default;

        /**
         * @brief
         * @param other
         */
        XLWorksheet(XLWorksheet&& other) = default;

        /**
         * @brief
         */
        ~XLWorksheet() override = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLWorksheet& operator=(const XLWorksheet& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLWorksheet& operator=(XLWorksheet&& other) = default;

        /**
         * @brief Get a pointer to the XLCell object for the given cell reference.
         * @param ref An XLCellReference object with the address of the cell to get.
         * @return A reference to the requested XLCell object.
         */
        XLCell Cell(const XLCellReference& ref);

        /**
         * @brief Get a pointer to the XLCell object for the given cell reference.
         * @param ref An XLCellReference object with the address of the cell to get.
         * @return A const reference to the requested XLCell object.
         */
        const XLCell Cell(const XLCellReference& ref) const;

        /**
         * @brief Get the cell with the given address
         * @param address The address of the cell to get, e.g. 'A1'
         * @return A reference to the XLCell object at the given address
         */
        XLCell Cell(const std::string& address);

        /**
         * @brief Get the cell with the given address
         * @param address The address of the cell to get, e.g. 'A1'
         * @return A const reference to the XLCell object at the given address
         */
        const XLCell Cell(const std::string& address) const;

        /**
         * @brief Get the cell at the given coordinates.
         * @param rowNumber The row number (index base 1).
         * @param columnNumber The column number (index base 1).
         * @return A reference to the XLCell object at the given coordinates.
         */
        XLCell Cell(unsigned long rowNumber, unsigned int columnNumber);

        /**
         * @brief Get the cell at the given coordinates.
         * @param rowNumber The row number (index base 1).
         * @param columnNumber The column number (index base 1).
         * @return A const reference to the XLCell object at the given coordinates.
         */
        const XLCell Cell(unsigned long rowNumber, unsigned int columnNumber) const;

        /**
         * @brief Get a range for the area currently in use (i.e. from cell A1 to the last cell being in use).
         * @return An XLCellRange object with the entire range.
         */
        XLCellRange Range();

        /**
         * @brief Get a range for the area currently in use (i.e. from cell A1 to the last cell being in use).
         * @return A const XLCellRange object with the entire range.
         */
        const XLCellRange Range() const;

        /**
         * @brief Get a range with the given coordinates.
         * @param topLeft An XLCellReference object with the coordinates to the top left cell.
         * @param bottomRight An XLCellReference object with the coordinates to the bottom right cell.
         * @return An XLCellRange object with the requested range.
         */
        XLCellRange Range(const XLCellReference& topLeft, const XLCellReference& bottomRight);

        /**
         * @brief Get a range with the given coordinates.
         * @param topLeft An XLCellReference object with the coordinates to the top left cell.
         * @param bottomRight An XLCellReference object with the coordinates to the bottom right cell.
         * @return A const XLCellRange object with the requested range.
         */
        const XLCellRange Range(const XLCellReference& topLeft, const XLCellReference& bottomRight) const;

        /**
         * @brief Get the row with the given row number.
         * @param rowNumber The number of the row to retrieve.
         * @return A pointer to the XLRow object.
         */
        XLRow Row(unsigned long rowNumber);

        /**
         * @brief Get the row with the given row number.
         * @param rowNumber The number of the row to retrieve.
         * @return A const pointer to the XLRow object.
         */
        const XLRow Row(unsigned long rowNumber) const;

        /**
         * @brief Get the column with the given column number.
         * @param columnNumber The number of the column to retrieve.
         * @return A pointer to the XLColumn object.
         */
        XLColumn Column(unsigned int columnNumber);

        /**
         * @brief Get the column with the given column number.
         * @param columnNumber The number of the column to retrieve.
         * @return A const pointer to the XLColumn object.
         */
        const XLColumn Column(unsigned int columnNumber) const;

        /**
         * @brief Get an XLCellReference to the first (top left) cell in the worksheet.
         * @return An XLCellReference for the first cell.
         */
        XLCellReference FirstCell() const noexcept;

        /**
         * @brief Get an XLCellReference to the last (bottom right) cell in the worksheet.
         * @return An XLCellReference for the last cell.
         */
        XLCellReference LastCell() const noexcept;

        /**
         * @brief Get the number of columns in the worksheet.
         * @return The number of columns.
         */
        unsigned int ColumnCount() const noexcept;

        /**
         * @brief Get the number of rows in the worksheet.
         * @return The number of rows.
         */
        unsigned long RowCount() const noexcept;

        /**
         * @brief
         * @param fileName
         * @param decimal
         * @param delimiter
         */
        void Export(const std::string& fileName, char decimal = ',', char delimiter = ';');

        /**
         * @brief
         * @param fileName
         * @param delimiter
         */
        void Import(const std::string& fileName, const std::string& delimiter = ";");

    };
} // namespace OpenXLSX

#endif //OPENXLSX_XLWORKSHEET_HPP
