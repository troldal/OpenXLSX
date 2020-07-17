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

#ifndef OPENXLSX_IMPL_XLWORKSHEET_H
#define OPENXLSX_IMPL_XLWORKSHEET_H

#include "XLCell.hpp"
#include "XLCellReference.hpp"
#include "XLCellValue.hpp"
#include "XLColumn.hpp"
#include "XLRow.hpp"
#include "XLSheetBase.hpp"

#include <vector>

namespace OpenXLSX::Impl
{
    class XLCellRange;

    //======================================================================================================================
    //========== XLWorksheet Class
    //=========================================================================================
    //======================================================================================================================

    /**
     * @brief A class encapsulating an Excel worksheet. Access to XLWorksheet objects should be via the workbook object.
     */
    class XLWorksheet : public XLSheetBase<XLWorksheet>
    {
        friend class XLCell;
        friend class XLRow;
        friend class XLWorkbook;

        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    public:
        XLWorksheet() : XLSheetBase(nullptr) {};

        /**
         * @brief Constructor
         * @param parent A reference to the parent workbook.
         * @param name The name of the worksheet.
         * @param filePath The path to the worksheet .xml file.
         * @param xmlData
         */
        explicit XLWorksheet(XLXmlData* xmlData);

        /**
         * @brief Copy Constructor.
         * @note The copy constructor has been explicitly deleted.
         */
        XLWorksheet(const XLWorksheet& other) = default;

        /**
         * @brief Move Constructor.
         * @note The move constructor has been explicitly deleted.
         */
        XLWorksheet(XLWorksheet&& other) = default;

        /**
         * @brief Destructor.
         */
        ~XLWorksheet() override = default;

        /**
         * @brief Copy assignment operator.
         * @note The copy assignment operator has been explicitly deleted.
         */
        XLWorksheet& operator=(const XLWorksheet& other) = default;

        /**
         * @brief Move assignment operator.
         * @note The move assignment operator has been explicitly deleted.
         */
        XLWorksheet& operator=(XLWorksheet&& other) = default;

        Impl::XLCell Cell(const std::string& ref);

        Impl::XLCell Cell(const std::string& ref) const;

        /**
         * @brief Get a pointer to the XLCell object for the given cell reference.
         * @param ref An XLCellReference object with the address of the cell to get.
         * @return A const reference to the requested XLCell object.
         */
        Impl::XLCell Cell(const XLCellReference& ref);

        /**
         * @brief Get a pointer to the XLCell object for the given cell reference.
         * @param ref An XLCellReference object with the address of the cell to get.
         * @return A const reference to the requested XLCell object.
         */
        Impl::XLCell Cell(const XLCellReference& ref) const;

        /**
         * @brief Get the cell at the given coordinates.
         * @param rowNumber The row number (index base 1).
         * @param columnNumber The column number (index base 1).
         * @return A reference to the XLCell object at the given coordinates.
         */
        Impl::XLCell Cell(uint32_t rowNumber, uint16_t columnNumber);

        /**
         * @brief Get the cell at the given coordinates.
         * @param rowNumber The row number (index base 1).
         * @param columnNumber The column number (index base 1).
         * @return A const reference to the XLCell object at the given coordinates.
         */
        Impl::XLCell Cell(uint32_t rowNumber, uint16_t columnNumber) const;

        /**
         * @brief Get a range for the area currently in use (i.e. from cell A1 to the last cell being in use).
         * @return A const XLCellRange object with the entire range.
         */
        Impl::XLCellRange Range() const;

        /**
         * @brief Get a range with the given coordinates.
         * @param topLeft An XLCellReference object with the coordinates to the top left cell.
         * @param bottomRight An XLCellReference object with the coordinates to the bottom right cell.
         * @return A const XLCellRange object with the requested range.
         */
        Impl::XLCellRange Range(const XLCellReference& topLeft, const XLCellReference& bottomRight) const;

        /**
         * @brief Get the row with the given row number.
         * @param rowNumber The number of the row to retrieve.
         * @return A pointer to the XLRow object.
         */
        Impl::XLRow Row(uint32_t rowNumber);

        /**
         * @brief Get the row with the given row number.
         * @param rowNumber The number of the row to retrieve.
         * @return A const pointer to the XLRow object.
         */
        const XLRow* Row(uint32_t rowNumber) const;

        /**
         * @brief Get the column with the given column number.
         * @param columnNumber The number of the column to retrieve.
         * @return A pointer to the XLColumn object.
         */
        Impl::XLColumn Column(uint16_t columnNumber) const;

        /**
         * @brief Get an XLCellReference to the last (bottom right) cell in the worksheet.
         * @return An XLCellReference for the last cell.
         */
        XLCellReference LastCell() const noexcept;

        /**
         * @brief Get the number of columns in the worksheet.
         * @return The number of columns.
         */
        uint16_t ColumnCount() const noexcept;

        /**
         * @brief Get the number of rows in the worksheet.
         * @return The number of rows.
         */
        uint32_t RowCount() const noexcept;

        /**
         * @brief
         * @param oldName
         * @param newName
         */
        void UpdateSheetName(const std::string& oldName, const std::string& newName);

        /**
         * @brief
         * @return
         */
        std::string GetXmlData() const override;

        //----------------------------------------------------------------------------------------------------------------------
        //           Protected Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    protected:
        /**
         * @brief The overridden parseXMLData method is used to map or copy the XML data to the internal data
         * structures.
         * @return true on success; otherwise false.
         */
        bool ParseXMLData() override;

        /**
         * @brief Creates a clone of the sheet and adds it to the parent XLWorkbook object.
         * @param newName The name of the cloned sheet.
         * @return A pointer to the newly created clone.
         * @todo Not yet implemented.
         */
        XLWorksheet Clone(const std::string& newName);
    };
}    // namespace OpenXLSX::Impl

#endif    // OPENXLSX_IMPL_XLWORKSHEET_H
