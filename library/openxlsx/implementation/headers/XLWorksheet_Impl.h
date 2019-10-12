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

#include <vector>

#include "XLSheet_Impl.h"
#include "XLRow_Impl.h"
#include "XLColumn_Impl.h"
#include "XLCell_Impl.h"
#include "XLCellValue_Impl.h"
#include "XLCellReference_Impl.h"

namespace OpenXLSX::Impl
{
    class XLCellRange;

    //======================================================================================================================
    //========== XLColumnVector Alias ======================================================================================
    //======================================================================================================================

    /**
     * @brief A std::vector of std::unique_ptr's to XLColumn objects.
     */
    using XLColumns = std::map<unsigned int, XLColumn>;


    //======================================================================================================================
    //========== XLWorksheet Class =========================================================================================
    //======================================================================================================================

    /**
     * @brief A class encapsulating an Excel worksheet. Access to XLWorksheet objects should be via the workbook object.
     */
    class XLWorksheet : public XLSheet
    {

        friend class XLCell;

        friend class XLRow;

        friend class XLWorkbook;


        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor
         * @param parent A reference to the parent workbook.
         * @param name The name of the worksheet.
         * @param filePath The path to the worksheet .xml file.
         * @param xmlData
         */
        explicit XLWorksheet(XLWorkbook& parent, XMLAttribute name, const std::string& filePath,
                             const std::string& xmlData = "");

        /**
         * @brief Copy Constructor.
         * @note The copy constructor has been explicitly deleted.
         */
        XLWorksheet(const XLWorksheet& other) = delete;

        /**
         * @brief Move Constructor.
         * @note The move constructor has been explicitly deleted.
         */
        XLWorksheet(XLWorksheet&& other) = delete;

        /**
         * @brief Destructor.
         */
        ~XLWorksheet() override = default;

        /**
         * @brief Copy assignment operator.
         * @note The copy assignment operator has been explicitly deleted.
         */
        XLWorksheet& operator=(const XLWorksheet& other) = delete;

        /**
         * @brief Move assignment operator.
         * @note The move assignment operator has been explicitly deleted.
         */
        XLWorksheet& operator=(XLWorksheet&& other) = delete;

        /**
         * @brief Get a pointer to the XLCell object for the given cell reference.
         * @param ref An XLCellReference object with the address of the cell to get.
         * @return A reference to the requested XLCell object.
         */
        XLCell* Cell(const XLCellReference& ref);

        /**
         * @brief Get a pointer to the XLCell object for the given cell reference.
         * @param ref An XLCellReference object with the address of the cell to get.
         * @return A const reference to the requested XLCell object.
         */
        const XLCell* Cell(const XLCellReference& ref) const;

        /**
         * @brief Get the cell with the given address
         * @param address The address of the cell to get, e.g. 'A1'
         * @return A reference to the XLCell object at the given address
         */
        XLCell* Cell(const std::string& address);

        /**
         * @brief Get the cell with the given address
         * @param address The address of the cell to get, e.g. 'A1'
         * @return A const reference to the XLCell object at the given address
         */
        const XLCell* Cell(const std::string& address) const;

        /**
         * @brief Get the cell at the given coordinates.
         * @param rowNumber The row number (index base 1).
         * @param columnNumber The column number (index base 1).
         * @return A reference to the XLCell object at the given coordinates.
         */
        XLCell* Cell(unsigned long rowNumber, unsigned int columnNumber);

        /**
         * @brief Get the cell at the given coordinates.
         * @param rowNumber The row number (index base 1).
         * @param columnNumber The column number (index base 1).
         * @return A const reference to the XLCell object at the given coordinates.
         */
        const XLCell* Cell(unsigned long rowNumber, unsigned int columnNumber) const;

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
        XLRow* Row(unsigned long rowNumber);

        /**
         * @brief Get the row with the given row number.
         * @param rowNumber The number of the row to retrieve.
         * @return A const pointer to the XLRow object.
         */
        const XLRow* Row(unsigned long rowNumber) const;

        /**
         * @brief Get the column with the given column number.
         * @param columnNumber The number of the column to retrieve.
         * @return A pointer to the XLColumn object.
         */
        XLColumn* Column(unsigned int columnNumber);

        /**
         * @brief Get the column with the given column number.
         * @param columnNumber The number of the column to retrieve.
         * @return A const pointer to the XLColumn object.
         */
        const XLColumn* Column(unsigned int columnNumber) const;

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
         * @param oldName
         * @param newName
         */
        void UpdateSheetName(const std::string& oldName, const std::string& newName);

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

        /**
         * @brief
         * @return
         */
        const std::string& GetXmlData() const override;

        //----------------------------------------------------------------------------------------------------------------------
        //           Protected Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    protected:

        /**
         * @brief The overridden parseXMLData method is used to map or copy the XML data to the internal data structures.
         * @return true on success; otherwise false.
         */
        bool ParseXMLData() override;

        /**
         * @brief Creates a clone of the sheet and adds it to the parent XLWorkbook object.
         * @param newName The name of the cloned sheet.
         * @return A pointer to the newly created clone.
         * @todo Not yet implemented.
         */
        XLWorksheet* Clone(const std::string& newName) override;

        /**
         * @brief Get the data structure all columns in the worksheet.
         * @return A reference to the std::vector with the column data.
         */
        XLColumns* Columns();

        /**
         * @brief Get the data structure all columns in the worksheet.
         * @return A const reference to the std::vector with the column data.
         */
        const XLColumns* Columns() const;

        /**
         * @brief
         * @return
         */
        static std::string NewSheetXmlData();

        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    private:

        /**
         * @brief Get access to the dimension node in the underlying XML file.
         * @return A pointer to the XMLNode object.
         */
        XMLNode DimensionNode();

        /**
         * @brief Get read-only access to the dimension node in the underlying XML file.
         * @return A const pointer to the XMLNode object. Returned as pointer-to-const.
         */
        const XMLNode DimensionNode() const;

        /**
         * @brief Initialize the dimension node in the underlying XML file.
         */
        void InitDimensionNode();

        /**
         * @brief Get access to the sheet data node in the underlying XML file.
         * @return A pointer to the corresponding XMLNode object.
         */
        XMLNode SheetDataNode();

        /**
         * @brief Get access to the sheet data node in the underlying XML file.
         * @return A const pointer to the corresponding XMLNode object.
         */
        const XMLNode SheetDataNode() const;

        /**
         * @brief Initialize the sheet data node in the underlying XML file.
         */
        void InitSheetDataNode();

        /**
         * @brief Get a reference to the XMLNode object corresponding to the columns (cols) node in the XML file.
         * @return A XMLNode reference.
         */
        XMLNode ColumnsNode();

        /**
         * @brief Get a const reference to the XMLNode object corresponding to the columns (cols) node in the XML file.
         * @return A const XMLNode reference.
         */
        const XMLNode ColumnsNode() const;

        /**
         * @brief Initialize the columns (cols) node in the underlying XML file.
         */
        void InitColumnsNode();

        /**
         * @brief Get a reference to the XMLNode object corresponding to the sheet views node in the XML file.
         * @return An XMLNode reference.
         */
        XMLNode SheetViewsNode();

        /**
         * @brief Get a const reference to the XMLNode object corresponding to the sheet views node in the XML file.
         * @return A const XMLNode reference.
         */
        const XMLNode SheetViewsNode() const;

        /**
         * @brief Initialize the sheet views node in the underlying XML file.
         */
        void InitSheetViewsNode();

        /**
         * @brief Specify the first cell (upper left) of the worksheet.
         * @param cellRef An XLCellReference object with a reference to the first cell in the sheet.
         */
        void SetFirstCell(const XLCellReference& cellRef) noexcept;

        /**
         * @brief Specify the last cell (bottom right) of the spreadsheet.
         * @param cellRef An XLCellReference object with a reference to the last cell in the sheet.
         */
        void SetLastCell(const XLCellReference& cellRef) noexcept;



        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Variables
        //----------------------------------------------------------------------------------------------------------------------

    private:

        std::unique_ptr<XMLNode> m_dimensionNode; /**< The node specifying the dimensions of the sheet, e.g. "A1:AZ100"*/
        std::unique_ptr<XMLNode> m_sheetDataNode; /**< The node where the sheet data (i.e. rows and cells) begin. */
        std::unique_ptr<XMLNode> m_columnsNode; /**< The head node for sheet column data */
        std::unique_ptr<XMLNode> m_sheetViewsNode; /**< The head node for sheet views */

        /**< A pointer to the parent XLWorkbook object (const) */

        struct XLRowData
        {
            unsigned long rowIndex;
            std::unique_ptr<XLRow> rowItem = nullptr;
        };

        std::vector<XLRowData> m_rows; /**< A std::vector with pointers to all rows in the sheet. */
        XLColumns m_columns; /**< A std::vector with pointers to all columns in sheet. */

        XLCellReference m_firstCell; /**< The first cell in the sheet (i.e. the top left cell).*/
        mutable XLCellReference m_lastCell; /**<  The last cell in the sheet (i.e. the bottom right). */
        unsigned int m_maxColumn; /**< The last column with properties set */
    };
} // namespace OpenXLSX::Impl

#endif //OPENXLSX_IMPL_XLWORKSHEET_H
