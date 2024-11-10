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

#ifndef OPENXLSX_XLSHEET_HPP
#define OPENXLSX_XLSHEET_HPP

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#   pragma warning(push)
#   pragma warning(disable : 4251)
#   pragma warning(disable : 4275)
#endif // _MSC_VER

// ===== External Includes ===== //
#include <cstdint>    // uint8_t, uint16_t, uint32_t
#include <ostream>    // std::basic_ostream
#include <type_traits>
#include <variant>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLCell.hpp"
#include "XLCellReference.hpp"
#include "XLColor.hpp"
#include "XLColumn.hpp"
#include "XLCommandQuery.hpp"
#include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLMergeCells.hpp"
#include "XLRow.hpp"
#include "XLXmlFile.hpp"

namespace OpenXLSX
{
    constexpr const bool XLEmptyHiddenCells = true; // readability constant for XLWorksheet::mergeCells parameter emptyHiddenCells

    /**
     * @brief The XLSheetState is an enumeration of the possible (visibility) states, e.g. Visible or Hidden.
     */
    enum class XLSheetState : uint8_t { Visible, Hidden, VeryHidden };

    /**
     * @brief The XLSheetBase class is the base class for the XLWorksheet and XLChartsheet classes. However,
     * it is not a base class in the traditional sense. Rather, it provides common functionality that is
     * inherited via the CRTP (Curiously Recurring Template Pattern) pattern.
     * @tparam T Type that will inherit functionality. Restricted to types XLWorksheet and XLChartsheet.
     */
    template<typename T, typename = std::enable_if_t<std::is_same_v<T, XLWorksheet> || std::is_same_v<T, XLChartsheet>>>
    class OPENXLSX_EXPORT XLSheetBase : public XLXmlFile
    {
    public:
        /**
         * @brief Constructor
         */
        XLSheetBase() : XLXmlFile(nullptr) {};

        /**
         * @brief The constructor. There are no default constructor, so all parameters must be provided for
         * constructing an XLAbstractSheet object. Since this is a pure abstract class, instantiation is only
         * possible via one of the derived classes.
         * @param xmlData
         */
        explicit XLSheetBase(XLXmlData* xmlData) : XLXmlFile(xmlData) {};

        /**
         * @brief The copy constructor.
         * @param other The object to be copied.
         * @note The default copy constructor is used, i.e. only shallow copying of pointer data members.
         */
        XLSheetBase(const XLSheetBase& other) = default;

        /**
         * @brief
         * @param other
         */
        XLSheetBase(XLSheetBase&& other) noexcept = default;

        /**
         * @brief The destructor
         * @note The default destructor is used, since cleanup of pointer data members is not required.
         */
        ~XLSheetBase() = default;

        /**
         * @brief Assignment operator
         * @return A reference to the new object.
         * @note The default assignment operator is used, i.e. only shallow copying of pointer data members.
         */
        XLSheetBase& operator=(const XLSheetBase&) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLSheetBase& operator=(XLSheetBase&& other) noexcept = default;

        /**
         * @brief
         * @return
         */
        XLSheetState visibility() const
        {
            XLQuery query(XLQueryType::QuerySheetVisibility);
            query.setParam("sheetID", relationshipID());
            const auto state  = parentDoc().execQuery(query).template result<std::string>();
            auto result = XLSheetState::Visible;

            if (state == "visible" || state.empty()) {
                result = XLSheetState::Visible;
            }
            else if (state == "hidden") {
                result = XLSheetState::Hidden;
            }
            else if (state == "veryHidden") {
                result = XLSheetState::VeryHidden;
            }

            return result;
        }

        /**
         * @brief
         * @param state
         */
        void setVisibility(XLSheetState state)
        {
            auto stateString = std::string();
            switch (state) {
                case XLSheetState::Visible:
                    stateString = "visible";
                    break;

                case XLSheetState::Hidden:
                    stateString = "hidden";
                    break;

                case XLSheetState::VeryHidden:
                    stateString = "veryHidden";
                    break;
            }

            parentDoc().execCommand(XLCommand(XLCommandType::SetSheetVisibility)
                                        .setParam("sheetID", relationshipID())
                                        .setParam("sheetVisibility", stateString));
        }

        /**
         * @brief
         * @return
         * @todo To be implemented.
         */
        XLColor color() const
        {
            return static_cast<const T&>(*this).getColor_impl();
        }

        /**
         * @brief
         * @param color
         */
        void setColor(const XLColor& color)
        {
            static_cast<T&>(*this).setColor_impl(color);
        }

        /**
         * @brief
         * @return
         */
        uint16_t index() const
        {
//            return uint16_t(std::stoi(parentDoc().execQuery(R"({ "query": "QuerySheetIndex", "sheetID": ")" + relationshipID() + "\"}")));

            XLQuery query(XLQueryType::QuerySheetIndex);
            query.setParam("sheetID", relationshipID());
            return static_cast<uint16_t>(std::stoi(parentDoc().execQuery(query).template result<std::string>()));
        }

        /**
         * @brief
         * @param index
         */
        void setIndex(uint16_t index)
        {
            parentDoc().execCommand(XLCommand(XLCommandType::SetSheetIndex)
                                        .setParam("sheetID", relationshipID())
                                        .setParam("sheetIndex", index));
        }

        /**
         * @brief Method to retrieve the name of the sheet.
         * @return A std::string with the sheet name.
         */
        std::string name() const
        {
            XLQuery query(XLQueryType::QuerySheetName);
            query.setParam("sheetID", relationshipID());
            return parentDoc().execQuery(query).template result<std::string>();
        }

        /**
         * @brief Method for renaming the sheet.
         * @param sheetName A std::string with the new name.
         */
        void setName(const std::string& sheetName)
        {
            parentDoc().execCommand(XLCommand(XLCommandType::SetSheetName)
                                        .setParam("sheetID", relationshipID())
                                        .setParam("sheetName", name())
                                        .setParam("newName", sheetName));
        }

        /**
         * @brief
         * @return
         */
        bool isSelected() const
        {
            return static_cast<const T&>(*this).isSelected_impl();
        }

        /**
         * @brief
         * @param selected
         */
        void setSelected(bool selected)
        {
            static_cast<T&>(*this).setSelected_impl(selected);
        }

        /**
         * @brief
         * @return
         */
        bool isActive() const
        {
            return static_cast<const T&>(*this).isActive_impl();
        }

        /**
         * @brief
         */
        bool setActive()
        {
            return static_cast<T&>(*this).setActive_impl();
        }

        /**
         * @brief Method for cloning the sheet.
         * @param newName A std::string with the name of the clone
         * @return A pointer to the cloned object.
         * @note This is a pure abstract method. I.e. it is implemented in subclasses.
         */
        void clone(const std::string& newName)
        {
            parentDoc().execCommand(XLCommand(XLCommandType::CloneSheet)
                                        .setParam("sheetID", relationshipID())
                                        .setParam("cloneName", newName));
        }
    };

    /**
     * @brief A class encapsulating an Excel worksheet. Access to XLWorksheet objects should be via the workbook object.
     */
    class OPENXLSX_EXPORT XLWorksheet final : public XLSheetBase<XLWorksheet>
    {
        friend class XLCell;
        friend class XLRow;
        friend class XLWorkbook;
        friend class XLSheetBase<XLWorksheet>;

        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    public:
        /**
         * @brief Default constructor
         */
        XLWorksheet() : XLSheetBase(nullptr) {};

        /**
         * @brief
         * @param xmlData
         */
        explicit XLWorksheet(XLXmlData* xmlData);

        /**
         * @brief Destructor.
         */
        ~XLWorksheet() = default;

        /**
         * @brief Copy Constructor.
         */
        XLWorksheet(const XLWorksheet& other);

        /**
         * @brief Move Constructor.
         */
        XLWorksheet(XLWorksheet&& other);

        /**
         * @brief Copy assignment operator.
         */
        XLWorksheet& operator=(const XLWorksheet& other);

        /**
         * @brief Move assignment operator.
         */
        XLWorksheet& operator=(XLWorksheet&& other);


        /**
         * @brief
         * @param ref
         * @return
         */
        XLCellAssignable cell(const std::string& ref) const;

        /**
         * @brief Get an XLCell object for the given cell reference.
         * @param ref An XLCellReference object with the address of the cell to get.
         * @return A reference to the requested XLCell object.
         */
        XLCellAssignable cell(const XLCellReference& ref) const;

        /**
         * @brief Get the cell at the given coordinates.
         * @param rowNumber The row number (index base 1).
         * @param columnNumber The column number (index base 1).
         * @return A reference to the XLCell object at the given coordinates.
         */
        XLCellAssignable cell(uint32_t rowNumber, uint16_t columnNumber) const;

        /**
         * @brief Get a range for the area currently in use (i.e. from cell A1 to the last cell being in use).
         * @return An XLCellRange object with the entire range.
         */
        XLCellRange range() const;

        /**
         * @brief Get a range with the given coordinates.
         * @param topLeft An XLCellReference object with the coordinates to the top left cell.
         * @param bottomRight An XLCellReference object with the coordinates to the bottom right cell.
         * @return An XLCellRange object with the requested range.
         */
        XLCellRange range(const XLCellReference& topLeft, const XLCellReference& bottomRight) const;

        /**
         * @brief Get a range with the given coordinates.
         * @param topLeft A std::string that is convertible to an XLCellReference to the top left cell
         * @param bottomRight A std::string that is convertible to an XLCellReference to the bottom right cell.
         * @return An XLCellRange object with the requested range.
         */
        XLCellRange range(std::string const& topLeft, std::string const& bottomRight) const;
   
        /**
         * @brief Get a range with the given coordinates.
         * @param rangeReference A std::string that contains two XLCellReferences separated by a delimiter ':'
         *                       Example "A2:B5"
         * @return An XLCellRange object with the requested range.
         */
        XLCellRange range(std::string const& rangeReference) const;

        /**
         * @brief
         * @return
         */
        XLRowRange rows() const;

        /**
         * @brief
         * @param rowCount
         * @return
         */
        XLRowRange rows(uint32_t rowCount) const;

        /**
         * @brief
         * @param firstRow
         * @param lastRow
         * @return
         */
        XLRowRange rows(uint32_t firstRow, uint32_t lastRow) const;

        /**
         * @brief Get the row with the given row number.
         * @param rowNumber The number of the row to retrieve.
         * @return A reference to the XLRow object.
         */
        XLRow row(uint32_t rowNumber) const;

        /**
         * @brief Get the column with the given column number.
         * @param columnNumber The number of the column to retrieve.
         * @return A reference to the XLColumn object.
         */
        XLColumn column(uint16_t columnNumber) const;

        /**
         * @brief Get the column with the given column reference
         * @param columnRef The letters referencing the given column
         * @return A reference to the XLColumn object.
         */
        XLColumn column(std::string const& columnRef) const;

        /**
         * @brief Get an XLCellReference to the last (bottom right) cell in the worksheet.
         * @return An XLCellReference for the last cell.
         */
        XLCellReference lastCell() const noexcept;

        /**
         * @brief Get the number of columns in the worksheet.
         * @return The number of columns.
         */
        uint16_t columnCount() const noexcept;

        /**
         * @brief Get the number of rows in the worksheet.
         * @return The number of rows.
         */
        uint32_t rowCount() const noexcept;

        /**
         * @brief
         * @param oldName
         * @param newName
         */
        void updateSheetName(const std::string& oldName, const std::string& newName);

        /**
         * @brief get an XLMergeCells object to directly access the member functions
         * @returns an XLMergeCells object for this worksheet
         */
        XLMergeCells & merges();

        /**
         * @brief merge the cells indicated by range
         * @param rangeToMerge the XLCellRange to merge, can be obtained from XLWorksheet::range functions
         * @param emptyHiddenCells if true (XLEmptyHiddenCells), the values of hidden cells will be deleted
         *                         (only from the cells, not from the shared strings table, if used)
         * @throws XLInputException if range comprises < 2 cells or any cell within rangeToMerge is already part of an existing range
         */
        void mergeCells(XLCellRange const& rangeToMerge, bool emptyHiddenCells = false);
        /**
         * @brief Convenience wrapper for mergeCells with a std::string range reference
         * @param rangeReference A range reference string for the cells to merge
         * @param emptyHiddenCells as above
         */
        void mergeCells(const std::string& rangeReference, bool emptyHiddenCells = false);

        /**
         * @brief remove the merge setting for the indicated range
         * @param rangeToUnmerge the XLCellRange to unmerge, can be obtained from XLWorksheet::range functions
         * @throws XLInputException if no merged range exists that matches rangeToMerge precisely
         */
        void unmergeCells(XLCellRange const& rangeToUnmerge);
        /**
         * @brief Convenience wrapper for unmergeCells with a std::string range reference
         * @param rangeReference A range reference string for the cells to unmerge
         */
        void unmergeCells(const std::string& rangeReference);

        /**
         * @brief Get the array index of xl/styles.xml:<styleSheet>:<cellXfs> for the style assigned to a column.
         *        This value is stored in the col attributes like so: style="2"
         * @param column The column letter(s) or index (1-based)
         * @returns The index of the applicable format style
         */
        XLStyleIndex getColumnFormat(uint16_t column) const;
        XLStyleIndex getColumnFormat(const std::string& column) const;

        /**
         * @brief set the column style as a reference to the array index of xl/styles.xml:<styleSheet>:<cellXfs>
         *        Subsequently, set the same style for all existing(!) cells in that column
         * @param column the column letter(s) or index (1-based)
         * @param cellFormatIndex the style to set, corresponding to the index of XLStyles::cellStyles()
         * @returns true on success, false on failure
         */
        bool setColumnFormat(uint16_t column, XLStyleIndex cellFormatIndex);
        bool setColumnFormat(const std::string& column, XLStyleIndex cellFormatIndex);

        /**
         * @brief get the array index of xl/styles.xml:<styleSheet>:<cellXfs> for the style assigned to a row
         *        this value is stored in the row attributes like so: s="2"
         * @param row the row index (1-based)
         * @returns the index of the applicable format style
         */
        XLStyleIndex getRowFormat(uint16_t row) const;

        /**
         * @brief set the row style as a reference to the array index of xl/styles.xml:<styleSheet>:<cellXfs>
         *        Subsequently, set the same style for all existing(!) cells in that row
         * @param row the row index (1-based)
         * @param cellFormatIndex the style to set, corresponding to the index of XLStyles::cellStyles()
         * @returns true on success, false on failure
         */
        bool setRowFormat(uint32_t row, XLStyleIndex cellFormatIndex);

    private:

        /**
         * @brief
         * @return
         */
        XLColor getColor_impl() const;

        /**
         * @brief
         * @param color
         */
        void setColor_impl(const XLColor& color);

        /**
         * @brief
         * @return
         */
        bool isSelected_impl() const;

        /**
         * @brief
         * @param selected
         */
        void setSelected_impl(bool selected);

        /**
         * @brief
         * @return
         */
        bool isActive_impl() const;

        /**
         * @brief
         */
        bool setActive_impl();

        XLMergeCells m_merges;    /**< class handling the <mergeCells> */
    };

    /**
     * @brief Class representing the an Excel chartsheet.
     * @todo This class is largely unimplemented and works just as a placeholder.
     */
    class OPENXLSX_EXPORT XLChartsheet final : public XLSheetBase<XLChartsheet>
    {
        friend class XLSheetBase<XLChartsheet>;

        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    public:
        /**
         * @brief Default constructor
         */
        XLChartsheet() : XLSheetBase(nullptr) {};

        /**
         * @brief
         * @param xmlData
         */
        explicit XLChartsheet(XLXmlData* xmlData);

        /**
         * @brief
         * @param other
         */
        XLChartsheet(const XLChartsheet& other) = default;

        /**
         * @brief
         * @param other
         */
        XLChartsheet(XLChartsheet&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLChartsheet();

        /**
         * @brief
         * @return
         */
        XLChartsheet& operator=(const XLChartsheet& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLChartsheet& operator=(XLChartsheet&& other) noexcept = default;

    private:

        /**
         * @brief
         * @return
         */
        XLColor getColor_impl() const;

        /**
         * @brief
         * @param color
         */
        void setColor_impl(const XLColor& color);

        /**
         * @brief
         * @return
         */
        bool isSelected_impl() const;

        /**
         * @brief
         * @param selected
         */
        void setSelected_impl(bool selected);
    };

    /**
     * @brief The XLAbstractSheet is a generalized sheet class, which functions as superclass for specialized classes,
     * such as XLWorksheet. It implements functionality common to all sheet types. This is a pure abstract class,
     * so it cannot be instantiated.
     */
    class OPENXLSX_EXPORT XLSheet final : public XLXmlFile
    {
    public:
        /**
         * @brief The constructor. There are no default constructor, so all parameters must be provided for
         * constructing an XLAbstractSheet object. Since this is a pure abstract class, instantiation is only
         * possible via one of the derived classes.
         * @param xmlData
         */
        explicit XLSheet(XLXmlData* xmlData);

        /**
         * @brief The copy constructor.
         * @param other The object to be copied.
         * @note The default copy constructor is used, i.e. only shallow copying of pointer data members.
         */
        XLSheet(const XLSheet& other) = default;

        /**
         * @brief
         * @param other
         */
        XLSheet(XLSheet&& other) noexcept = default;

        /**
         * @brief The destructor
         * @note The default destructor is used, since cleanup of pointer data members is not required.
         */
        ~XLSheet() = default;

        /**
         * @brief Assignment operator
         * @return A reference to the new object.
         * @note The default assignment operator is used, i.e. only shallow copying of pointer data members.
         */
        XLSheet& operator=(const XLSheet& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLSheet& operator=(XLSheet&& other) noexcept = default;

        /**
         * @brief Method for getting the current visibility state of the sheet.
         * @return An XLSheetState enum object, with the current sheet state.
         */
        XLSheetState visibility() const;

        /**
         * @brief Method for setting the state of the sheet.
         * @param state An XLSheetState enum object with the new state.
         * @bug For some reason, this method doesn't work. The data is written correctly to the xml file, but the sheet
         * is not hidden when opening the file in Excel.
         */
        void setVisibility(XLSheetState state);

        /**
         * @brief
         * @return
         */
        XLColor color() const;

        /**
         * @brief
         * @param color
         */
        void setColor(const XLColor& color);

        /**
         * @brief Method for getting the index of the sheet.
         * @return An int with the index of the sheet.
         */
        uint16_t index() const;

        /**
         * @brief Method for setting the index of the sheet. This effectively moves the sheet to a different position.
         */
        void setIndex(uint16_t index);

        /**
         * @brief Method to retrieve the name of the sheet.
         * @return A std::string with the sheet name.
         */
        std::string name() const;

        /**
         * @brief Method for renaming the sheet.
         * @param name A std::string with the new name.
         */
        void setName(const std::string& name);

        /**
         * @brief Determine whether the sheet is selected
         * @return
         */
        bool isSelected() const;

        /**
         * @brief
         * @param selected
         */
        void setSelected(bool selected);

        /**
         * @brief Method to get the type of the sheet.
         * @return An XLSheetType enum object with the sheet type.
         */
        template<
            typename SheetType,
                 typename = std::enable_if_t<std::is_same_v<SheetType, XLWorksheet> || std::is_same_v<SheetType, XLChartsheet>>>
        bool isType() const
        {
            return std::holds_alternative<SheetType>(m_sheet);
        }

        /**
         * @brief Method for cloning the sheet.
         * @param newName A std::string with the name of the clone
         * @return A pointer to the cloned object.
         * @note This is a pure abstract method. I.e. it is implemented in subclasses.
         */
        void clone(const std::string& newName);

        /**
         * @brief
         * @tparam T
         * @return
         */
        template<typename T, typename = std::enable_if_t<std::is_same_v<T, XLWorksheet> || std::is_same_v<T, XLChartsheet>>>
        T get() const
        {
            try {
                if constexpr (std::is_same<T, XLWorksheet>::value)
                    return std::get<XLWorksheet>(m_sheet);

                else if constexpr (std::is_same<T, XLChartsheet>::value)
                    return std::get<XLChartsheet>(m_sheet);
            }

            catch (const std::bad_variant_access&) {
                throw XLSheetError("XLSheet object does not contain the requested sheet type.");
            }
        }

        /**
         * @brief
         * @return
         */
        operator XLWorksheet() const;    // NOLINT

        /**
         * @brief
         * @return
         */
        operator XLChartsheet() const;    // NOLINT

        /**
         * @brief Print the XML contents of the XLSheet using the underlying XMLNode print function
         */
        void print(std::basic_ostream<char>& ostr) const;

        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Variables
        //----------------------------------------------------------------------------------------------------------------------

    private:
        std::variant<XLWorksheet, XLChartsheet> m_sheet; /**<  */
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER    // conditionally enable MSVC specific pragmas to avoid other compilers warning about unknown pragmas
#   pragma warning(pop)
#endif // _MSC_VER

#endif    // OPENXLSX_XLSHEET_HPP
