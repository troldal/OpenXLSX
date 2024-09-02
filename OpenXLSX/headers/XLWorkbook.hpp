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

#ifndef OPENXLSX_XLWORKBOOK_HPP
#define OPENXLSX_XLWORKBOOK_HPP

#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wunknown-pragmas" // disable warning about below #pragma warning being unknown
#ifdef _MSC_VER                                    // additional condition because the previous line does not work on gcc 12.2
#   pragma warning(push)
#   pragma warning(disable : 4251)
#   pragma warning(disable : 4275)
#endif // _MSC_VER
#pragma GCC diagnostic pop

// ===== External Includes ===== //
#include <ostream>    // std::basic_ostream
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"

namespace OpenXLSX
{
    class XLSharedStrings;

    class XLSheet;

    class XLWorksheet;

    class XLChartsheet;

    /**
     * @brief The XLSheetType class is an enumeration of the available sheet types, e.g. Worksheet (ordinary
     * spreadsheets), and Chartsheet (sheets with only a chart).
     */
    enum class XLSheetType { Worksheet, Chartsheet, Dialogsheet, Macrosheet };

    /**
     * @brief This class encapsulates the concept of a Workbook. It provides access to the individual sheets
     * (worksheets or chartsheets), as well as functionality for adding, deleting, moving and renaming sheets.
     */
    class OPENXLSX_EXPORT XLWorkbook : public XLXmlFile
    {
        friend class XLSheet;
        friend class XLDocument;

    public:    // ---------- Public Member Functions ---------- //
        /**
         * @brief Default constructor. Creates an empty ('null') XLWorkbook object.
         */
        XLWorkbook() = default;

        /**
         * @brief Constructor. Takes a pointer to an XLXmlData object (stored in the parent XLDocument object).
         * @param xmlData A pointer to the underlying XLXmlData object, which holds the XML data.
         * @note Do not create an XLWorkbook object directly. Get access through the an XLDocument object.
         */
        explicit XLWorkbook(XLXmlData* xmlData);

        /**
         * @brief Copy Constructor.
         * @param other The XLWorkbook object to be copied.
         * @note The copy constructor has been explicitly defaulted.
         */
        XLWorkbook(const XLWorkbook& other) = default;

        /**
         * @brief Move constructor.
         * @param other The XLWorkbook to be moved.
         * @note The move constructor has been explicitly defaulted.
         */
        XLWorkbook(XLWorkbook&& other) = default;

        /**
         * @brief Destructor
         * @note Default destructor specified
         */
        ~XLWorkbook();

        /**
         * @brief Copy assignment operator.
         * @param other The XLWorkbook object to be assigned to the current.
         * @return A reference to *this
         * @note The copy assignment operator has been explicitly deleted, as XLWorkbook objects should not be copied.
         */
        XLWorkbook& operator=(const XLWorkbook& other) = default;

        /**
         * @brief Move assignment operator.
         * @param other The XLWorkbook to be move assigned.
         * @return A reference to *this
         * @note The move assignment operator has been explicitly deleted, as XLWorkbook objects should not be moved.
         */
        XLWorkbook& operator=(XLWorkbook&& other) = default;

        /**
         * @brief Get the sheet (worksheet or chartsheet) at the given index.
         * @param index The index at which the desired sheet is located.
         * @return A pointer to an XLAbstractSheet with the sheet at the index.
         * @note The index must be 1-based (rather than 0-based) as this is the default for Excel spreadsheets.
         */
        XLSheet sheet(uint16_t index);

        /**
         * @brief Get the sheet (worksheet or chartsheet) with the given name.
         * @param sheetName The name of the desired sheet.
         * @return A pointer to an XLAbstractSheet with the sheet at the index.
         */
        XLSheet sheet(const std::string& sheetName);

        /**
         * @brief Get the worksheet with the given name.
         * @param sheetName The name of the desired worksheet.
         * @return
         */
        XLWorksheet worksheet(const std::string& sheetName);

        /**
         * @brief Get the worksheet at the given index.
         * @param index The index at which the desired sheet is located.
         * @return
         */
        XLWorksheet worksheet(uint16_t index);

        /**
         * @brief Get the chartsheet with the given name.
         * @param sheetName The name of the desired chartsheet.
         * @return
         */
        XLChartsheet chartsheet(const std::string& sheetName);

        /**
         * @brief Get the chartsheet at the given index.
         * @param index The index at which the desired sheet is located.
         * @return
         */
        XLChartsheet chartsheet(uint16_t index);

        /**
         * @brief Delete sheet (worksheet or chartsheet) from the workbook.
         * @param sheetName Name of the sheet to delete.
         * @throws XLException An exception will be thrown if trying to delete the last worksheet in the workbook
         * @warning A workbook must contain at least one worksheet. Trying to delete the last worksheet from the
         * workbook will trow an exception.
         */
        void deleteSheet(const std::string& sheetName);

        /**
         * @brief
         * @param sheetName
         */
        void addWorksheet(const std::string& sheetName);

        /**
         * @brief
         * @param existingName
         * @param newName
         */
        void cloneSheet(const std::string& existingName, const std::string& newName);

        /**
         * @brief
         * @param sheetName
         * @param index
         */
        void setSheetIndex(const std::string& sheetName, unsigned int index);

        /**
         * @brief
         * @param sheetName
         * @return
         */
        unsigned int indexOfSheet(const std::string& sheetName) const;

        /**
         * @brief
         * @param sheetName
         * @return
         */
        XLSheetType typeOfSheet(const std::string& sheetName) const;

        /**
         * @brief
         * @param index
         * @return
         */
        XLSheetType typeOfSheet(unsigned int index) const;

        /**
         * @brief
         * @return
         */
        unsigned int sheetCount() const;

        /**
         * @brief
         * @return
         */
        unsigned int worksheetCount() const;

        /**
         * @brief
         * @return
         */
        unsigned int chartsheetCount() const;

        /**
         * @brief
         * @return
         */
        std::vector<std::string> sheetNames() const;

        /**
         * @brief
         * @return
         */
        std::vector<std::string> worksheetNames() const;

        /**
         * @brief
         * @return
         */
        std::vector<std::string> chartsheetNames() const;

        /**
         * @brief
         * @param sheetName
         * @return
         */
        bool sheetExists(const std::string& sheetName) const;

        /**
         * @brief
         * @param sheetName
         * @return
         */
        bool worksheetExists(const std::string& sheetName) const;

        /**
         * @brief
         * @param sheetName
         * @return
         */
        bool chartsheetExists(const std::string& sheetName) const;

        /**
         * @brief
         * @param oldName
         * @param newName
         */
        void updateSheetReferences(const std::string& oldName, const std::string& newName);

        /**
         * @brief
         * @return
         */
        XLSharedStrings sharedStrings();

        /**
         * @brief
         * @return
         */
        bool hasSharedStrings() const;

        /**
         * @brief
         */
        void deleteNamedRanges();

        /**
         * @brief set a flag to force full calculation upon loading the file in Excel
         */
        void setFullCalculationOnLoad();

        /**
         * @brief print the XML contents of the workbook.xml using the underlying XMLNode print function
         */
        void print(std::basic_ostream<char>& ostr) const;

    private:    // ---------- Private Member Functions ---------- //
        /**
         * @brief
         * @return
         */
        uint16_t createInternalSheetID();

        /**
         * @brief
         * @param sheetName
         * @return
         */
        std::string sheetID(const std::string& sheetName);

        /**
         * @brief
         * @param sheetID
         * @return
         */
        std::string sheetName(const std::string& sheetID) const;

        /**
         * @brief
         * @param sheetID
         * @return
         */
        std::string sheetVisibility(const std::string& sheetID) const;

        /**
         * @brief
         * @param sheetName
         * @param internalID
         */
        void prepareSheetMetadata(const std::string& sheetName, uint16_t internalID);

        /**
         * @brief
         * @param sheetRID
         * @param newName
         */
        void setSheetName(const std::string& sheetRID, const std::string& newName);

        /**
         * @brief
         * @param sheetRID
         * @param state
         */
        void setSheetVisibility(const std::string& sheetRID, const std::string& state);

        /**
         * @brief
         * @param sheetRID
         * @return
         */
        bool sheetIsActive(const std::string& sheetRID) const;

        /**
         * @brief
         * @param sheetRID
         * @return true if sheed with sheedRID could be set to active (or was already active), otherwise false
         */
        bool setSheetActive(const std::string& sheetRID);

        /**
         * @brief Check whether attribute string state matches a value that is considered not visible
         * @param state
         * @return true if state does not match a value that is considered not visible (hidden, veryHidden), otherwise false
         */
        bool isVisibleState(std::string const& state) const;

        /**
         * @brief Check whether sheetNode is not empty, and in case it has an attribute "state", that the state does not reflect hidden-ness
         * @param sheetNode
         * @return true if sheetNode can be considered visible (and could be activated)
         */
        bool isVisible(XMLNode const& sheetNode) const;
    };
}    // namespace OpenXLSX

#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wunknown-pragmas" // disable warning about below #pragma warning being unknown
#ifdef _MSC_VER                                    // additional condition because the previous line does not work on gcc 12.2
#   pragma warning(pop)
#endif // _MSC_VER
#pragma GCC diagnostic pop

#endif    // OPENXLSX_XLWORKBOOK_HPP
