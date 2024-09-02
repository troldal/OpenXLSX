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

#ifndef OPENXLSX_XLPROPERTIES_HPP
#define OPENXLSX_XLPROPERTIES_HPP

#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wunknown-pragmas" // disable warning about below #pragma warning being unknown
#ifdef _MSC_VER                                    // additional condition because the previous line does not work on gcc 12.2
#   pragma warning(push)
#   pragma warning(disable : 4251)
#   pragma warning(disable : 4275)
#endif // _MSC_VER
#pragma GCC diagnostic pop

// ===== External Includes ===== //
#include <string>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"

namespace OpenXLSX
{
    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLProperties : public XLXmlFile
    {
    private:
        /**
         * @brief constructor helper function: create core.xml content from template
         * @param workbook
         */
        void createFromTemplate();

        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    public:
        /**
         * @brief
         */
        XLProperties() = default;

        /**
         * @brief
         * @param xmlData
         */
        explicit XLProperties(XLXmlData* xmlData);

        /**
         * @brief
         * @param other
         */
        XLProperties(const XLProperties& other) = default;

        /**
         * @brief
         * @param other
         */
        XLProperties(XLProperties&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLProperties();

        /**
         * @brief
         * @param other
         * @return
         */
        XLProperties& operator=(const XLProperties& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLProperties& operator=(XLProperties&& other) = default;

        /**
         * @brief
         * @param name
         * @param value
         * @return
         */
        void setProperty(const std::string& name, const std::string& value);

        /**
         * @brief
         * @param name
         * @param value
         * @return
         */
        void setProperty(const std::string& name, int value);

        /**
         * @brief
         * @param name
         * @param value
         * @return
         */
        void setProperty(const std::string& name, double value);

        /**
         * @brief
         * @param name
         * @return
         */
        std::string property(const std::string& name) const;

        /**
         * @brief
         * @param name
         */
        void deleteProperty(const std::string& name);

        //----------------------------------------------------------------------------------------------------------------------
        //           Protected Member Functions
        //----------------------------------------------------------------------------------------------------------------------
    };

    /**
     * @brief This class is a specialization of the XLAbstractXMLFile, with the purpose of the representing the
     * document app properties in the app.xml file (docProps folder) in the .xlsx package.
     */
    class OPENXLSX_EXPORT XLAppProperties : public XLXmlFile
    {
    private:
        /**
         * @brief constructor helper function: create app.xml content from template
         * @param workbook
         */
        void createFromTemplate(XMLDocument const & workbookXml);

        //--------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //--------------------------------------------------------------------------------------------------------------

    public:
        /**
         * @brief
         */
        XLAppProperties() = default;

        /**
         * @brief enable XLAppProperties to re-create a worksheet list in docProps/app.xml <TitlesOfParts> element from workbookXml
         * @param xmlData
         * @param workbook
         */
        explicit XLAppProperties(XLXmlData* xmlData, XMLDocument const & workbookXml);

        /**
         * @brief
         * @param xmlData
         */
        explicit XLAppProperties(XLXmlData* xmlData);

        /**
         * @brief
         * @param other
         */
        XLAppProperties(const XLAppProperties& other) = default;

        /**
         * @brief
         * @param other
         */
        XLAppProperties(XLAppProperties&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLAppProperties();

        /**
         * @brief
         * @param other
         * @return
         */
        XLAppProperties& operator=(const XLAppProperties& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLAppProperties& operator=(XLAppProperties&& other) noexcept = default;

        /**
         * @brief update the "HeadingPairs" entry for "Worksheets" *and* the "TitlesOfParts" vector size
         * @param increment change the sheet count by this (negative = decrement)
         * @throws XLInternalError when sheet count would become < 1
         */
        void incrementSheetCount(int16_t increment);

        /**
         * @brief initialize <TitlesOfParts> to contain all and only entries from workbookSheetNames & ensure HeadingPairs entry for Worksheets has the correct count
         * @param workbookSheetNames the vector of sheet names as returned by XLWorkbook::sheetNames()
         * @throws XLInternalError thrown by the underlying sheetNames call upon failure
         */
        void alignWorksheets(std::vector<std::string> const & workbookSheetNames);

        /**
         * @brief
         * @param title
         * @return
         */
        void addSheetName(const std::string& title);

        /**
         * @brief
         * @param title
         */
        void deleteSheetName(const std::string& title);

        /**
         * @brief
         * @param oldTitle
         * @param newTitle
         */
        void setSheetName(const std::string& oldTitle, const std::string& newTitle);

        /**
         * @brief
         * @param name
         * @param value
         */
        void addHeadingPair(const std::string& name, int value);

        /**
         * @brief
         * @param name
         */
        void deleteHeadingPair(const std::string& name);

        /**
         * @brief
         * @param name
         * @param newValue
         */
        void setHeadingPair(const std::string& name, int newValue);

        /**
         * @brief
         * @param name
         * @param value
         * @return
         */
        void setProperty(const std::string& name, const std::string& value);

        /**
         * @brief
         * @param name
         * @return
         */
        std::string property(const std::string& name) const;

        /**
         * @brief
         * @param name
         */
        void deleteProperty(const std::string& name);

        /**
         * @brief
         * @param sheetName
         * @return
         */
        void appendSheetName(const std::string& sheetName);

        /**
         * @brief
         * @param sheetName
         * @return
         */
        void prependSheetName(const std::string& sheetName);

        /**
         * @brief
         * @param sheetName
         * @param index
         * @return
         */
        void insertSheetName(const std::string& sheetName, unsigned int index);
    };

}    // namespace OpenXLSX

#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wunknown-pragmas" // disable warning about below #pragma warning being unknown
#ifdef _MSC_VER                                    // additional condition because the previous line does not work on gcc 12.2
#   pragma warning(pop)
#endif // _MSC_VER
#pragma GCC diagnostic pop

#endif    // OPENXLSX_XLPROPERTIES_HPP
