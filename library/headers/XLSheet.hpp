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

#ifndef OPENXLSX_IMPL_XLABSTRACTSHEET_H
#define OPENXLSX_IMPL_XLABSTRACTSHEET_H

#include "XLAbstractXMLFile.hpp"
#include "XLCell.hpp"
#include "XLCellReference.hpp"
#include "XLCellValue.hpp"
#include "XLChartsheet.hpp"
#include "XLColor.hpp"
#include "XLColumn.hpp"
#include "XLDefinitions.hpp"
#include "XLException.hpp"
#include "XLRow.hpp"
#include "XLWorksheet.hpp"

#include <type_traits>
#include <variant>
#include <vector>

namespace OpenXLSX
{
    class XLWorkbook;

    /**
     * @brief The XLAbstractSheet is a generalized sheet class, which functions as superclass for specialized classes,
     * such as XLWorksheet. It implements functionality common to all sheet types. This is a pure abstract class,
     * so it cannot be instantiated.
     */
    class XLSheet : public XLAbstractXMLFile
    {
        friend class XLWorkbook;

    public:
        /**
         * @brief The constructor. There are no default constructor, so all parameters must be provided for
         * constructing an XLAbstractSheet object. Since this is a pure abstract class, instantiation is only
         * possible via one of the derived classes.
         * @param parent A pointer to the parent XLDocument object.
         * @param name The name of the new sheet.
         * @param filepath A std::string with the relative path to the sheet file in the .xlsx package.
         * @param xmlData
         */
        explicit XLSheet(XLXmlData* xmlData);

        /**
         * @brief The copy constructor.
         * @param other The object to be copied.
         * @note The default copy constructor is used, i.e. only shallow copying of pointer data members.
         * @todo Can this method be deleted?
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
        ~XLSheet() override;

        /**
         * @brief Assignment operator
         * @return A reference to the new object.
         * @note The default assignment operator is used, i.e. only shallow copying of pointer data members.
         * @todo Can this method be deleted?
         */
        XLSheet& operator=(const XLSheet& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLSheet& operator=(XLSheet&& other) noexcept = default;

        /**
         * @brief Method to retrieve the name of the sheet.
         * @return A std::string with the sheet name.
         */
        std::string Name() const;

        /**
         * @brief Method for renaming the sheet.
         * @param name A std::string with the new name.
         */
        void SetName(const std::string& name);

        /**
         * @brief Method for getting the current visibility state of the sheet.
         * @return An XLSheetState enum object, with the current sheet state.
         */
        XLSheetState State() const;

        /**
         * @brief Method for setting the state of the sheet.
         * @param state An XLSheetState enum object with the new state.
         * @bug For some reason, this method doesn't work. The data is written correctly to the xml file, but the sheet
         * is not hidden when opening the file in Excel.
         */
        void SetState(XLSheetState state);

        /**
         * @brief
         * @return
         */
        XLColor Color();

        /**
         * @brief
         * @param color
         */
        void SetColor(const XLColor& color);

        /**
         * @brief
         * @param selected
         */
        void SetSelected(bool selected);

        /**
         * @brief Method to get the type of the sheet.
         * @return An XLSheetType enum object with the sheet type.
         */
        // XLSheetType Type() const = 0;

        /**
         * @brief Method for cloning the sheet.
         * @param newName A std::string with the name of the clone
         * @return A pointer to the cloned object.
         * @note This is a pure abstract method. I.e. it is implemented in subclasses.
         */
        // XLSheet* Clone(const std::string& newName) = 0;

        /**
         * @brief Method for getting the index of the sheet.
         * @return An int with the index of the sheet.
         */
        unsigned int Index() const;

        /**
         * @brief Method for setting the index of the sheet. This effectively moves the sheet to a different position.
         */
        void SetIndex();

        template<typename T>
        T Get()
        {
            if constexpr (std::is_same<T, XLWorksheet>::value)
                return std::get<XLWorksheet>(m_sheet);

            else if constexpr (std::is_same<T, XLChartsheet>::value)
                return std::get<XLChartsheet>(m_sheet);

            else
                throw XLException("Invalid sheet type.");
        }

        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Variables
        //----------------------------------------------------------------------------------------------------------------------

    private:
        std::variant<XLWorksheet, XLChartsheet> m_sheet; /**<  */
    };
}    // namespace OpenXLSX

#endif    // OPENXLSX_IMPL_XLABSTRACTSHEET_H
