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

#ifndef OPENXLSX_IMPL_XLSHAREDSTRINGS_H
#define OPENXLSX_IMPL_XLSHAREDSTRINGS_H

// ===== Standard Library Includes ===== //
#include <vector>
#include <string_view>

// ===== OpenXLSX Includes ===== //
#include "XLAbstractXMLFile_Impl.hpp"
#include "XLXml_Impl.hpp"

namespace OpenXLSX::Impl
{

    class XLDocument;
    //======================================================================================================================
    //========== XLSharedStrings Class =====================================================================================
    //======================================================================================================================

    /**
     * @brief This class encapsulate the Excel concept of Shared Strings. In Excel, instead of havig individual strings
     * in each cell, cells have a reference to an entry in the SharedStrings register. This results in smalle file sizes,
     * as repeated strings are referenced easily.
     * @todo Consider defining a static method for creating a new shared strings object + XML file.
     */
    class XLSharedStrings : public XLAbstractXMLFile
    {

        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor
         * @param parent A pointer to the parent XLDocument
         * @param filePath The path to the sharedStrings.xml file
         */
        explicit XLSharedStrings(XLDocument& parent);

        /**
         * @brief Destructor
         */
        ~XLSharedStrings() override = default;

        /**
         * @brief
         * @param other
         */
        XLSharedStrings(const XLSharedStrings& other) = delete;

        /**
         * @brief
         * @param other
         */
        XLSharedStrings(XLSharedStrings&& other) noexcept = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLSharedStrings& operator=(const XLSharedStrings& other) = delete;

        /**
         * @brief
         * @param other
         * @return
         */
        XLSharedStrings& operator=(XLSharedStrings&& other) noexcept = default;

        /**
         * @brief
         * @param str
         * @return
         */
        int32_t GetStringIndex(const std::string& str) const;

        /**
         * @brief
         * @param str
         * @return
         */
        bool StringExists(const std::string& str) const;

        /**
         * @brief
         * @param index
         * @return
         */
        bool StringExists(uint32_t index) const;

        /**
         * @brief
         * @param index
         * @return
         */
        const char* GetString(uint32_t index) const;

        /**
         * @brief Append a new string to the list of shared strings.
         * @param str The string to append.
         * @return A long int with the index of the appended string
         */
        int32_t AppendString(const std::string& str);

        /**
         * @brief Clear the string at the given index.
         * @param index The index to clear.
         * @note There is no 'deleteString' member function, as deleting a shared string node will invalidate the
         * shared string indices for the cells in the spreadsheet. Instead use this member functions, which clears
         * the contents of the string, but keeps the XMLNode holding the string.
         */
        void ClearString(int index);

        //----------------------------------------------------------------------------------------------------------------------
        //           Protected Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    protected:

        /**
         * @brief Parse the contents of the underlying XML file and fills the datastructure with the data from the XML file.
         * @return true if successful; otherwise false.
         */
        bool ParseXMLData() override;
    };

} // namespace OpenXLSX::Impl

#endif //OPENXLSX_IMPL_XLSHAREDSTRINGS_H
