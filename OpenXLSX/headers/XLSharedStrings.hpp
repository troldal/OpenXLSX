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

#ifndef OPENXLSX_XLSHAREDSTRINGS_HPP
#define OPENXLSX_XLSHAREDSTRINGS_HPP

#ifdef _MSC_VER
#   pragma warning(push)
#   pragma warning(disable : 4251)
#   pragma warning(disable : 4275)
#endif // _MSC_VER

#include <deque>
#include <limits>     // std::numeric_limits
#include <ostream>    // std::basic_ostream
#include <string>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"

namespace OpenXLSX
{
    constexpr size_t XLMaxSharedStrings = (std::numeric_limits< int32_t >::max)();    // pull request #261: wrapped max in parentheses to prevent expansion of windows.h "max" macro

    /**
     * @brief This class encapsulate the Excel concept of Shared Strings. In Excel, instead of havig individual strings
     * in each cell, cells have a reference to an entry in the SharedStrings register. This results in smalle file
     * sizes, as repeated strings are referenced easily.
     */
    class OPENXLSX_EXPORT XLSharedStrings : public XLXmlFile
    {
        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    public:
        /**
         * @brief
         */
        XLSharedStrings() = default;

        /**
         * @brief
         * @param xmlData
         * @param stringCache
         */
        explicit XLSharedStrings(XLXmlData* xmlData, std::deque<std::string>* stringCache);

        /**
         * @brief Destructor
         */
        ~XLSharedStrings();

        /**
         * @brief
         * @param other
         */
        XLSharedStrings(const XLSharedStrings& other) = default;

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
        XLSharedStrings& operator=(const XLSharedStrings& other) = default;

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
        int32_t getStringIndex(const std::string& str) const;

        /**
         * @brief
         * @param str
         * @return
         */
        bool stringExists(const std::string& str) const;

        /**
         * @brief
         * @param index
         * @return
         */
        const char* getString(int32_t index) const;

        /**
         * @brief Append a new string to the list of shared strings.
         * @param str The string to append.
         * @return An int32_t with the index of the appended string
         */
        int32_t appendString(const std::string& str);

        /**
         * @brief Clear the string at the given index.
         * @param index The index to clear.
         * @note There is no 'deleteString' member function, as deleting a shared string node will invalidate the
         * shared string indices for the cells in the spreadsheet. Instead use this member functions, which clears
         * the contents of the string, but keeps the XMLNode holding the string.
         */
        void clearString(int32_t index);

        // 2024-06-18 TBD if this is ever needed
        // /**
        //  * @brief check m_stringCache is initialized
        //  * @return true if m_stringCache != nullptr, false otherwise
        //  * @note 2024-05-28 added function to enable other classes to check m_stringCache status
        //  */
        // bool initialized() const { return m_stringCache != nullptr; }

        /**
         * @brief print the XML contents of the shared strings document using the underlying XMLNode print function
         */
        void print(std::basic_ostream<char>& ostr) const;

    private:
        std::deque<std::string>* m_stringCache {}; /** < Each string must have an unchanging memory address; hence the use of std::deque */
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER
#   pragma warning(pop)
#endif // _MSC_VER

#endif    // OPENXLSX_XLSHAREDSTRINGS_HPP
