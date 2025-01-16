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

#ifndef OPENXLSX_XLTABLES_HPP
#define OPENXLSX_XLTABLES_HPP

// ===== External Includes ===== //
#include <cstdint>    // uint8_t, uint16_t, uint32_t
#include <ostream>    // std::basic_ostream
// #include <type_traits>
// #include <variant>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
// #include "XLDocument.hpp"
#include "XLException.hpp"
#include "XLXmlData.hpp"
#include "XLXmlFile.hpp"

namespace OpenXLSX
{
    /**
     * @brief The XLTables class is the base class for worksheet tables
     */
    class OPENXLSX_EXPORT XLTables : public XLXmlFile
    {
        friend class XLWorksheet;   // for access to XLXmlFile::getXmlPath
    public:
        /**
         * @brief Constructor
         */
        XLTables() : XLXmlFile(nullptr) {};

        /**
         * @brief The constructor.
         * @param xmlData the source XML of the table file
         */
        XLTables(XLXmlData* xmlData);

        /**
         * @brief The copy constructor.
         * @param other The object to be copied.
         * @note The default copy constructor is used, i.e. only shallow copying of pointer data members.
         */
        XLTables(const XLTables& other) = default;

        /**
         * @brief
         * @param other
         */
        XLTables(XLTables&& other) noexcept = default;

        /**
         * @brief The destructor
         * @note The default destructor is used, since cleanup of pointer data members is not required.
         */
        ~XLTables() = default;

        /**
         * @brief Assignment operator
         * @return A reference to the new object.
         * @note The default assignment operator is used, i.e. only shallow copying of pointer data members.
         */
        XLTables& operator=(const XLTables&) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLTables& operator=(XLTables&& other) noexcept = default;

        // /**
        //  * @brief getters
        //  */
        // std::string get(std::string cellRef) const;
        // 
        // /**
        //  * @brief setters
        //  */
        // bool set(std::string cellRef);

        /**
         * @brief Print the XML contents of this XLTables instance using the underlying XMLNode print function
         */
        void print(std::basic_ostream<char>& ostr) const;
    };
}    // namespace OpenXLSX

#endif    // OPENXLSX_XLTABLES_HPP
