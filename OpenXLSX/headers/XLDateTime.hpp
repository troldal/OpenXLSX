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

#ifndef OPENXLSX_XLDATETIME_HPP
#define OPENXLSX_XLDATETIME_HPP

#ifdef _MSC_VER                                    // additional condition because the previous line does not work on gcc 12.2
#   pragma warning(push)
#   pragma warning(disable : 4251)
#   pragma warning(disable : 4275)
#endif // _MSC_VER

// ===== External Includes ===== //
#include <ctime>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLException.hpp"

// ========== CLASS AND ENUM TYPE DEFINITIONS ========== //
namespace OpenXLSX
{
    class OPENXLSX_EXPORT XLDateTime
    {
    public:
        /**
         * @brief Constructor.
         */
        XLDateTime();

        /**
         * @brief Constructor taking an Excel time point serial number as an argument.
         * @param serial Excel time point serial number.
         */
        explicit XLDateTime(double serial);

        /**
         * @brief Constructor taking a std::tm struct as an argument.
         * @param timepoint A std::tm struct.
         */
        explicit XLDateTime(const std::tm& timepoint);

        /**
         * @brief Constructor taking a unixtime format (seconds since 1/1/1970) as an argument.
         * @param unixtime A time_t number.
         */
        explicit XLDateTime(time_t unixtime);

        /**
         * @brief Copy constructor.
         * @param other Object to be copied.
         */
        XLDateTime(const XLDateTime& other);

        /**
         * @brief Move constructor.
         * @param other Object to be moved.
         */
        XLDateTime(XLDateTime&& other) noexcept;

        /**
         * @brief Destructor
         */
        ~XLDateTime();

        /**
         * @brief Copy assignment operator.
         * @param other Object to be copied.
         * @return Reference to the copied-to object.
         */
        XLDateTime& operator=(const XLDateTime& other);

        /**
         * @brief Move assignment operator.
         * @param other Object to be moved.
         * @return Reference to the moved-to object.
         */
        XLDateTime& operator=(XLDateTime&& other) noexcept;

        /**
         * @brief Assignment operator taking an Excel date/time serial number as an argument.
         * @param serial A floating point value with the serial number.
         * @return Reference to the copied-to object.
         */
        XLDateTime& operator=(double serial);

        /**
         * @brief Assignment operator taking a std::tm object as an argument.
         * @param timepoint std::tm object with the time point
         * @return Reference to the copied-to object.
         */
        XLDateTime& operator=(const std::tm& timepoint);

        /**
         * @brief Implicit conversion to Excel date/time serial number (any floating point type).
         * @tparam T Type to convert to (any floating point type).
         * @return Excel date/time serial number.
         */
        template<typename T,
                 typename = std::enable_if_t<std::is_floating_point_v<T>>>
        operator T() const    // NOLINT
        {
            return serial();
        }

        /**
         * @brief Implicit conversion to std::tm object.
         * @return std::tm object.
         */
        operator std::tm() const;    // NOLINT

        /**
         * @brief Get the date/time in the form of an Excel date/time serial number.
         * @return A double with the serial number.
         */
        double serial() const;

        /**
         * @brief Get the date/time in the form of a std::tm struct.
         * @return A std::tm struct with the time point.
         */
        std::tm tm() const;

    private:
        double m_serial { 1.0 }; /**<  */
    };
}    // namespace OpenXLSX

#ifdef _MSC_VER                                    // additional condition because the previous line does not work on gcc 12.2
#   pragma warning(pop)
#endif // _MSC_VER

#endif    // OPENXLSX_XLDATETIME_HPP
