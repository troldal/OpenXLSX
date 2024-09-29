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

#ifndef OPENXLSX_XLCOLOR_HPP
#define OPENXLSX_XLCOLOR_HPP

#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wunknown-pragmas" // disable warning about below #pragma warning being unknown
#ifdef _MSC_VER                                    // additional condition because the previous line does not work on gcc 12.2
#   pragma warning(push)
#   pragma warning(disable : 4251)
#   pragma warning(disable : 4275)
#endif // _MSC_VER
#pragma GCC diagnostic pop

// ===== External Includes ===== //
#include <cstdint>    // Pull request #276
#include <string>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"

namespace OpenXLSX
{
    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLColor
    {
        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

        friend bool operator==(const XLColor& lhs, const XLColor& rhs);
        friend bool operator!=(const XLColor& lhs, const XLColor& rhs);

    public:
        /**
         * @brief
         */
        XLColor();

        /**
         * @brief
         * @param alpha
         * @param red
         * @param green
         * @param blue
         */
        XLColor(uint8_t alpha, uint8_t red, uint8_t green, uint8_t blue);

        /**
         * @brief
         * @param red
         * @param green
         * @param blue
         */
        XLColor(uint8_t red, uint8_t green, uint8_t blue);

        /**
         * @brief
         * @param hexCode
         */
        explicit XLColor(const std::string& hexCode);

        /**
         * @brief
         * @param other
         */
        XLColor(const XLColor& other);

        /**
         * @brief
         * @param other
         */
        XLColor(XLColor&& other) noexcept;

        /**
         * @brief
         */
        ~XLColor();

        /**
         * @brief
         * @param other
         * @return
         */
        XLColor& operator=(const XLColor& other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLColor& operator=(XLColor&& other) noexcept;

        /**
         * @brief
         * @param alpha
         * @param red
         * @param green
         * @param blue
         */
        void set(uint8_t alpha, uint8_t red, uint8_t green, uint8_t blue);

        /**
         * @brief
         * @param red
         * @param green
         * @param blue
         */
        void set(uint8_t red = 0, uint8_t green = 0, uint8_t blue = 0);

        /**
         * @brief
         * @param hexCode
         */
        void set(const std::string& hexCode);

        /**
         * @brief
         * @return
         */
        uint8_t alpha() const;

        /**
         * @brief
         * @return
         */
        uint8_t red() const;

        /**
         * @brief
         * @return
         */
        uint8_t green() const;

        /**
         * @brief
         * @return
         */
        uint8_t blue() const;

        /**
         * @brief
         * @return
         */
        std::string hex() const;

        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Variables
        //----------------------------------------------------------------------------------------------------------------------

    private:
        uint8_t m_alpha { 255 };

        uint8_t m_red { 0 };

        uint8_t m_green { 0 };

        uint8_t m_blue { 0 };
    };

}    // namespace OpenXLSX

namespace OpenXLSX
{
    /**
     * @brief
     * @param lhs
     * @param rhs
     * @return
     */
    inline bool operator==(const XLColor& lhs, const XLColor& rhs)
    {
        return lhs.alpha() == rhs.alpha() && lhs.red() == rhs.red() && lhs.green() == rhs.green() && lhs.blue() == rhs.blue();
    }

    /**
     * @brief
     * @param lhs
     * @param rhs
     * @return
     */
    inline bool operator!=(const XLColor& lhs, const XLColor& rhs) { return !(lhs == rhs); }

}    // namespace OpenXLSX

#pragma GCC diagnostic push
#pragma GCC diagnostic ignored "-Wunknown-pragmas" // disable warning about below #pragma warning being unknown
#ifdef _MSC_VER                                    // additional condition because the previous line does not work on gcc 12.2
#   pragma warning(pop)
#endif // _MSC_VER
#pragma GCC diagnostic pop

#endif    // OPENXLSX_XLCOLOR_HPP
