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

// ===== External Includes ===== //
#include <cstdint>    // pull requests #216, #232
#include <ios>        // std::hex
#include <sstream>

// ===== OpenXLSX Includes ===== //
#include "XLColor.hpp"
#include "XLException.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
XLColor::XLColor() = default;

/**
 * @details
 */
XLColor::XLColor(uint8_t alpha, uint8_t red, uint8_t green, uint8_t blue) : m_alpha(alpha), m_red(red), m_green(green), m_blue(blue) {}

/**
 * @details
 */
XLColor::XLColor(uint8_t red, uint8_t green, uint8_t blue) : m_red(red), m_green(green), m_blue(blue) {}

/**
 * @details
 */
XLColor::XLColor(const std::string& hexCode) { set(hexCode); }

/**
 * @details
 */
XLColor::XLColor(const XLColor& other) = default;

/**
 * @details
 */
XLColor::XLColor(XLColor&& other) noexcept = default;

/**
 * @details
 */
XLColor::~XLColor() = default;

/**
 * @details
 */
XLColor& XLColor::operator=(const XLColor& other) = default;

/**
 * @details
 */
XLColor& XLColor::operator=(XLColor&& other) noexcept = default;

/**
 * @details
 */
void XLColor::set(uint8_t alpha, uint8_t red, uint8_t green, uint8_t blue)
{
    m_alpha = alpha;
    m_red   = red;
    m_green = green;
    m_blue  = blue;
}

/**
 * @details
 */
void XLColor::set(uint8_t red, uint8_t green, uint8_t blue)
{
    m_red   = red;
    m_green = green;
    m_blue  = blue;
}

/**
 * @details
 */
void XLColor::set(const std::string& hexCode)
{
    std::string alpha;
    std::string red;
    std::string green;
    std::string blue;

    auto temp = hex();

    constexpr int hexCodeSizeWithoutAlpha = 6;
    constexpr int hexCodeSizeWithAlpha    = 8;

    if (hexCode.size() == hexCodeSizeWithoutAlpha) {
        alpha = hex().substr(0, 2);
        red   = hexCode.substr(0, 2);
        green = hexCode.substr(2, 2);
        blue  = hexCode.substr(4, 2);
    }

    else if (hexCode.size() == hexCodeSizeWithAlpha) {
        alpha = hexCode.substr(0, 2);
        red   = hexCode.substr(2, 2);
        green = hexCode.substr(4, 2);
        blue  = hexCode.substr(6, 2);    // NOLINT
    }
    else
        throw XLInputError("Invalid color code");

    constexpr int hexBase = 16;
    m_alpha               = static_cast<uint8_t>(stoul(alpha, nullptr, hexBase));
    m_red                 = static_cast<uint8_t>(stoul(red, nullptr, hexBase));
    m_green               = static_cast<uint8_t>(stoul(green, nullptr, hexBase));
    m_blue                = static_cast<uint8_t>(stoul(blue, nullptr, hexBase));
}

/**
 * @details
 */
uint8_t XLColor::alpha() const { return m_alpha; }

/**
 * @details
 */
uint8_t XLColor::red() const { return m_red; }

/**
 * @details
 */
uint8_t XLColor::green() const { return m_green; }

/**
 * @details
 */
uint8_t XLColor::blue() const { return m_blue; }

/**
 * @details
 */
std::string XLColor::hex() const
{
    std::stringstream str;
    constexpr int     hexBase = 16;

    if (m_alpha < hexBase) str << "0";
    str << std::hex << static_cast<int>(m_alpha);

    if (m_red < hexBase) str << "0";
    str << std::hex << static_cast<int>(m_red);

    if (m_green < hexBase) str << "0";
    str << std::hex << static_cast<int>(m_green);

    if (m_blue < hexBase) str << "0";
    str << std::hex << static_cast<int>(m_blue);

    return (str.str());
}
