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

#ifndef OPENXLSX_XLCELLVALUE_HPP
#define OPENXLSX_XLCELLVALUE_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <cstdint>
#include <string>
#include <variant>
#include <iostream>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
//#include "XLCellValueProxy.hpp"
#include "XLSharedStrings.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    class XLCellValueProxy;

    /**
     * @brief
     */
    enum class XLValueType { Empty, Boolean, Integer, Float, Error, String };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLCellValue
    {
        friend bool operator==(const XLCellValue& lhs, const XLCellValue& rhs);
        friend bool operator!=(const XLCellValue& lhs, const XLCellValue& rhs);
        friend bool operator<(const XLCellValue& lhs, const XLCellValue& rhs);
        friend bool operator>(const XLCellValue& lhs, const XLCellValue& rhs);
        friend bool operator<=(const XLCellValue& lhs, const XLCellValue& rhs);
        friend bool operator>=(const XLCellValue& lhs, const XLCellValue& rhs);
        friend std::ostream & operator<<(std::ostream &os, const XLCellValue& p);

    public:
        /**
         * @brief
         */
        XLCellValue();

        /**
         * @brief
         * @tparam T
         * @param value
         */
        template<typename T>
        XLCellValue(T value);

        XLCellValue(const XLCellValueProxy& proxy);

        /**
         * @brief
         * @param other
         */
        XLCellValue(const XLCellValue& other);

        /**
         * @brief
         * @param other
         */
        XLCellValue(XLCellValue&& other) noexcept;

        /**
         * @brief
         */
        ~XLCellValue();

        /**
         * @brief
         * @param other
         * @return
         */
        XLCellValue& operator=(const XLCellValue& other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLCellValue& operator=(XLCellValue&& other) noexcept;

        /**
         * @brief
         * @tparam T
         * @param value
         * @return
         */
        template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type* = nullptr>
        XLCellValue& operator=(T value);

        template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
        XLCellValue& operator=(T value);

        template<typename T,
            typename std::enable_if<std::is_constructible<T, char*>::value && !std::is_same<T, bool>::value, char*>::type* = nullptr>
        XLCellValue& operator=(T value);

        /**
         * @brief
         * @tparam T
         * @param value
         */
        template<typename T>
        void set(T value);

        /**
         * @brief
         * @tparam T
         * @return
         */
        template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type* = nullptr>
        T get() const;

        /**
         * @brief
         * @tparam T
         * @return
         */
        template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
        T get() const;

        /**
         * @brief
         */
        XLCellValue& clear();

        /**
         * @brief
         * @tparam T
         * @return
         */
        template<typename T,
            typename std::enable_if<std::is_constructible<T, char*>::value && !std::is_same<T, bool>::value, char*>::type* = nullptr>
        T get() const;

        /**
         * @brief
         * @return
         */
        XLCellValue& setError();

        /**
         * @brief
         * @return
         */
        XLValueType type() const;

    private:
        XLValueType                                      m_type { XLValueType::Empty }; /**< */
        std::variant<std::string, int64_t, double, bool> m_value { "" };                /**< */
    };

#include "impl/XLCellValue.inl"

    /**
     * @brief
     * @param lhs
     * @param rhs
     * @return
     */
    inline bool operator==(const XLCellValue& lhs, const XLCellValue& rhs)
    {
        return lhs.m_value == rhs.m_value;
    }

    /**
     * @brief
     * @param lhs
     * @param rhs
     * @return
     */
    inline bool operator!=(const XLCellValue& lhs, const XLCellValue& rhs)
    {
        return lhs.m_value != rhs.m_value;
    }

    /**
     * @brief
     * @param lhs
     * @param rhs
     * @return
     */
    inline bool operator<(const XLCellValue& lhs, const XLCellValue& rhs)
    {
        return lhs.m_value < rhs.m_value;
    }

    /**
     * @brief
     * @param lhs
     * @param rhs
     * @return
     */
    inline bool operator>(const XLCellValue& lhs, const XLCellValue& rhs)
    {
        return lhs.m_value > rhs.m_value;
    }

    /**
     * @brief
     * @param lhs
     * @param rhs
     * @return
     */
    inline bool operator<=(const XLCellValue& lhs, const XLCellValue& rhs)
    {
        return lhs.m_value <= rhs.m_value;
    }

    /**
     * @brief
     * @param lhs
     * @param rhs
     * @return
     */
    inline bool operator>=(const XLCellValue& lhs, const XLCellValue& rhs)
    {
        return lhs.m_value >= rhs.m_value;
    }

    inline std::ostream & operator<<(std::ostream &os, const XLCellValue& value)
    {
        switch (value.type()) {
            case XLValueType::Empty:
                return os << "";
            case XLValueType::Boolean:
                return os << value.get<bool>();
            case XLValueType::Integer:
                return os << value.get<int64_t>();
            case XLValueType::Float:
                return os << value.get<double>();
            case XLValueType::String:
                return os << value.get<std::string_view>();
            default:
                return os << "";
        }
    }
}    // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_XLCELLVALUE_HPP
