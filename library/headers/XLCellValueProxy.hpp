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

#ifndef OPENXLSX_XLCELLVALUEPROXY_HPP
#define OPENXLSX_XLCELLVALUEPROXY_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <cstdint>
#include <type_traits>

// ===== OpenXLSX Includes ===== //
#include "XLCellValue.hpp"
#include "XLXmlParser.hpp"

// ========== CLASS AND ENUM TYPE DEFINITIONS ========== //
namespace OpenXLSX
{
    class XLCell;

    /**
     * @brief
     */
    class XLCellValueProxy
    {
        friend class XLCell;
        friend class XLCellValue;

    public:
        //---------- Public Member Functions ----------//

        /**
         * @brief
         */
        ~XLCellValueProxy();

        /**
         * @brief
         * @param other
         * @return
         */
        XLCellValueProxy& operator=(const XLCellValueProxy& other);

        /**
         * @brief
         * @tparam T
         * @param numberValue
         * @return
         */
        template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type* = nullptr>
        XLCellValueProxy& operator=(T numberValue);    // NOLINT

        /**
         * @brief
         * @tparam T
         * @param numberValue
         * @return
         */
        template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
        XLCellValueProxy& operator=(T numberValue);    // NOLINT

        /**
         * @brief
         * @tparam T
         * @param stringValue
         * @return
         */
        template<typename T,
                 typename std::enable_if<!std::is_same<T, XLCellValue>::value && !std::is_same<T, bool>::value &&
                                             std::is_constructible<T, const char*>::value,
                                         T>::type* = nullptr>
        XLCellValueProxy& operator=(T stringValue);    // NOLINT

        /**
         * @brief
         * @tparam T
         * @param value
         * @return
         */
        template<typename T, typename std::enable_if<std::is_same<T, XLCellValue>::value, T>::type* = nullptr>
        XLCellValueProxy& operator=(T value);    // NOLINT

        /**
         * @brief
         * @tparam T
         * @param numberValue
         */
        template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type* = nullptr>
        void set(T numberValue);

        /**
         * @brief
         * @tparam T
         * @param numberValue
         */
        template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
        void set(T numberValue);

        /**
         * @brief
         * @tparam T
         * @param stringValue
         */
        template<typename T,
                 typename std::enable_if<!std::is_same<T, XLCellValue>::value && !std::is_same<T, bool>::value &&
                                             std::is_constructible<T, const char*>::value,
                                         T>::type* = nullptr>
        void set(T stringValue);

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
        XLCellValueProxy& clear();

        /**
         * @brief
         * @return
         */
        XLCellValueProxy& setError();

        /**
         * @brief
         * @return
         */
        XLValueType type() const;

        /**
         * @brief
         * @return
         */
        std::string typeAsString() const;

        /**
         * @brief
         * @return
         */
        operator XLCellValue();    // NOLINT

    private:
        //---------- Private Member Functions ---------- //

        /**
         * @brief
         * @param cell
         * @param cellNode
         */
        XLCellValueProxy(XLCell* cell, XMLNode* cellNode);

        /**
         * @brief
         * @param other
         */
        XLCellValueProxy(const XLCellValueProxy& other);

        /**
         * @brief
         * @param other
         */
        XLCellValueProxy(XLCellValueProxy&& other) noexcept;

        /**
         * @brief
         * @param other
         * @return
         */
        XLCellValueProxy& operator=(XLCellValueProxy&& other) noexcept;

        /**
         * @brief
         * @param numberValue
         */
        void setInteger(int64_t numberValue);

        /**
         * @brief
         * @param numberValue
         */
        void setBoolean(bool numberValue);

        /**
         * @brief
         * @param numberValue
         */
        void setFloat(double numberValue);

        /**
         * @brief
         * @param stringValue
         */
        void setString(const char* stringValue);

        /**
         * @brief Get a reference to the XLCellValue object for the cell.
         * @return A reference to an XLCellValue object.
         */
        XLCellValue getValue() const;

        //---------- Private Member Variables ---------- //

        XLCell*  m_cell;     /**< */
        XMLNode* m_cellNode; /**< */
    };
}  // namespace OpenXLSX

// ========== TEMPLATE MEMBER IMPLEMENTATIONS ========== //
namespace OpenXLSX {
    /**
     * @details
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type*>
    XLCellValueProxy& XLCellValueProxy::operator=(T numberValue)
    {
        if constexpr (std::is_same<T, bool>::value) {    // if bool
            setBoolean(numberValue);
        }
        else {    // if not bool
            setInteger(numberValue);
        }

        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type*>
    XLCellValueProxy& XLCellValueProxy::operator=(T numberValue)
    {
        setFloat(numberValue);
        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    template<typename T,
             typename std::enable_if<!std::is_same<T, XLCellValue>::value && !std::is_same<T, bool>::value &&
                                         std::is_constructible<T, const char*>::value,
                                     T>::type*>
    XLCellValueProxy& XLCellValueProxy::operator=(T stringValue)
    {
        if constexpr (std::is_same<const char*, typename std::remove_reference<typename std::remove_cv<T>::type>::type>::value)
            setString(stringValue);
        else if constexpr (std::is_same<std::string_view, T>::value)
            setString(std::string(stringValue).c_str());
        else
            setString(stringValue.c_str());
        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_same<T, XLCellValue>::value, T>::type*>
    XLCellValueProxy& XLCellValueProxy::operator=(T value)
    {
        switch (value.type()) {
            case XLValueType::Boolean:
                setBoolean(value.template get<bool>());
                break;
            case XLValueType::Integer:
                setInteger(value.template get<int64_t>());
                break;
            case XLValueType::Float:
                setFloat(value.template get<double>());
                break;
            case XLValueType::String:
                setString(value.template get<std::string>().c_str());
                break;
            case XLValueType::Empty:
                break;
            default:
                break;
        }

        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type*>
    void XLCellValueProxy::set(T numberValue)
    {
        *this = numberValue;
    }

    /**
     * @details
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type*>
    void XLCellValueProxy::set(T numberValue)
    {
        *this = numberValue;
    }

    /**
     * @details
     * @pre
     * @post
     */
    template<typename T,
             typename std::enable_if<!std::is_same<T, XLCellValue>::value && !std::is_same<T, bool>::value &&
                                         std::is_constructible<T, const char*>::value,
                                     T>::type*>
    void XLCellValueProxy::set(T stringValue)
    {
        *this = stringValue;
    }

    /**
     * @brief
     * @tparam T
     * @return
     */
    template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type*>
    T XLCellValueProxy::get() const {
        return getValue().get<T>();
    }

    /**
     * @brief
     * @tparam T
     * @return
     */
    template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type*>
    T XLCellValueProxy::get() const {
        return getValue().get<T>();
    }

    /**
     * @brief
     * @tparam T
     * @return
     */
    template<typename T,
        typename std::enable_if<std::is_constructible<T, char*>::value && !std::is_same<T, bool>::value, char*>::type*>
    T XLCellValueProxy::get() const {
        return getValue().get<T>();
    }

}    // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_XLCELLVALUEPROXY_HPP
