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

#ifndef OPENXLSX_XLCELLVALUE_H
#define OPENXLSX_XLCELLVALUE_H

#include "XLProperty.h"
#include <type_traits>
#include <string>

namespace OpenXLSX {
    namespace Impl {
        class XLCellValue;
    }

    /**
     * @brief
     */
    class XLCellValue
    {
    public:

        /**
         * @brief
         * @param value
         */
        explicit XLCellValue(Impl::XLCellValue& value);

        /**
         * @brief
         * @param other
         */
        XLCellValue(const XLCellValue& other) = default;

        /**
         * @brief
         * @param other
         */
        XLCellValue(XLCellValue&& other) = default;

        /**
         * @brief
         */
        virtual ~XLCellValue() = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLCellValue& operator=(const XLCellValue& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLCellValue& operator=(XLCellValue&& other) = default;

        /**
         * @brief
         * @tparam T
         * @param numberValue
         * @return
         */
        template <typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type* = nullptr>
        XLCellValue& operator=(T numberValue);

        /**
         * @brief
         * @tparam T
         * @param numberValue
         * @return
         */
        template <typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
        XLCellValue& operator=(T numberValue);

        /**
         * @brief
         * @param stringValue
         * @return
         */
        XLCellValue& operator=(const char* stringValue);

        /**
         * @brief
         * @param stringValue
         * @return
         */
        XLCellValue& operator=(const std::string& stringValue);

        /**
         * @brief
         * @return
         */
        std::string AsString() const;

        /**
         * @brief
         * @tparam T
         * @param numberValue
         */
        template <typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type* = nullptr>
        void Set(T numberValue);

        /**
         * @brief
         * @tparam T
         * @param numberValue
         */
        template <typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
        void Set(T numberValue);

        /**
         * @brief
         * @param stringValue
         */
        void Set(const char* stringValue);

        /**
         * @brief
         * @param stringValue
         */
        void Set(const std::string& stringValue);

        /**
         * @brief
         * @tparam T
         * @return
         */
        template <typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type* = nullptr>
        T Get() const;

        /**
         * @brief
         * @tparam T
         * @return
         */
        template <typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
        T Get() const;

        /**
         * @brief
         * @tparam T
         * @return
         */
        template <typename T, typename std::enable_if<std::is_constructible<T, char*>::value && !std::is_same<T,bool>::value,char*>::type* = nullptr>
        T Get() const;

        /**
         * @brief
         */
        void Clear();

        /**
         * @brief
         * @return
         */
        XLValueType ValueType() const;


    private:

        /**
         * @brief
         * @param numberValue
         */
        void SetInteger(long long int numberValue);

        /**
         * @brief
         * @param numberValue
         */
        void SetBoolean(bool numberValue);

        /**
         * @brief
         * @param numberValue
         */
        void SetFloat(long double numberValue);

        /**
         * @brief
         * @return
         */
        long long int GetInteger() const;

        /**
         * @brief
         * @return
         */
        bool GetBoolean() const;

        /**
         * @brief
         * @return
         */
        long double GetFloat() const;

        /**
         * @brief
         * @return
         */
        const char* GetString() const;

    private:
        Impl::XLCellValue* m_value; /**< */
    };

    /**
     * @details
     */
    template <typename T, typename std::enable_if<std::is_integral<T>::value,
                                                  long long int>::type*>
    XLCellValue& XLCellValue::operator=(T numberValue) {

        Set(numberValue);
        return *this;
    }

    /**
     * @details
     */
    template <typename T, typename std::enable_if<std::is_floating_point<T>::value,
                                                  long double>::type*>
    XLCellValue& XLCellValue::operator=(T numberValue) {

        Set(numberValue);
        return *this;
    }

    /**
     * @details
     */
    template <typename T, typename std::enable_if<std::is_integral<T>::value,
                                                  long long int>::type*>
    void XLCellValue::Set(T numberValue) {

        if constexpr (std::is_same<T, bool>::value) { // if bool
            SetBoolean(numberValue);
        }
        else { // if not bool
            SetInteger(numberValue);
        }
    }

    /**
     * @details
     */
    template <typename T, typename std::enable_if<std::is_floating_point<T>::value,
                                                  long double>::type*>
    void XLCellValue::Set(T numberValue) {

        SetFloat(numberValue);
    }

    /**
     * @details
     */
    template <typename T, typename std::enable_if<std::is_integral<T>::value,
                                                  long long int>::type*>
    T XLCellValue::Get() const {

        if constexpr (std::is_same<T, bool>::value) {
            return GetBoolean();
        }
        else {
            return GetInteger();
        }
    }

    /**
     * @details
     */
    template <typename T, typename std::enable_if<std::is_floating_point<T>::value,
                                                  long double>::type*>
    T XLCellValue::Get() const {

        return GetFloat();
    }

    /**
     * @details
     */
    template <typename T, typename std::enable_if<std::is_constructible<T,
                                                                        char*>::value && !std::is_same<T,
                                                                                                       bool>::value,
                                                  char*>::type*>
    T XLCellValue::Get() const {

        return T(GetString());
    }

}



#endif //OPENXLSX_XLCELLVALUE_H
