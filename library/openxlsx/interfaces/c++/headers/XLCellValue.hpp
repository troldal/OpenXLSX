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

#include "openxlsx_export.h"
#include "XLDefinitions.hpp"
#include <type_traits>
#include <string>

namespace OpenXLSX
{
    namespace Impl
    {
        class XLCellValue;
    } // namespace Impl

    /**
     * @brief The XLCellValue class represents the concept of a cell value. This can be in the form of a number
     * (an integer or a float), a string, a boolean or no value (empty).
     *
     * @details The XLCellValue class encapsulates the concept of a cell value, i.e. the string, number, boolean or
     * empty value that a cell can contain. Other cell contents, such as formatting or formulas, are handled by separate
     * classes.
     *
     * ### Usage ###
     * An XLCellValue object is not created directly by the user, but rather accessed and/or modified through the XLCell
     * interface (the Value() member function).
     * The value can be set either through the Set() member function or the operator=(). Both have been overloaded to take
     * any integer or floating point value, bool value or string value (either std::string or char* string). In order to
     * set the cell value to empty, use the Clear() member function.
     *
     * It is also possible to assign one XLCellValue to another, using the overloaded copy assignment operator (operator=()).
     * Note that this behaviour is different from most of the classes in the OpenXLSX interface, where each object behaves
     * like an opaque pointer and the copy assignment operator essentially produces a new pointer to the same implementation
     * object. For XLCellValue objects, however, operator=() will copy the value of the source object to the destination object.
     * The copy/move **constructors**, however, has default behaviour; the new object will point to the same implementation
     * object as the source. Note also that a move assignment operator has not been implemented, as the value **must**
     * always be copied when assigning to an existing object.
     *
     * To get the current value of an XLCellValue object, use the Get<>() member function. The Get<>() member function is a
     * template function, taking the value return type as template argument. It has been overloaded to take any integer
     * or floating point type, bool type or string type that can be constructed from a char* string (e.g. std::string or
     * std::string_view). It is also possible to get the value as a std::string, regardless of the value type, using the
     * AsString() member function.
     *
     * ### Example ###
     * Here follows an example that sets the value and the prints it to std::cout:
     * ```cpp
     * // Code to create document/workbool/worksheet...
     *
     * wks->Cell("A1")->Value() = 3.14159;
     * wks->Cell("B1")->Value() = 42;
     * wks->Cell("C1")->Value() = "Hello OpenXLSX!";
     * wks->Cell("D1")->Value() = true;

     * auto A1 = wks->Cell("A1")->Value().Get<double>();
     * auto B1 = wks->Cell("B1")->Value().Get<int>();
     * auto C1 = wks->Cell("C1")->Value().Get<std::string>();
     * auto D1 = wks->Cell("D1")->Value().Get<bool>();

     * cout << "Cell A1: " << A1 << endl;
     * cout << "Cell B1: " << B1 << endl;
     * cout << "Cell C1: " << C1 << endl;
     * cout << "Cell D1: " << D1 << endl;
     * ```
     */
    class OPENXLSX_EXPORT XLCellValue
    {
        friend class XLCell;

    private:
        /**
         * @brief
         * @param value
         */
        explicit XLCellValue(Impl::XLCellValue& value);
    public:
        /**
         * @brief
         * @param other
         */
        XLCellValue(const XLCellValue& other) = delete;

        /**
         * @brief
         * @param other
         */
        XLCellValue(XLCellValue&& other) = delete;

        /**
         * @brief
         */
        virtual ~XLCellValue() = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLCellValue& operator=(const XLCellValue& other);

        /**
         * @brief
         * @tparam T
         * @param numberValue
         * @return
         */
        template<typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type* = nullptr>
        XLCellValue& operator=(T numberValue);

        /**
         * @brief
         * @tparam T
         * @param numberValue
         * @return
         */
        template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
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
        template<typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type* = nullptr>
        void Set(T numberValue);

        /**
         * @brief
         * @tparam T
         * @param numberValue
         */
        template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
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
        template<typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type* = nullptr>
        T Get() const;

        /**
         * @brief
         * @tparam T
         * @return
         */
        template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
        T Get() const;

        /**
         * @brief
         * @tparam T
         * @return
         */
        template<typename T, typename std::enable_if<
                std::is_constructible<T, char*>::value && !std::is_same<T, bool>::value,
                char*>::type* = nullptr>
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
        Impl::XLCellValue* m_value; /**< A pointer to the underlying implementation object, Impl::XLCellValue*/
    };

    /**
     * @details
     */
    template<typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type*>
    XLCellValue& XLCellValue::operator=(T numberValue) {

        Set(numberValue);
        return *this;
    }

    /**
     * @details
     */
    template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type*>
    XLCellValue& XLCellValue::operator=(T numberValue) {

        Set(numberValue);
        return *this;
    }

    /**
     * @details
     */
    template<typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type*>
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
    template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type*>
    void XLCellValue::Set(T numberValue) {

        SetFloat(numberValue);
    }

    /**
     * @details
     */
    template<typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type*>
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
    template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type*>
    T XLCellValue::Get() const {

        return GetFloat();
    }

    /**
     * @details
     */
    template<typename T, typename std::enable_if<
            std::is_constructible<T, char*>::value && !std::is_same<T, bool>::value,
            char*>::type*>
    T XLCellValue::Get() const {

        return T(GetString());
    }

} // namespace OpenXLSX

#endif //OPENXLSX_XLCELLVALUE_HPP
