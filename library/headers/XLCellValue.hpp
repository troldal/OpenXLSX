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
#include <memory>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlParser.hpp"

namespace OpenXLSX
{
    class XLCell;
    class XLSharedStrings;

    /**
     * @brief The XLValueType class is an enumeration of the possible cell value types.
     */
    enum class XLValueType { Empty, Boolean, Integer, Float, Error, String };

    /**
     * @brief The XLCellValue class represents the concept of a cell value. This can be in the form of a number
     * (an integer or a float), a string, a boolean or no value (empty).
     * @details The XLCellValue class encapsulates the concept of a cell value, i.e. the string, number, boolean or
     * empty value that a cell can contain. Other cell contents, such as formatting or formulas, are handled by separate
     * classes.
     *
     * ## Usage ##
     * An XLCellValue object is not created directly by the user, but rather accessed and/or modified through the XLCell
     * interface (the Value() member function).
     * The value can be set either through the Set() member function or the operator=(). Both have been overloaded to
     take
     * any integer or floating point value, bool value or string value (either std::string or char* string). In order to
     * set the cell value to empty, use the Clear() member function.
     *
     * To get the current value of an XLCellValue object, use the Get<>() member function. The Get<>() member function
     is a
     * template function, taking the value return type as template argument. It has been overloaded to take any integer
     * or floating point type, bool type or string type that can be constructed from a char* string (e.g. std::string or
     * std::string_view). It is also possible to get the value as a std::string, regardless of the valuetype, using the
     * AsString() member function.
     *
     * ## Example ##
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
     *
     * ## Developer Notes ##
     * The only member variable is a reference to the parent XLCell object. All functionality works by manipulating
     * the parent cell through the reference.
     */
    class OPENXLSX_EXPORT XLCellValue
    {
    public:
        //---------- PUBLIC MEMBER FUNCTIONS ----------//

        /**
         * @brief
         * @param cellNode
         * @param sharedStrings
         */
        explicit XLCellValue(const XMLNode& cellNode, XLSharedStrings* sharedStrings) noexcept;

        /**
         * @brief Copy constructor.
         * @details The copy constructor and copy assignment operator works differently for XLCellValue objects.
         * The copy constructor creates an exact copy of the object, with the same parent XLCell object. The copy
         * assignment operator only copies the underlying cell value and type attribute to the target object.
         * @param other object to be copied.
         * @note The default copy constructor has been used.
         */
        XLCellValue(const XLCellValue& other);

        /**
         * @brief Move constructor.
         * @param other the object to be moved.
         * @note The move constructor has been explicitly deleted. Move will not be allowed, as an XLCellValue must
         * always remain valid. Moving will invalidate the source object.
         */
        XLCellValue(XLCellValue&& other) noexcept = delete;

        /**
         * @brief Destructor.
         * @note The destructor uses the default implementation, as no pointer data members exist.
         */
        ~XLCellValue();

        /**
         * @brief Copy assignment operator.
         * @param other the object to copy values from.
         * @return A reference to the current object, with the new value.
         */
        XLCellValue& operator=(const XLCellValue& other);

        /**
         * @brief Move assignment operator.
         * @param other The object to be moved.
         * @return A reference to the current object, with the new value.
         * @note The move assignment operator has been explicitly deleted. Move will not be allowed, as an XLCellValue
         * must always remain valid. Moving will invalidate the source object.
         */
        XLCellValue& operator=(XLCellValue&& other) noexcept;

        /**
         * @brief Assignment operator
         * @tparam T The integer type to assign
         * @param numberValue The integer value to assign to the XLCellValue object.
         * @return A reference to the current object, with the new value.
         */
        template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type* = nullptr>
        XLCellValue& operator=(T numberValue);

        /**
         * @brief Assignment operator.
         * @tparam T the floating point type to assign.
         * @param numberValue The floating point value to assign to the XLCellValue object.
         * @return A reference to the current object, with the new value.
         */
        template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
        XLCellValue& operator=(T numberValue);

        /**
         * @brief Assignment operator.
         * @param stringValue A char* string to assign to the XLCellValue object.
         * @return A reference to the current object, with the new value.
         */
        XLCellValue& operator=(const char* stringValue);

        /**
         * @brief Assignment operator.
         * @param stringValue A std::string to assign to the XLCellValue object.
         * @return A reference to the current object, with the new value.
         */
        XLCellValue& operator=(const std::string& stringValue);

        /**
         * @brief Set the object to an integer value.
         * @tparam T The integer type to assign.
         * @param numberValue The integer value to assign.
         */
        template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type* = nullptr>
        void set(T numberValue);

        /**
         * @brief Set the object to a floating point value.
         * @tparam T The floating point type to assign.
         * @param numberValue The floating point value to assign.
         */
        template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
        void set(T numberValue);

        /**
         * @brief Set the object to a string value.
         * @param stringValue A char* string to assign.
         */
        void set(const char* stringValue);

        /**
         * @brief Set the object to a string value.
         * @param stringValue A std::string_view to assign.
         */
        void set(const std::string& stringValue);

        /**
         * @brief Get integer value.
         * @tparam T The integer type to get.
         */
        template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type* = nullptr>
        T get();

        /**
         * @brief Get floating point value.
         * @tparam T The floating point type to get.
         */
        template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
        T get();

        /**
         * @brief Get string value.
         * @tparam T The string type to get.
         */
        template<typename T,
                 typename std::enable_if<std::is_constructible<T, char*>::value && !std::is_same<T, bool>::value, char*>::type* = nullptr>
        T get();

        /**
         * @brief Get the value of the object as a string, regardless of the value type.
         * @return The value as a string.
         */
        std::string asString() const;

        /**
         * @brief Clear the cell value.
         */
        void clear();

        /**
         * @brief Get the value type of the cell.
         * @return An XLValueType object with the value type.
         */
        XLValueType valueType() const;

    private:
        //---------- PRIVATE MEMBER FUNCTIONS ----------//

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
         * @return
         */
        int64_t getInteger() const;

        /**
         * @brief
         * @return
         */
        bool getBoolean() const;

        /**
         * @brief
         * @return
         */
        long double getFloat() const;

        /**
         * @brief
         * @return
         */
        const char* getString() const;

        //---------- PRIVATE MEMBER VARIABLES ----------//

        std::unique_ptr<XMLNode> m_cellNode;
        XLSharedStrings*         m_sharedStrings;
    };

    //----------------------------------------------------------------------------------------------------------------------
    //           Implementation of Template Member Functions
    //----------------------------------------------------------------------------------------------------------------------

    /**
     * @details This is a template assignment operator for all integer-type parameters. The function calls the Set
     * template function and returns *this. It has been implemented using the std::enable_if template function which
     * wil SFINAE out any non-integer types.
     * @pre The input parameter, the XLCellValue object and the underlying xml object are valid.
     * @post The underlying xml object has been modified to hold the value of the input parameter and the type attribute
     * has been set accordingly.
     */
    template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type*>
    XLCellValue& XLCellValue::operator=(T numberValue)
    {
        set(numberValue);
        return *this;
    }

    /**
     * @details This is a template assignment operator for all floating point-type paremeters. The function calls the
     * Set() template function and returns *this. It has been implemented using the std::enable_if template function
     * which wil SFINAE out any non-floating point types.
     * @pre The input parameter, the XLCellValue object and the underlying xml object are valid.
     * @post The underlying xml object has been modified to hold the value of the input parameter and the type attribute
     * has been set accordingly.
     */
    template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type*>
    XLCellValue& XLCellValue::operator=(T numberValue)
    {
        set(numberValue);
        return *this;
    }

    /**
     * @details This is a template function for all integer-type parameters. It has been implemented using the
     * std::enable_if template function which will SFINAE out any non-integer types. Bools are handled as a special
     * case using the 'if constexpr' construct.
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type*>
    void XLCellValue::set(T numberValue)
    {
        if constexpr (std::is_same<T, bool>::value) {    // if bool
            setBoolean(numberValue);
        }
        else {    // if not bool
            setInteger(numberValue);
        }
    }

    /**
     * @details This is a template function for all floating point-type parameters. It has been implemented using the
     * std::enable_if template function which will SFINAE out any non-floating point types.
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type*>
    void XLCellValue::set(T numberValue)
    {
        setFloat(numberValue);
    }

    /**
     * @details This is a template function for all integer-type parameters. It has been implemented using the
     * std::enable_if template function which will SFINAE out any non-integer types. Bools are handled as a special
     * case using the 'if constexpr' construct.
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_integral<T>::value, int64_t>::type*>
    T XLCellValue::get()
    {
        if constexpr (std::is_same<T, bool>::value) {
            return getBoolean();
        }
        else {
            return static_cast<T>(getInteger());
        }
    }

    /**
     * @details This is a template function for all floating point-type parameters. It has been implemented using the
     * std::enable_if template function which will SFINAE out any non-floating point types.
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type*>
    T XLCellValue::get()
    {
        return static_cast<T>(getFloat());
    }

    /**
     * @details This is a template function for all types that can be constructed from a char* string, including
     * std::string and std::string_view. All other types are SFINAE'd out using std::enable_if.
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_constructible<T, char*>::value && !std::is_same<T, bool>::value, char*>::type*>
    T XLCellValue::get()
    {
        return T(getString());
    }
}    // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_XLCELLVALUE_HPP
