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

#ifndef OPENXLSX_IMPL_XLCELLVALUE_H
#define OPENXLSX_IMPL_XLCELLVALUE_H

#include "XLDefinitions.h"
#include "XLCellType_Impl.h"
#include "XLException.h"
#include "XLXml_Impl.h"
#include "XLCellValue.h"
#include "XLSharedStrings_Impl.h"

#include <string>
#include <string_view>

namespace OpenXLSX::Impl
{

    class XLCell;

    //======================================================================================================================
    //========== XLCellValue Class =========================================================================================
    //======================================================================================================================

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
     * The value can be set either through the Set() member function or the operator=(). Both have been overloaded to take
     * any integer or floating point value, bool value or string value (either std::string or char* string). In order to
     * set the cell value to empty, use the Clear() member function.
     *
     * To get the current value of an XLCellValue object, use the Get<>() member function. The Get<>() member function is a
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
    class XLCellValue
    {
        friend class OpenXLSX::XLCellValue;

        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor
         * @param parent A reference to the parent XLCell object.
         */
        explicit XLCellValue(XLCell& parent) noexcept;

        /**
         * @brief Copy constructor.
         * @details The copy constructor and copy assignment operator works differently for XLCellValue objects.
         * The copy constructor creates an exact copy of the object, with the same parent XLCell object. The copy
         * assignment operator only copies the underlying cell value and type attribute to the target object.
         * @param other object to be copied.
         * @note The default copy constructor has been used.
         */
        XLCellValue(const XLCellValue& other) = delete;

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
        virtual ~XLCellValue() = default;

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
         * @note The move assignment operator has been explicitly deleted. Move will not be allowed, as an XLCellValue must
         * always remain valid. Moving will invalidate the source object.
         */
        XLCellValue& operator=(XLCellValue&& other) noexcept = delete;

        /**
         * @brief Assignment operator
         * @tparam T The integer type to assign
         * @param numberValue The integer value to assign to the XLCellValue object.
         * @return A reference to the current object, with the new value.
         */
        template<typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type* = nullptr>
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
         * @brief Get the value of the object as a string, regardless of the value type.
         * @return The value as a string.
         */
        std::string AsString() const;

        /**
         * @brief Set the object to an integer value.
         * @tparam T The integer type to assign.
         * @param numberValue The integer value to assign.
         */
        template<typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type* = nullptr>
        void Set(T numberValue);

        /**
         * @brief Set the object to a floating point value.
         * @tparam T The floating point type to assign.
         * @param numberValue The floating point value to assign.
         */
        template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
        void Set(T numberValue);

        /**
         * @brief Set the object to a string value.
         * @param stringValue A char* string to assign.
         */
        void Set(const char* stringValue);

        /**
         * @brief Set the object to a string value.
         * @param stringValue A std::string_view to assign.
         */
        void Set(const std::string& stringValue);

        /**
         * @brief Get integer value.
         * @tparam T The integer type to get.
         */
        template<typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type* = nullptr>
        T Get();

        /**
         * @brief Get floating point value.
         * @tparam T The floating point type to get.
         */
        template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type* = nullptr>
        T Get();

        /**
         * @brief Get string value.
         * @tparam T The string type to get.
         */
        template<typename T, typename std::enable_if<
                std::is_constructible<T, char*>::value && !std::is_same<T, bool>::value,
                char*>::type* = nullptr>
        T Get();

        /**
         * @brief Clear the cell value.
         */
        void Clear();

        /**
         * @brief Get the value type of the cell.
         * @return An XLValueType object with the value type.
         */
        XLValueType ValueType() const;

        //----------------------------------------------------------------------------------------------------------------------
        //           Protected Member Functions
        //----------------------------------------------------------------------------------------------------------------------
    protected:

        /**
         * @brief Get the cell type of the cell.
         * @return An XLCellType object corresponding to the cell type.
         */
        XLCellType CellType() const;

        /**
         * @brief Get a reference to the parent cell of the XLCellValue object.
         * @return A reference to the parent XLCell object.
         */
        XLCell* Cell();

        /**
         * @brief Get a const reference to the parent cell of the XLCellValue object.
         * @return A const reference to the parent XLCell object.
         */
        const XLCell* Cell() const;

        /**
         * @brief Get a pointer to the value node in the underlying XML file.
         * @return A pointer to an XMLNode object corresponding to the value node.
         */
        XMLNode ValueNode();

        /**
         * @brief Get a const pointer to the value node in the underlying XML file.
         * @return A const pointer to an XMLNode object corresponding to the value node.
         */
        const XMLNode ValueNode() const;

        /**
         * @brief Confirm whether or not a value node exists.
         * @return true if a value node exists; otherwise false.
         */
        bool HasValueNode() const;

        /**
         * @brief Set the value of the value node.
         * @param value A std::string with the value.
         */
        void SetValueNode(std::string_view value);

        /**
         * @brief Delete the value node in the underlying XML file.
         */
        void DeleteValueNode();

        /**
         * @brief Create a new value node.
         * @return A pointer to the XMLNode object for the value node.
         */
        XMLNode CreateValueNode();

        /**
         * @brief Get a pointer to the type attribute in the underlying XML file.
         * @return A pointer to an XMLAttribute object corresponding to the type attribute; nullptr if it
         * doesn't exist.
         */
        XMLAttribute TypeAttribute();

        /**
         * @brief Get a const pointer to the type attribute in the underlying XML file.
         * @return A const pointer to an XMLAttribute object corresponding to the type attribute; nullptr if it
         * doesn't exist.
         */
        const XMLAttribute TypeAttribute() const;

        /**
         * @brief Confirm whether or not a type attribute exists.
         * @return true if a type attribute exists; otherwise false.
         */
        bool HasTypeAttribute() const;

        /**
         * @brief Set the value of the type attribute.
         * @param typeString A std::string with the value.
         */
        void SetTypeAttribute(const std::string& typeString);

        /**
         * @brief Delete the type attribute in the underlying XML file.
         */
        void DeleteTypeAttribute();

        /**
         * @brief Create a new type attribute.
         * @return A pointer to the XMLAttribute object for the type attribute.
         */
        XMLAttribute CreateTypeAttribute();


        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    private:

        /**
         * @brief Determine if the input string holds an integer or a floating point value.
         * @param numberString The string holding a number.
         * @return An XLNumberType::Integer or XLNumberType::Float
         */
        XLNumberType DetermineNumberType(const std::string& numberString) const;

        /**
         * @brief Get the xml node with the requested shared string.
         * @param index The index of the requested shared string.
         * @return An xml node with the shared string.
         */
        XMLNode SharedStringNode(unsigned long index) const;

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


        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Variables
        //----------------------------------------------------------------------------------------------------------------------

    private:
        XLCell& m_parentCell; /**< A reference to the parent XLCell object. */
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
    template<typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type*>
    XLCellValue& XLCellValue::operator=(T numberValue) {

        Set(numberValue);
        return *this;
    }

    /**
     * @details This is a template assignment operator for all floating point-type paremeters. The function calls the
     * Set() template function and returns *this. It has been implemented using the std::enable_if template function which
     * wil SFINAE out any non-floating point types.
     * @pre The input parameter, the XLCellValue object and the underlying xml object are valid.
     * @post The underlying xml object has been modified to hold the value of the input parameter and the type attribute
     * has been set accordingly.
     */
    template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type*>
    XLCellValue& XLCellValue::operator=(T numberValue) {

        Set(numberValue);
        return *this;
    }

    /**
     * @details This is a template function for all integer-type parameters. It has been implemented using the
     * std::enable_if template function which will SFINAE out any non-integer types. Bools are handled as a special
     * case using the 'if constexpr' construct.
     * @pre
     * @post
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
     * @details This is a template function for all floating point-type parameters. It has been implemented using the
     * std::enable_if template function which will SFINAE out any non-floating point types.
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type*>
    void XLCellValue::Set(T numberValue) {

        SetFloat(numberValue);
    }

    /**
     * @details This is a template function for all integer-type parameters. It has been implemented using the
     * std::enable_if template function which will SFINAE out any non-integer types. Bools are handled as a special
     * case using the 'if constexpr' construct.
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type*>
    T XLCellValue::Get() {

        if constexpr (std::is_same<T, bool>::value) {
            return GetBoolean();
        }
        else {
            return GetInteger();
        }
    }

    /**
     * @details This is a template function for all floating point-type parameters. It has been implemented using the
     * std::enable_if template function which will SFINAE out any non-floating point types.
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type*>
    T XLCellValue::Get() {

        return GetFloat();
    }

    /**
     * @details This is a template function for all types that can be constructed from a char* string, including
     * std::string and std::string_view. All other types are SFINAE'd out using std::enable_if.
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<
            std::is_constructible<T, char*>::value && !std::is_same<T, bool>::value,
            char*>::type*>
    T XLCellValue::Get() {

        return T(GetString());
    }
} // namespace OpenXLSX::Impl

#endif //OPENXLSX_XLCELLVALUE_H
