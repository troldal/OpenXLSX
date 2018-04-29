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

#ifndef OPENXLEXE_XLCELLVALUE_H
#define OPENXLEXE_XLCELLVALUE_H

#include "XLValueString.h"
#include "XLValueNumber.h"
#include "XLValueBoolean.h"
#include "XLValueEmpty.h"
#include "XLValueError.h"
#include "XLCellType.h"
#include "XLCell.h"

#include <string>

namespace OpenXLSX
{

//======================================================================================================================
//========== XLCellValue Class =========================================================================================
//======================================================================================================================

    /**
     * @brief The XLCellValue class represents the concept of a cell value. This can be in the form of a number
     * (an integer or a float), a string, a boolean or no value (empty).
     */
    class XLCellValue
    {
        friend class XLValue;
        friend class XLValueString;
        friend class XLValueNumber;
        friend class XLValueBoolean;
        friend class XLValueEmpty;
        friend class XLValueError;


//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor
         * @param parent A reference to the parent XLCell object.
         */
        explicit XLCellValue(XLCell &parent);

        /**
         * @brief Copy constructor.
         * @param other object to be copied.
         * @note The copy constructor has been explicitly deleted.
         */
        XLCellValue(const XLCellValue &other) = delete;

        /**
         * @brief Move constructor.
         * @param other the object to be moved.
         * @note The move constructor has been explicitly deleted.
         */
        XLCellValue(XLCellValue &&other) noexcept = delete;

        /**
         * @brief Destructor.
         * @note The destructor uses the default implementation
         */
        virtual ~XLCellValue() = default;

        /**
         * @brief Copy assignment operator.
         * @param other the object to copy values from.
         * @return A reference to the current object, with the new value.
         */
        XLCellValue &operator=(const XLCellValue &other);

        /**
         * @brief Move assignment operator.
         * @param other The object to be moved.
         * @return A reference to the current object, with the new value.
         */
        XLCellValue &operator=(XLCellValue &&other) noexcept;

        /**
         * @brief Assignment operator.
         * @param boolValue The boolean value to assign to the XLCellValue object.
         * @return A reference to the current object, with the new value.
         */
        XLCellValue &operator=(XLBool boolValue);

        /**
         * @brief Assignment operator.
         * @return A reference to the current object, with the new value.
         * @note This operator, taking a bool as parameter, has been explicitly deleted.
         */
        XLCellValue &operator=(bool) = delete;

        /**
         * @brief Assignment operator
         * @tparam T The integer type to assign
         * @param numberValue The integer value to assign to the XLCellValue object.
         * @return A reference to the current object, with the new value.
         */
        template<typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type * = nullptr>
        XLCellValue &operator=(T numberValue);

        /**
         * @brief Assignment operator.
         * @tparam T the floating point type to assign.
         * @param numberValue The floating point value to assign to the XLCellValue object.
         * @return A reference to the current object, with the new value.
         */
        template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type * = nullptr>
        XLCellValue &operator=(T numberValue);

        /**
         * @brief Assingment operator.
         * @param stringValue A char* string to assign to the XLCellValue object.
         * @return A reference to the current object, with the new value.
         */
        XLCellValue &operator=(const char *stringValue);

        /**
         * @brief Assignment operator.
         * @param stringValue A std::string to assign to the XLCellValue object.
         * @return A reference to the current object, with the new value.
         */
        XLCellValue &operator=(const std::string &stringValue);

        /**
         * @brief Get the value of the object as a boolean.
         * @return The value as an XLBool object.
         */
        XLBool Boolean() const;

        /**
         * @brief Get the value of the object as a floating point number.
         * @return The value as a long double.
         */
        long double Float() const;

        /**
         * @brief Get the value of the object as an integer number.
         * @return The value as an integer.
         */
        long long int Integer() const;
        /**
         * @brief Get the value of the object as a string.
         * @return The value as a string.
         */
        const std::string &String() const;

        /**
         * @brief Get the value of the object as a string, regardless of the value type.
         * @return The value as a string.
         */
        std::string AsString() const;

        /**
         * @brief Set the object a boolean value.
         * @param boolValue an XLBool value.
         */
        void Set(XLBool boolValue);

        /**
         * @brief Set the object to a boolean value.
         * @note This method has been explicitly deleted.
         */
        void Set(bool) = delete;

        /**
         * @brief Set the object to an integer value.
         * @tparam T The integer type to assign.
         * @param numberValue The integer value to assign.
         */
        template<typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type * = nullptr>
        void Set(T numberValue);

        /**
         * @brief Set the object to a floating point value.
         * @tparam T The floating point type to assign.
         * @param numberValue The floating point value to assign.
         */
        template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type * = nullptr>
        void Set(T numberValue);

        /**
         * @brief Set the object to a string value.
         * @param stringValue A char* string to assign.
         */
        void Set(const char *stringValue);

        /**
         * @brief Set the object to a string value.
         * @param stringValue A std::string to assign.
         */
        void Set(const std::string &stringValue);

        /**
         * @brief Clear the cell value.
         */
        void SetEmpty();

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
         * @brief Get a string corresponding to the type attribute in the underlying XML file.
         * @return A std::string with the type string.
         */
        std::string TypeString() const;

        /**
         * @brief Get the cell type of the cell.
         * @return An XLCellType object corresponding to the cell type.
         */
        XLCellType CellType() const;

        /**
         * @brief Get a reference to the parent cell of the XLCellValue object.
         * @return A reference to the parent XLCell object.
         */
        XLCell *ParentCell();

        /**
         * @brief Get a const reference to the parent cell of the XLCellValue object.
         * @return A const reference to the parent XLCell object.
         */
        const XLCell *ParentCell() const;

        /**
         * @brief Get a pointer to the value node in the underlying XML file.
         * @return A pointer to an XMLNode object corresponding to the value node.
         */
        XMLNode *ValueNode();

        /**
         * @brief Get a const pointer to the value node in the underlying XML file.
         * @return A const pointer to an XMLNode object corresponding to the value node.
         */
        const XMLNode *ValueNode() const;

        /**
         * @brief Confirm whether or not a value node exists.
         * @return true if a value node exists; otherwise false.
         */
        bool HasValueNode() const;

        /**
         * @brief Set the value of the value node.
         * @param value A std::string with the value.
         */
        void SetValueNode(const std::string &value);

        /**
         * @brief Delete the value node in the underlying XML file.
         */
        void DeleteValueNode();

        /**
         * @brief Create a new value node.
         * @return A pointer to the XMLNode object for the value node.
         */
        XMLNode *CreateValueNode();

        /**
         * @brief Get a pointer to the type attribute in the underlying XML file.
         * @return A pointer to an XMLAttribute object corresponding to the type attribute; nullptr if it
         * doesn't exist.
         */
        XMLAttribute *TypeAttribute();

        /**
         * @brief Get a const pointer to the type attribute in the underlying XML file.
         * @return A const pointer to an XMLAttribute object corresponding to the type attribute; nullptr if it
         * doesn't exist.
         */
        const XMLAttribute *TypeAttribute() const;

        /**
         * @brief Confirm whether or not a type attribute exists.
         * @return true if a type attribute exists; otherwise false.
         */
        bool HasTypeAttribute() const;

        /**
         * @brief Set the value of the type attribute.
         * @param typeString A std::string with the value.
         */
        void SetTypeAttribute(const std::string &typeString);

        /**
         * @brief Delete the type attribute in the underlying XML file.
         */
        void DeleteTypeAttribute();

        /**
         * @brief Create a new type attribute.
         * @return A pointer to the XMLAttribute object for the type attribute.
         */
        XMLAttribute *CreateTypeAttribute();


//----------------------------------------------------------------------------------------------------------------------
//           Private Member Functions
//----------------------------------------------------------------------------------------------------------------------

    private:

        /**
         * @brief Initialize the object based on the contents of the XML file.
         */
        void Initialize();


//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:
        std::unique_ptr<XLValue> m_value; /**< A pointer to the value object. */
        XLCell &m_parentCell; /**< A reference to the parent XLCell object. */
    };

//----------------------------------------------------------------------------------------------------------------------
//           Implementation of Template Member Functions
//----------------------------------------------------------------------------------------------------------------------

    /**
     * @details This is a template assignment operator for all integer-type paremeters. The function calls the Set
     * template function and returns *this.
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type *>
    XLCellValue &XLCellValue::operator=(T numberValue)
    {
        Set(static_cast<long long int>(numberValue));
        return *this;
    }

    /**
     * @details This is a template assignment operator for all floating point-type paremeters. The function calls the
     * Set template function and returns *this.
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type *>
    XLCellValue &XLCellValue::operator=(T numberValue)
    {
        Set(static_cast<long double>(numberValue));
        return *this;
    }

    /**
     * @details This is a template function for all integer-type parameters. If the current cell value is not already
     * of integer ot floating point type (i.e. number type), set the m_value variable accordingly. Call the Set
     * method of the m_value member variable.
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_integral<T>::value, long long int>::type *>
    void XLCellValue::Set(T numberValue)
    {
        if (ValueType() != XLValueType::Integer || ValueType() != XLValueType::Float)
            m_value = std::make_unique<XLValueNumber>(*this);

        static_cast<XLValueNumber &>(*m_value).Set(static_cast<long long int>(numberValue));

        ParentCell()->SetModified();

    }

    /**
     * @details This is a template function for all floating point-type parameters. If the current cell value is not
     * already of integer ot floating point type (i.e. number type), set the m_value variable accordingly. Call the Set
     * method of the m_value member variable.
     * @pre
     * @post
     */
    template<typename T, typename std::enable_if<std::is_floating_point<T>::value, long double>::type *>
    void XLCellValue::Set(T numberValue)
    {
        if (ValueType() != XLValueType::Integer || ValueType() != XLValueType::Float)
            m_value = std::make_unique<XLValueNumber>(*this);

        static_cast<XLValueNumber &>(*m_value).Set(static_cast<long double>(numberValue));

        ParentCell()->SetModified();
    }
}

#endif //OPENXLEXE_XLCELLVALUE_H
