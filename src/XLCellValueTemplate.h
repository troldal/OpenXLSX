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

#ifndef OPENXLEXE_XLCELLVALUETEMPLATE_H
#define OPENXLEXE_XLCELLVALUETEMPLATE_H

#include "XLValueString.h"
#include "XLValueNumber.h"
#include "XLValueBoolean.h"
#include "XLValueEmpty.h"
#include "XLValueError.h"
#include "XLCellType.h"
#include "XLCell.h"
#include "Utilities/XML/XML.h"

#include <string>
#include <type_traits>

namespace OpenXLSX
{

//======================================================================================================================
//========== XLCellValue Class =========================================================================================
//======================================================================================================================

    /**
     * @brief The XLCellValue class represents the concept of a cell value. This can be in the form of a number
     * (an integer or a float), a string, a boolean or no value (empty).
     */
    template<typename T>
    class XLCellValueTemplate
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
        explicit XLCellValueTemplate(XLCell &parent, T value = 0.0);

        /**
         * @brief Copy constructor.
         * @param other object to be copied.
         * @note The copy constructor has been explicitly deleted.
         */
        XLCellValueTemplate(const XLCellValue &other) = delete;

        /**
         * @brief Move constructor.
         * @param other the object to be moved.
         * @note The move constructor has been explicitly deleted.
         */
        XLCellValueTemplate(XLCellValue &&other) noexcept = delete;

        /**
         * @brief Destructor.
         * @note The destructor uses the default implementation
         */
        virtual ~XLCellValueTemplate() = default;

        /**
         * @brief Copy assignment operator.
         * @param other the object to copy values from.
         * @return A reference to the current object, with the new value.
         */
        XLCellValueTemplate &operator=(const XLCellValueTemplate &other);

        /**
         * @brief Move assignment operator.
         * @param other The object to be moved.
         * @return A reference to the current object, with the new value.
         */
        XLCellValueTemplate &operator=(XLCellValueTemplate &&other) noexcept;

        /**
         * @brief Assignment operator.
         * @param boolValue The boolean value to assign to the XLCellValue object.
         * @return A reference to the current object, with the new value.
         */
        XLCellValueTemplate &operator=(XLBool boolValue);

        /**
         * @brief Assignment operator.
         * @return A reference to the current object, with the new value.
         * @note This operator, taking a bool as parameter, has been explicitly deleted.
         */
        XLCellValueTemplate &operator=(bool) = delete;

        /**
         * @brief Assignment operator
         * @tparam U The integer type to assign
         * @param numberValue The integer value to assign to the XLCellValue object.
         * @return A reference to the current object, with the new value.
         */
        template<typename U, typename std::enable_if<std::is_integral<U>::value, long long int>::type * = nullptr>
        XLCellValueTemplate &operator=(U numberValue);

        /**
         * @brief Assignment operator.
         * @tparam T the floating point type to assign.
         * @param numberValue The floating point value to assign to the XLCellValue object.
         * @return A reference to the current object, with the new value.
         */
        template<typename U, typename std::enable_if<std::is_floating_point<U>::value, long double>::type * = nullptr>
        XLCellValueTemplate &operator=(U numberValue);

        /**
         * @brief Assingment operator.
         * @param stringValue A char* string to assign to the XLCellValue object.
         * @return A reference to the current object, with the new value.
         */
        XLCellValueTemplate &operator=(const char *stringValue);

        /**
         * @brief Assignment operator.
         * @param stringValue A std::string to assign to the XLCellValue object.
         * @return A reference to the current object, with the new value.
         */
        XLCellValueTemplate &operator=(const std::string &stringValue);

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
        template<typename U, typename std::enable_if<std::is_integral<U>::value, long long int>::type * = nullptr>
        void Set(U numberValue);

        /**
         * @brief Set the object to a floating point value.
         * @tparam T The floating point type to assign.
         * @param numberValue The floating point value to assign.
         */
        template<typename U, typename std::enable_if<std::is_floating_point<U>::value, long double>::type * = nullptr>
        void Set(U numberValue);

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
        void SetValueNode(const std::string &value);

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
        void SetTypeAttribute(const std::string &typeString);

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
         * @brief Initialize the object based on the contents of the XML file.
         */
        void Initialize();


//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:
        T m_value; /**< A pointer to the value object. */
        XLCell &m_parentCell; /**< A reference to the parent XLCell object. */
    };

//----------------------------------------------------------------------------------------------------------------------
//           Implementation of Template Member Functions
//----------------------------------------------------------------------------------------------------------------------

    /**
     * @details The constructor sets the m_value to nullptr and the m_parentCell to the value of the input parameter.
     * Subsequently, the Initialize member function is called, which sets up the m_value variable.
     * @pre The parent input parameter is a valid XLCell object.
     * @post A valid XLCellValueTemplate object has been constructed.
     */
    template<typename T>
    XLCellValueTemplate<T>::XLCellValueTemplate(XLCell &parent, T value)
        : m_value(value),
          m_parentCell(parent)
    {
        Initialize();
    }

    /**
     * @details Initializes the object, by creating the right value object, based on the cell type, which is ascertained
     * based on the content of the underlying XML file.
     * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
     * @post The object has been initialized with the correct value type.
     */
    template<typename T>
    void XLCellValueTemplate<T>::Initialize()
    { /*
    switch (CellType()) {
        case XLCellType::Empty:
            m_value = make_unique<XLValueEmpty>(*this);
            break;

        case XLCellType::Number:
            m_value = make_unique<XLValueNumber>(*this);
            break;

        case XLCellType::String:
            m_value = make_unique<XLValueString>(*this);
            break;

        case XLCellType::Boolean:
            m_value = make_unique<XLValueBoolean>(*this);
            break;

        case XLCellType::Error:
            m_value = make_unique<XLValueError>(*this);
            break;
    } */
    }

    /**
     * @details The copy assignment operator sets the value node and type attribute to the corresponding items in
     * the object being assigned from. Subsequently, the Initialize method is called, which initializes the object
     * based on the underlying XML file.
     * @pre
     * @post
     */
    template<typename T>
    XLCellValueTemplate<T> &XLCellValueTemplate<T>::operator=(const XLCellValueTemplate<T> &other)
    {
        SetValueNode(!other.ValueNode() ? "" : other.ValueNode().text().get());
        SetTypeAttribute(other.TypeString());

        Initialize();
        ParentCell()->SetModified();
        return *this;
    }

    /**
     * @details The copy assignment operator sets the value node and type attribute to the corresponding items in
     * the object being assigned from. Subsequently, the Initialize method is called, which initializes the object
     * based on the underlying XML file. This is identical to the copy assignment operator.
     * @pre
     * @post
     * @todo Implement a real move assignment operator
     */
    template<typename T>
    XLCellValueTemplate<T> &XLCellValueTemplate<T>::operator=(XLCellValueTemplate<T> &&other) noexcept
    {
        SetValueNode(!other.ValueNode() ? "" : other.ValueNode().value());
        SetTypeAttribute(other.TypeString());

        Initialize();
        ParentCell()->SetModified();
        return *this;
    }

    /**
     * @details The assignment operator taking a XLBool object as a parameter, calls the corresponding Set method
     * and returns the resulting object.
     * @pre
     * @post
     */
    template<typename T>
    XLCellValueTemplate<T> &XLCellValueTemplate<T>::operator=(XLBool boolValue)
    {
        Set(boolValue);
        return *this;
    }

    /**
     * @details The assignment operator taking a std::tring object as a parameter, calls the corresponding Set method
     * and returns the resulting object.
     * @pre
     * @post
     */
    template<typename T>
    XLCellValueTemplate<T> &XLCellValueTemplate<T>::operator=(const std::string &stringValue)
    {
        Set(stringValue);
        return *this;
    }

    /**
     * @details The assignment operator taking a char* string as a parameter, calls the corresponding Set method
     * and returns the resulting object.
     * @pre
     * @post
     */
    template<typename T>
    XLCellValueTemplate<T> &XLCellValueTemplate<T>::operator=(const char *stringValue)
    {
        Set(stringValue);
        return *this;
    }

    /**
     * @details If the value type is already a boolean, the value will be set to the new value. Otherwise, the m_value
     * member variable will be set to an XLBoolean object with the given value.
     * @pre
     * @post
     */
    template<typename T>
    void XLCellValueTemplate<T>::Set(XLBool boolValue)
    {
        if (ValueType() == XLValueType::Boolean)
            static_cast<XLValueBoolean &>(*m_value).Set(boolValue);
        else
            m_value = std::make_unique<XLValueBoolean>(boolValue, *this);

        ParentCell()->SetModified();
    }

    /**
     * @details If the value type is already a String, the value will be set to the new value. Otherwise, the m_value
     * member variable will be set to an XLString object with the given value.
     * @pre
     * @post
     */
    template<typename T>
    void XLCellValueTemplate<T>::Set(const std::string &stringValue)
    {
        if (ValueType() == XLValueType::String)
            static_cast<XLValueString &>(*m_value).Set(stringValue);
        else
            m_value = std::make_unique<XLValueString>(stringValue, *this);

        ParentCell()->SetModified();
    }

    /**
     * @details Converts the char* parameter to a std::string and calls the corresponding Set method.
     * @pre
     * @post
     */
    template<typename T>
    void XLCellValueTemplate<T>::Set(const char *stringValue)
    {
        Set(std::string(stringValue));
    }

    /**
     * @details Returns the boolean value of the cell, provided the value type is 'Boolean'.
     * @exception XLException If the value type is not 'Boolean', an XLException will be thrown.
     * @pre The m_value member variable is valid and the value type is 'Boolean'
     * @post The current object, and any associated objects, are unchanged.
     */
    template<typename T>
    XLBool XLCellValueTemplate<T>::Boolean() const
    {
        if (ValueType() != XLValueType::Boolean) throw XLException("Cell value is not Boolean");
        return static_cast<XLValueBoolean &>(*m_value).Boolean();
    }

    /**
     * @details Returns the float value of the cell, provided the value type is 'Float'.
     * @exception XLException If the value type is not 'Float', an XLException will be thrown.
     * @pre The m_value member variable is valid and the value type is 'Float'
     * @post The current object, and any associated objects, are unchanged.
     */
    template<typename T>
    long double XLCellValueTemplate<T>::Float() const
    {
        if (ValueType() != XLValueType::Float) throw XLException("Cell value is not Float");
        return static_cast<XLValueNumber &>(*m_value).Float();
    }

    /**
     * @details Returns the integer value of the cell, provided the value type is 'Integer'.
     * @exception XLException If the value type is not 'Integer', an XLException will be thrown.
     * @pre The m_value member variable is valid and the value type is 'Integer'
     * @post The current object, and any associated objects, are unchanged.
     */
    template<typename T>
    long long int XLCellValueTemplate<T>::Integer() const
    {
        if (ValueType() != XLValueType::Integer) throw XLException("Cell value is not Integer");
        return static_cast<XLValueNumber &>(*m_value).Integer();
    }

    /**
     * @details Returns the string value of the cell, provided the value type is 'String'.
     * @exception XLException If the value type is not 'String', an XLException will be thrown.
     * @pre The m_value member variable is valid and the value type is 'String'
     * @post The current object, and any associated objects, are unchanged.
     */
    template<typename T>
    const std::string &XLCellValueTemplate<T>::String() const
    {
        if (ValueType() != XLValueType::String) throw XLException("Cell value is not String");
        return static_cast<XLValueString &>(*m_value).String();
    }

    /**
     * @details Return the cell value as a string, by calling the AsString method of the m_value member variable.
     * @pre The m_value member variable is valid.
     * @post The current object, and any associated objects, are unchanged.
     */
    template<typename T>
    std::string XLCellValueTemplate<T>::AsString() const
    {
        return (*m_value).AsString();
    }

    /**
     * @details Determine the value type, based on the cell type, and return the corresponding XLValueType object.
     * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
     * @post The current object, and any associated objects, are unchanged.
     */
    template<typename T>
    XLValueType XLCellValueTemplate<T>::ValueType() const
    {
        switch (CellType()) {
            case XLCellType::Empty:
                return XLValueType::Empty;

            case XLCellType::Error:
                return XLValueType::Error;

            case XLCellType::Number: {
                if (dynamic_cast<XLValueNumber &>(*m_value).NumberType() == XLNumberType::Integer)
                    return XLValueType::Integer;
                else
                    return XLValueType::Float;
            }

            case XLCellType::Boolean:
                return XLValueType::Boolean;

            default:
                return XLValueType::String;
        }
    }

    /**
     * @details Determine the cell type, based on the contents of the underlying XML file, and returns the corresponding
     * XLCellType object.
     * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
     * @post The current object, and any associated objects, are unchanged.
     */
    template<typename T>
    XLCellType XLCellValueTemplate<T>::CellType() const
    {
        // Determine the type of the Cell
        // If neither a Type attribute or a value node is present, the cell is empty.
        if (!m_parentCell.HasTypeAttribute() && !HasValueNode()) {
            return XLCellType::Empty;
        }

            // If a Type attribute is not present, but a value node is, the cell contains a number.
        else if (!m_parentCell.HasTypeAttribute() && HasValueNode()) {
            return XLCellType::Number;
        }

            // If the cell is of type "s", the cell contains a shared string.
        else if (m_parentCell.HasTypeAttribute() && std::string(m_parentCell.TypeAttribute().value()) == "s") {
            return XLCellType::String;
        }

            // If the cell is of type "inlineStr", the cell contains an inline string.
        else if (m_parentCell.HasTypeAttribute() && std::string(m_parentCell.TypeAttribute().value()) == "inlineStr") {
            return XLCellType::String;
        }

            // If the cell is of type "str", the cell contains an ordinary string.
        else if (m_parentCell.HasTypeAttribute() && std::string(m_parentCell.TypeAttribute().value()) == "str") {
            return XLCellType::String;
        }

            // If the cell is of type "b", the cell contains a boolean.
        else if (m_parentCell.HasTypeAttribute() && std::string(m_parentCell.TypeAttribute().value()) == "b") {
            return XLCellType::Boolean;
        }

            // Otherwise, the cell contains an error.
        else
            return XLCellType::Error; //the m_typeAttribute has the ValueAsString "e"
    }

    /**
     * @details Returns the type string by calling the TypeString method of the m_value member variable.
     * @pre The m_value member variable is valid.
     * @post The current object, and any associated objects, are unchanged.
     */
    template<typename T>
    std::string XLCellValueTemplate<T>::TypeString() const
    {
        if (std::is_same<T, bool>) return "b";
        else if (std::is_arithmetic<T>) return "";
        return (*m_value).TypeString();
    }

    /**
     * @details Returns the value of the m_parentCell member variable.
     * @pre The parent XLCell object is valid.
     * @post The current object, and any associated objects, are unchanged.
     */
    template<typename T>
    XLCell *XLCellValueTemplate<T>::ParentCell()
    {
        return &m_parentCell;
    }

    /**
     * @details Returns the value of the m_parentCell member variable.
     * @pre The parent XLCell object is valid.
     * @post The current object, and any associated objects, are unchanged.
     */
    template<typename T>
    const XLCell *XLCellValueTemplate<T>::ParentCell() const
    {
        return &m_parentCell;
    }

    /**
     * @details If the input value is an empty string, object will be set to empty. Otherwise, a value node will be
     * created (if it doesn't exists) and the value set to the input value.
     * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
     * @post The cell is either set to empty, or a value node exists and has been set to the input value.
     * @note If the input string is empty, the cell type will be set to empty.
     */
    template<typename T>
    void XLCellValueTemplate<T>::SetValueNode(const std::string &value)
    {
        ParentCell()->CreateValueNode();
        ValueNode().text().set(value.c_str());
    }

    /**
     * @details If the type string is empty, the type attribute will be deleted. Otherwise, the type attribute will be
     * created and set to the value of the type string.
     * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
     * @post The type attribute is either deleted or has been set to the value of the input type string.
     */
    template<typename T>
    void XLCellValueTemplate<T>::SetTypeAttribute(const std::string &typeString)
    {
        if (typeString.empty())
            DeleteTypeAttribute();
        else {
            CreateTypeAttribute();
            ParentCell()->CellNode().attribute("t").set_value(typeString.c_str());
        }
    }

    /**
     * @details If a type attribute exists, it will be deleted from the XML file.
     * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
     * @post No type attribute for the cell exists in the XML file.
     */
    template<typename T>
    void XLCellValueTemplate<T>::DeleteTypeAttribute()
    {
        if (HasTypeAttribute()) ParentCell()->CellNode().remove_attribute("t");
    }

    /**
     * @details If a type attribute does not exist, a new one will be created, added to the XML document and returned.
     * Otherwise a pointer to the existing type attribute will be returned.
     * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
     * @post An empty type attribute has been added to the XML file.
     * @note An empty type attribute may not be valid in Excel. Therefore, the value must be set immediately afterwards.
     */
    template<typename T>
    XMLAttribute XLCellValueTemplate<T>::CreateTypeAttribute()
    {
        if (!HasTypeAttribute())
            ParentCell()->CellNode().append_attribute("t") = "";

        return ParentCell()->CellNode().attribute("t");
    }

    /**
     * @details Obtain a pointer to the corresponding XMLNode object, requesting child node "v" from the cell node. If
     * no type node exists, a nullptr will be returned.
     * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
     * @post The current object, and any associated objects, are unchanged.
     */
    template<typename T>
    XMLNode XLCellValueTemplate<T>::ValueNode()
    {
        if (!HasValueNode())
            ParentCell()->CreateValueNode();
        return ParentCell()->CellNode().child("v");
    }

    /**
     * @details Obtain a pointer to the corresponding XMLNode object, requesting child node "v" from the cell node. If
     * no type node exists, a nullptr will be returned.
     * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
     * @post The current object, and any associated objects, are unchanged.
     */
    template<typename T>
    const XMLNode XLCellValueTemplate<T>::ValueNode() const
    {
        return ParentCell()->CellNode().child("v");
    }

    /**
     * @details Determines the existance of the value node, by checking if a nullptr is returned.
     * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
     * @post The current object, and any associated objects, are unchanged.
     */
    template<typename T>
    bool XLCellValueTemplate<T>::HasValueNode() const
    {
        if (ParentCell()->CellNode().child("v"))
            return true;
        else
            return false;
    }

    /**
     * @details Obtain a pointer to the corresponding XMLAttribute object, requesting attribute "t" from the cell node. If
     * no type attribute exists, a nullptr will be returned.
     * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
     * @post The current object, and any associated objects, are unchanged.
     */
    template<typename T>
    XMLAttribute XLCellValueTemplate<T>::TypeAttribute()
    {
        return ParentCell()->CellNode().attribute("t");
    }

    /**
     * @details Obtain a pointer to the corresponding XMLAttribute object, requesting attribute "t" from the cell node. If
     * no type attribute exists, a nullptr will be returned.
     * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
     * @post The current object, and any associated objects, are unchanged.
     */
    template<typename T>
    const XMLAttribute XLCellValueTemplate<T>::TypeAttribute() const
    {
        return ParentCell()->CellNode().attribute("t");
    }

    /**
     * @details Determines the existance of the type attribute, by checking if a nullptr is returned.
     * @pre The parent XLCell object is valid and has a corresponding node in the underlying XML file.
     * @post The current object, and any associated objects, are unchanged.
     */
    template<typename T>
    bool XLCellValueTemplate<T>::HasTypeAttribute() const
    {
        if (!ParentCell()->CellNode().attribute("t"))
            return false;
        else
            return true;
    }


    /**
     * @details This is a template assignment operator for all integer-type paremeters. The function calls the Set
     * template function and returns *this.
     * @pre
     * @post
     */
    template<typename T>
    template<typename U, typename std::enable_if<std::is_integral<U>::value, long long int>::type *>
    XLCellValueTemplate<T> &XLCellValueTemplate<T>::operator=(U numberValue)
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
    template<typename T>
    template<typename U, typename std::enable_if<std::is_floating_point<U>::value, long double>::type *>
    XLCellValueTemplate<T> &XLCellValueTemplate<T>::operator=(U numberValue)
    {
        Set(static_cast<long double>(numberValue));
        return *this;
    }

    /**
     * @details This is a template function for all integer-type parameters. If the current cell value is not already
     * of integer or floating point type (i.e. number type), set the m_value variable accordingly. Call the Set
     * method of the m_value member variable.
     * @pre
     * @post
     */
    template<typename T>
    template<typename U, typename std::enable_if<std::is_integral<U>::value, long long int>::type *>
    void XLCellValueTemplate<T>::Set(U numberValue)
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
    template<typename T>
    template<typename U, typename std::enable_if<std::is_floating_point<U>::value, long double>::type *>
    void XLCellValueTemplate<T>::Set(U numberValue)
    {
        if (ValueType() != XLValueType::Integer || ValueType() != XLValueType::Float)
            m_value = std::make_unique<XLValueNumber>(*this);

        static_cast<XLValueNumber &>(*m_value).Set(static_cast<long double>(numberValue));

        ParentCell()->SetModified();
    }
}

#endif //OPENXLEXE_XLCELLVALUETEMPLATE_H
