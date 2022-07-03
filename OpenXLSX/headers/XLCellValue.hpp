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
#include <cmath>
#include <cstdint>
#include <iostream>
#include <string>
#include <variant>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLDateTime.hpp"
#include "XLException.hpp"
#include "XLXmlParser.hpp"

// ========== CLASS AND ENUM TYPE DEFINITIONS ========== //
namespace OpenXLSX
{
    //---------- Forward Declarations ----------//
    class XLCellValueProxy;
    class XLCell;

    /**
     * @brief Enum defining the valid value types for a an Excel spreadsheet cell.
     */
    enum class XLValueType { Empty, Boolean, Integer, Float, Error, String };

    /**
     * @brief Class encapsulating a cell value.
     */
    class OPENXLSX_EXPORT XLCellValue
    {
        //---------- Friend Declarations ----------//

        // TODO: Consider template functions to compare to ints, floats etc.
        friend bool          operator==(const XLCellValue& lhs, const XLCellValue& rhs);
        friend bool          operator!=(const XLCellValue& lhs, const XLCellValue& rhs);
        friend bool          operator<(const XLCellValue& lhs, const XLCellValue& rhs);
        friend bool          operator>(const XLCellValue& lhs, const XLCellValue& rhs);
        friend bool          operator<=(const XLCellValue& lhs, const XLCellValue& rhs);
        friend bool          operator>=(const XLCellValue& lhs, const XLCellValue& rhs);
        friend std::ostream& operator<<(std::ostream& os, const XLCellValue& value);
        friend std::hash<OpenXLSX::XLCellValue>;

    public:
        //---------- Public Member Functions ----------//

        /**
         * @brief Default constructor
         */
        XLCellValue();

        /**
         * @brief A templated constructor. Any value convertible to a valid cell value can be used as argument.
         * @tparam T The type of the argument (will be automatically deduced).
         * @param value The value.
         * @todo Consider changing the enable_if statement to check for objects with a .c_str() member function.
         */
        template<
            typename T,
            typename std::enable_if<std::is_integral_v<T> || std::is_floating_point_v<T> || std::is_same_v<std::decay_t<T>, std::string> ||
                                    std::is_same_v<std::decay_t<T>, std::string_view> || std::is_same_v<std::decay_t<T>, const char*> ||
                                    std::is_same_v<std::decay_t<T>, char*> || std::is_same_v<T, XLDateTime>>::type* = nullptr>
        XLCellValue(T value)    // NOLINT
        {
            // ===== If the argument is a bool, set the m_type attribute to Boolean.
            if constexpr (std::is_integral_v<T> && std::is_same_v<T, bool>) {
                m_type  = XLValueType::Boolean;
                m_value = value;
            }

            // ===== If the argument is an integral type, set the m_type attribute to Integer.
            else if constexpr (std::is_integral_v<T> && !std::is_same_v<T, bool>) {
                m_type  = XLValueType::Integer;
                m_value = int64_t(value);
            }

            // ===== If the argument is a string type (i.e. is constructable from *char),
            // ===== set the m_type attribute to String.
            else if constexpr (std::is_same_v<std::decay_t<T>, std::string> || std::is_same_v<std::decay_t<T>, std::string_view> ||
                               std::is_same_v<std::decay_t<T>, const char*> ||
                               (std::is_same_v<std::decay_t<T>, char*> && !std::is_same_v<T, bool>))
            {
                m_type = XLValueType::String;
                m_value = std::string(value);
            }

            // ===== If the argument is an XLDateTime, set the value to the date/time serial number.
            else if constexpr (std::is_same_v<T, XLDateTime>) {
                m_type  = XLValueType::Float;
                m_value = value.serial();
            }

            // ===== If the argument is a floating point type, set the m_type attribute to Float.
            // ===== If not, a static_assert will result in compilation error.
            else {
                static_assert(std::is_floating_point_v<T>, "Invalid argument for constructing XLCellValue object");
                if (std::isfinite(value)) {
                    m_type  = XLValueType::Float;
                    m_value = double(value);
                }
                else {
                    m_type = XLValueType::Error;
                    m_value = std::string("#NUM!");
                }
            }
        }

        /**
         * @brief Copy constructor.
         * @param other The object to be copied.
         */
        XLCellValue(const XLCellValue& other);

        /**
         * @brief Move constructor.
         * @param other The object to be moved.
         */
        XLCellValue(XLCellValue&& other) noexcept;

        /**
         * @brief Destructor
         */
        ~XLCellValue();

        /**
         * @brief Copy assignment operator.
         * @param other Object to be copied.
         * @return Reference to the copied-to object.
         */
        XLCellValue& operator=(const XLCellValue& other);

        /**
         * @brief Move assignment operator.
         * @param other Object to be moved.
         * @return Reference to the moved-to object.
         */
        XLCellValue& operator=(XLCellValue&& other) noexcept;

        /**
         * @brief Templated copy assignment operator.
         * @tparam T The type of the value argument.
         * @param value The value.
         * @return A reference to the assigned-to object.
         */
        template<
            typename T,
            typename std::enable_if<std::is_integral_v<T> || std::is_floating_point_v<T> || std::is_same_v<std::decay_t<T>, std::string> ||
                                    std::is_same_v<std::decay_t<T>, std::string_view> || std::is_same_v<std::decay_t<T>, const char*> ||
                                    std::is_same_v<std::decay_t<T>, char*> || std::is_same_v<T, XLDateTime>>::type* = nullptr>
        XLCellValue& operator=(T value)
        {
            // ===== Implemented using copy-and-swap.
            XLCellValue temp(value);
            std::swap(*this, temp);
            return *this;
        }

        /**
         * @brief Templated setter for integral and bool types.
         * @tparam T The type of the value argument.
         * @param numberValue The value
         */
        template<
            typename T,
            typename std::enable_if<std::is_same_v<T, XLCellValue> || std::is_integral_v<T> || std::is_floating_point_v<T> ||
                                    std::is_same_v<std::decay_t<T>, std::string> || std::is_same_v<std::decay_t<T>, std::string_view> ||
                                    std::is_same_v<std::decay_t<T>, const char*> || std::is_same_v<std::decay_t<T>, char*> ||
                                    std::is_same_v<T, XLDateTime>>::type* = nullptr>
        void set(T numberValue)
        {
            // ===== Implemented using the assignment operator.
            *this = numberValue;
        }

        /**
         * @brief Templated getter.
         * @tparam T The type of the value to be returned.
         * @return The value as a type T object.
         * @throws XLValueTypeError if the XLCellValue object does not contain a compatible type.
         */
        template<
            typename T,
            typename std::enable_if<std::is_integral_v<T> || std::is_floating_point_v<T> || std::is_same_v<std::decay_t<T>, std::string> ||
                                    std::is_same_v<std::decay_t<T>, std::string_view> || std::is_same_v<std::decay_t<T>, const char*> ||
                                    std::is_same_v<std::decay_t<T>, char*> || std::is_same_v<T, XLDateTime>>::type* = nullptr>
        T get() const
        {
            try {
                if constexpr (std::is_integral_v<T> && std::is_same_v<T, bool>) return std::get<bool>(m_value);

                if constexpr (std::is_integral_v<T> && !std::is_same_v<T, bool>) return static_cast<T>(std::get<int64_t>(m_value));

                if constexpr (std::is_floating_point_v<T>) {
                    if (m_type == XLValueType::Error) return std::nan("1");
                    return static_cast<T>(std::get<double>(m_value));
                }

                if constexpr (std::is_same_v<std::decay_t<T>, std::string> || std::is_same_v<std::decay_t<T>, std::string_view> ||
                              std::is_same_v<std::decay_t<T>, const char*> ||
                              (std::is_same_v<std::decay_t<T>, char*> && !std::is_same_v<T, bool>))
                    return std::get<std::string>(m_value).c_str();

                if constexpr (std::is_same_v<T, XLDateTime>) return XLDateTime(std::get<double>(m_value));
            }

            catch (const std::bad_variant_access& ) {
                throw XLValueTypeError("XLCellValue object does not contain the requested type.");
            }
        }

        /**
         * @brief Explicit conversion operator for easy conversion to supported types.
         * @tparam T The type to cast to.
         * @return The XLCellValue object cast to requested type.
         * @throws XLValueTypeError if the XLCellValue object does not contain a compatible type.
         */
        template<
            typename T,
            typename std::enable_if<std::is_integral_v<T> || std::is_floating_point_v<T> || std::is_same_v<std::decay_t<T>, std::string> ||
                                    std::is_same_v<std::decay_t<T>, std::string_view> || std::is_same_v<std::decay_t<T>, const char*> ||
                                    std::is_same_v<std::decay_t<T>, char*> || std::is_same_v<T, XLDateTime>>::type* = nullptr>
        operator T() const
        {
            return this->get<T>();
        }

        /**
         * @brief Clears the contents of the XLCellValue object.
         * @return Returns a reference to the current object.
         */
        XLCellValue& clear();

        /**
         * @brief Sets the value type to XLValueType::Error.
         * @return Returns a reference to the current object.
         */
        XLCellValue& setError(const std::string &error);

        /**
         * @brief Get the value type of the current object.
         * @return An XLValueType for the current object.
         */
        XLValueType type() const;

        /**
         * @brief Get the value type of the current object, as a string representation
         * @return A std::string representation of the value type.
         */
        std::string typeAsString() const;

    private:
        //---------- Private Member Variables ---------- //

        std::variant<std::string, int64_t, double, bool> m_value { std::string("") };                /**< The value contained in the cell. */
        XLValueType                                      m_type { XLValueType::Empty }; /**< The value type of the cell. */
    };

    /**
     * @brief The XLCellValueProxy class is used for proxy (or placeholder) objects for XLCellValue objects.
     * @details The XLCellValueProxy class is used for proxy (or placeholder) objects for XLCellValue objects.
     * The purpose is to enable implicit conversion during assignment operations. XLCellValueProxy objects
     * can not be constructed manually by the user, only through XLCell objects.
     */
    class OPENXLSX_EXPORT XLCellValueProxy
    {
        friend class XLCell;
        friend class XLCellValue;

    public:
        //---------- Public Member Functions ----------//

        /**
         * @brief Destructor
         */
        ~XLCellValueProxy();

        /**
         * @brief Copy assignment operator.
         * @param other XLCellValueProxy object to be copied.
         * @return A reference to the current object.
         */
        XLCellValueProxy& operator=(const XLCellValueProxy& other);

        /**
         * @brief Templated assignment operator
         * @tparam T The type of numberValue assigned to the object.
         * @param value The value.
         * @return A reference to the current object.
         */
        template<
            typename T,
            typename std::enable_if<std::is_integral_v<T> || std::is_floating_point_v<T> || std::is_same_v<std::decay_t<T>, std::string> ||
                                    std::is_same_v<std::decay_t<T>, std::string_view> || std::is_same_v<std::decay_t<T>, const char*> ||
                                    std::is_same_v<std::decay_t<T>, char*> || std::is_same_v<T, XLCellValue> ||
                                    std::is_same_v<T, XLDateTime>>::type* = nullptr>
        XLCellValueProxy& operator=(T value)
        {    // NOLINT

            if constexpr (std::is_integral_v<T> && std::is_same_v<T, bool>)    // if bool
                setBoolean(value);

            else if constexpr (std::is_integral_v<T> && !std::is_same_v<T, bool>)    // if integer
                setInteger(value);

            else if constexpr (std::is_floating_point_v<T>)    // if floating point
                setFloat(value);

            else if constexpr (std::is_same_v<T, XLDateTime>)
                setFloat(value.serial());

            else if constexpr (std::is_same_v<std::decay_t<T>, std::string> || std::is_same_v<std::decay_t<T>, std::string_view> ||
                               std::is_same_v<std::decay_t<T>, const char*> ||
                               (std::is_same_v<std::decay_t<T>, char*> && !std::is_same_v<T, bool> && !std::is_same_v<T, XLCellValue>))
            {
                if constexpr (std::is_same_v<std::decay_t<T>, const char*> || std::is_same_v<std::decay_t<T>, char*>)
                    setString(value);
                else if constexpr (std::is_same_v<std::decay_t<T>, std::string_view>)
                    setString(std::string(value).c_str());
                else
                    setString(value.c_str());
            }

            if constexpr (std::is_same_v<T, XLCellValue>) {
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
                        setString(value.template get<const char*>());
                        break;
                    case XLValueType::Empty:
                        clear();
                        break;
                    default:
                        setError("#N/A");
                        break;
                }
            }

            return *this;
        }

        /**
         * @brief
         * @tparam T
         * @param value
         */
        template<
            typename T,
            typename std::enable_if<std::is_integral_v<T> || std::is_floating_point_v<T> || std::is_same_v<std::decay_t<T>, std::string> ||
                                    std::is_same_v<std::decay_t<T>, std::string_view> || std::is_same_v<std::decay_t<T>, const char*> ||
                                    std::is_same_v<std::decay_t<T>, char*> || std::is_same_v<T, XLCellValue> ||
                                    std::is_same_v<T, XLDateTime>>::type* = nullptr>
        void set(T value)
        {
            *this = value;
        }

        /**
         * @brief
         * @tparam T
         * @return
         * @todo Is an explicit conversion operator needed as well?
         */
        template<
            typename T,
            typename std::enable_if<std::is_integral_v<T> || std::is_floating_point_v<T> || std::is_same_v<std::decay_t<T>, std::string> ||
                                    std::is_same_v<std::decay_t<T>, std::string_view> || std::is_same_v<std::decay_t<T>, const char*> ||
                                    std::is_same_v<std::decay_t<T>, char*> || std::is_same_v<T, XLDateTime>>::type* = nullptr>
        T get() const
        {
            return getValue().get<T>();
        }

        /**
         * @brief Clear the contents of the cell.
         * @return A reference to the current object.
         */
        XLCellValueProxy& clear();

        /**
         * @brief Set the cell value to a error state.
         * @return A reference to the current object.
         */
        XLCellValueProxy& setError(const std::string & error);

        /**
         * @brief Get the value type for the cell.
         * @return An XLCellValue corresponding to the cell value.
         */
        XLValueType type() const;

        /**
         * @brief Get the value type of the current object, as a string representation
         * @return A std::string representation of the value type.
         */
        std::string typeAsString() const;

        /**
         * @brief Implicitly convert the XLCellValueProxy object to a XLCellValue object.
         * @return An XLCellValue object, corresponding to the cell value.
         */
        operator XLCellValue();    // NOLINT

        /**
         * @brief
         * @tparam T
         * @return
         */
        template<
            typename T,
            typename std::enable_if<std::is_integral_v<T> || std::is_floating_point_v<T> || std::is_same_v<std::decay_t<T>, std::string> ||
                                    std::is_same_v<std::decay_t<T>, std::string_view> || std::is_same_v<std::decay_t<T>, const char*> ||
                                    std::is_same_v<std::decay_t<T>, char*> || std::is_same_v<T, XLDateTime>>::type* = nullptr>
        operator T() const
        {
            return getValue().get<T>();
        }

    private:
        //---------- Private Member Functions ---------- //

        /**
         * @brief Constructor
         * @param cell Pointer to the parent XLCell object.
         * @param cellNode Pointer to the corresponding XMLNode object.
         */
        XLCellValueProxy(XLCell* cell, XMLNode* cellNode);

        /**
         * @brief Copy constructor
         * @param other Object to be copied.
         */
        XLCellValueProxy(const XLCellValueProxy& other);

        /**
         * @brief Move constructor
         * @param other Object to be moved.
         */
        XLCellValueProxy(XLCellValueProxy&& other) noexcept;

        /**
         * @brief Move assignment operator
         * @param other Object to be moved
         * @return Reference to moved-to pbject.
         */
        XLCellValueProxy& operator=(XLCellValueProxy&& other) noexcept;

        /**
         * @brief Set cell to an integer value.
         * @param numberValue The value to be set.
         */
        void setInteger(int64_t numberValue);

        /**
         * @brief Set the cell to a bool value.
         * @param numberValue The value to be set.
         */
        void setBoolean(bool numberValue);

        /**
         * @brief Set the cell to a floating point value.
         * @param numberValue The value to be set.
         */
        void setFloat(double numberValue);

        /**
         * @brief Set the cell to a string value.
         * @param stringValue The value to be set.
         */
        void setString(const char* stringValue);

        /**
         * @brief Get a copy of the XLCellValue object for the cell.
         * @return An XLCellValue object.
         */
        XLCellValue getValue() const;

        //---------- Private Member Variables ---------- //

        XLCell*  m_cell;     /**< Pointer to the owning XLCell object. */
        XMLNode* m_cellNode; /**< Pointer to corresponding XML cell node. */
    };

}    // namespace OpenXLSX

// TODO: Consider comparison operators on fundamental datatypes
// ========== FRIEND FUNCTION IMPLEMENTATIONS ========== //
namespace OpenXLSX
{
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

    /**
     * @brief
     * @param os
     * @param value
     * @return
     */
    inline std::ostream& operator<<(std::ostream& os, const XLCellValue& value)
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

    inline std::ostream& operator<<(std::ostream& os, const XLCellValueProxy& value)
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

namespace std
{
    template<>
    struct hash<OpenXLSX::XLCellValue> // NOLINT
    {
        std::size_t operator()(const OpenXLSX::XLCellValue& value) const noexcept
        {
            return std::hash<std::variant<std::string, int64_t, double, bool>> {}(value.m_value);
        }
    };
}    // namespace std

#pragma warning(pop)
#endif    // OPENXLSX_XLCELLVALUE_HPP
