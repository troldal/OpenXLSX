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

#ifndef OPENXL_XLSTRING_H
#define OPENXL_XLSTRING_H

#include <string>
#include "XLCellType.h"
#include "XLValue.h"
#include "../@thirdparty/XML/XML.h"

namespace OpenXLSX
{
    class XLCell;

//======================================================================================================================
//========== XLValueString Class ============================================================================================
//======================================================================================================================

    /**
     * @brief This class represents the different types of string that are used in Excel. For the user, it behaves the
     * same, independent of string type.
     */
    class XLValueString: public XLValue
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor
         * @param parent A reference to the parent XLCellValue object.
         */
        explicit XLValueString(XLCellValue &parent);

        /**
         * @brief Constructor
         * @param stringValue A std::string with the string value to initialize the object with.
         * @param parent A reference to the parent XLCellValue object.
         */
        explicit XLValueString(std::string_view stringValue, XLCellValue &parent);

        /**
         * @brief Constructor
         * @param stringIndex The index of the relevant shared string.
         * @param parent A reference to the parent XLCellValue object.
         */
        explicit XLValueString(unsigned long stringIndex, XLCellValue &parent);

        /**
         * @brief Copy constructor
         * @param other The object to be copied.
         * @todo A copy constructor does not really make sense, as an XLCellValueType object will belong to a specific
         * XLCell object. However, a copy constructor is required in order to use the std::variant data structure. A
         * solution should be found, that does not require a copy constructor.
         */
        XLValueString(const XLValueString &other) = default;

        /**
         * @brief Move constructor
         * @param other The object to be moved.
         * @note The move constructor has been explicitly deleted.
         */
        XLValueString(XLValueString &&other) noexcept = default;

        /**
         * @brief Destructor
         * @note The default implementation has been used.
         */
        ~XLValueString() noexcept override = default;

        /**
         * @brief Copy assignment operator
         * @param other The object to be copied.
         * @return A reference to the object with the new value.
         */
        XLValueString &operator=(const XLValueString &other);

        /**
         * @brief Move assignment operator
         * @param other the object to be moved
         * @return A reference to the object with the new value.
         */
        XLValueString &operator=(XLValueString &&other) noexcept;

        /**
         * @brief Assignment operator
         * @param str A std::string to be assigned
         * @return A reference to the object with the new value.
         */
        XLValueString &operator=(const std::string &str);

        /**
         * @brief Set the string value
         * @param stringValue A std::string with the new value
         * @param type The type of the string; the default is an ordinary string.
         */
        void Set(std::string_view stringValue, XLStringType type = XLStringType::String);

        /**
         * @brief Get the string value as a std::string.
         * @return A std::string with the string value.
         */
        const std::string &String() const;

        /**
         * @brief Get the value as a std::string, regardless of the type. For all string types, it is identical
         * to the String method, and defined only to comply with the XLCellValueType interface.
         * @return A std::string with the string value.
         */
        std::string AsString() const override;

        /**
         * @brief Get the string type of the string.
         * @return An XLStringType with the type of the current string.
         */
        XLStringType StringType() const;

        /**
         * @brief Get the value type of the object.
         * @return An XLValueType object with the correct value type for the object.
         */
        XLValueType ValueType() const override;

        /**
         * @brief Get the cell type of the object.
         * @return An XLCellType object with the correct cell type for the object.
         */
        XLCellType CellType() const override;

        /**
         * @brief Get the type string for the object.
         * @return A std::string with a string corresponding to the type attribute for the object.
         */
        std::string TypeString() const override;

        /**
         * @brief Get the index of the string, if it is of type SharedString.
         * @return A long int with the index of the shared string.
         * @exception Throws an XLException if the type is not SharedString.
         */
        long SharedStringIndex() const;

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Functions
//----------------------------------------------------------------------------------------------------------------------

    private:

        /**
         * @brief Private function used to initialize a string object.
         * @param stringValue A std::string to initialize the object with.
         */
        void Initialize(std::string_view stringValue);



//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:
        XLStringType m_type; /**< The type of the string, i.e. String (ordinary), SharedString or InlineString */
        mutable std::string m_cache;

    };

}


#endif //OPENXLEXE_XLSTRING_H
