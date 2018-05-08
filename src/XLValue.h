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

#ifndef OPENXLEXE_XLCELLVALUETYPE_H
#define OPENXLEXE_XLCELLVALUETYPE_H

#include <string>
#include <memory>
#include "XLCellType.h"
#include "Utilities/XML/XML.h"

namespace OpenXLSX
{

    class XLCell;
    class XLCellValue;

//======================================================================================================================
//========== XLValue Class =============================================================================================
//======================================================================================================================

    /**
     * @brief The XLCellValueType class is a pure virtual class, functioning as parent class to any class representing
     * an Excel value type.
     */
    class XLValue
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor
         * @param parent A reference to the parent XLCellValue object.
         */
        explicit XLValue(XLCellValue &parent);

        /**
         * @brief Copy constructor
         * @param other The object to be copied.
         * @todo A copy constructor does not really make sense, as an XLCellValueType object will belong to a specific
         * XLCell object. However, a copy constructor is required in order to use the std::variant data structure. A
         * solution should be found, that does not require a copy constructor.
         */
        XLValue(const XLValue &other) = default;

        /**
         * @brief Move construcor
         * @param other The object to be move constructed.
         * @note This has been explicitly deleted.
         */
        XLValue(XLValue &&other) = delete;

        /**
         * @brief Destructor
         * @note Default implementation has been used.
         */
        virtual ~XLValue() = default;

        /**
         * @brief Copy assignment operator
         * @param other The object to be assigned from.
         * @return A reference to the current object with the new value-
         */
        XLValue &operator=(const XLValue &other);

        /**
         * @brief Move assignment operator.
         * @param other The object to be moved.
         * @return A reference to the object after the move.
         */
        XLValue &operator=(XLValue &&other) noexcept;

        /**
         * @brief Creates a polymorphic clone of the object
         * @param parent A reference to the parent XLCell object of the clone
         * @return A unique_ptr holding the clone.
         */
        virtual std::unique_ptr<XLValue> Clone(XLCell &parent) = 0;

        /**
         * @brief Get the value type of the object.
         * @return Returns an XLValueType object with the corresponding type.
         */
        virtual XLValueType ValueType() const = 0;

        /**
         * @brief Get the cell type of the object.
         * @return Return an XLCellType object with the corresponding type.
         */
        virtual XLCellType CellType() const = 0;

        /**
         * @brief Produce a string corrsponding to the type attribute in the underlying XML file.
         * @return A string with the type attribute to e used.
         */
        virtual std::string TypeString() const = 0;

        /**
         * @brief Produce a string with the current value, regardless of value type.
         * @return A string with the value as a string.
         */
        virtual std::string AsString() const = 0;

//----------------------------------------------------------------------------------------------------------------------
//           Protected Member Functions
//----------------------------------------------------------------------------------------------------------------------

    protected:

        /**
         * @brief Get the parent XLCellValue object.
         * @return A reference to the parent XLCellValue object.
         */
        XLCellValue *ParentCellValue();

        /**
         * @brief Get the parent XLCellValue object.
         * @return A const reference to the parent XLCellValue object.
         */
        const XLCellValue *ParentCellValue() const;

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:
        XLCellValue &m_parentCellValue; /**<  */

    };

}


#endif //OPENXLEXE_XLCELLVALUETYPE_H
