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

#ifndef OPENXLEXE_XLBOOLEAN_H
#define OPENXLEXE_XLBOOLEAN_H

#include "XLValue.h"
#include "XLCellType.h"

namespace OpenXLSX
{

//======================================================================================================================
//========== XLValueString Class ============================================================================================
//======================================================================================================================

    /**
     * @brief The XLBoolean class encapsulates the concept of a boolean value. The reason to use this rather than the
     * built-in type bool, is to enable polymorphic exchange of types and to avoid implicit conversion between e.g. int
     * and bool.
     */
    class XLValueBoolean: public XLValue
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor
         * @param parent A reference to the parent XLCellValue object.
         */
        explicit XLValueBoolean();

        /**
         * @brief Copy constructor
         * @param other The object to be copied.
         * @note This uses the default implementation.
         */
        XLValueBoolean(const XLValueBoolean &other) = default;

        /**
         * @brief Move constructor
         * @param other The object to be moved.
         * @note This uses the default implementation.
         */
        XLValueBoolean(XLValueBoolean &&other) noexcept = default;

        /**
         * @brief Destructor
         * @note This uses the default implementation.
         */
        ~XLValueBoolean() noexcept override = default;

        /**
         * @brief Copy assignment operator.
         * @param other The object to be assigned from.
         * @return A reference to this object, with the new value.
         * @note This uses the default implementation.
         */
        XLValueBoolean &operator=(const XLValueBoolean &other) = default;

        /**
         * @brief Move assignment operator.
         * @param other The object to be moved.
         * @return A reference to this object, with the moved value.
         * @note This uses the default implementation.
         * @todo Is this the right way to do it?
         */
        XLValueBoolean &operator=(XLValueBoolean &&other) noexcept = default;

        /**
         * @brief Set the bool value
         * @param boolValue The boolean value to the the object to.
         */
        const std::string& Set(bool boolValue);

        /**
         * @brief Get the value of the boolean.
         * @return An XLBool object with the value.
         */
        bool Get() const;

        /**
         * @brief Get the value as a string.
         * @return A std::string with "True" or "False", depending on the value.
         */
        std::string AsString() const override;

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

    };

}


#endif //OPENXLEXE_XLBOOLEAN_H
