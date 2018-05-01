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

#ifndef OPENXLEXE_XLEMPTY_H
#define OPENXLEXE_XLEMPTY_H

#include "XLValue.h"

namespace OpenXLSX
{

//======================================================================================================================
//========== XLValueEmpty Class =============================================================================================
//======================================================================================================================

    /**
     * @brief This class encapsulates an 'empty' value, i.e. a cell without a value or type.
     */
    class XLValueEmpty: public XLValue
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor
         * @param parent A reference to the parent XLCellValue object.
         */
        explicit XLValueEmpty(XLCellValue &parent);

        /**
         * @brief Copy constructor
         * @param other The XLEmpty object to be copied.
         * @note The default implementation has been used.
         */
        XLValueEmpty(const XLValueEmpty &other) = default;

        /**
         * @brief Move constructor
         * @param other the XLEmpty object to move.
         * @note The move constructor has been explicitly deleted.
         */
        XLValueEmpty(XLValueEmpty &&other) = delete;

        /**
         * @brief Destructor
         * @note The default implementation has been used.
         */
        ~XLValueEmpty() override = default;

        /**
         * @brief Copy assignment operator.
         * @param other The XLEmpty object to copy.
         * @return A reference to the object with the new value.
         * @note The default implementation has been used.
         */
        XLValueEmpty &operator=(const XLValueEmpty &other) = default;

        /**
         * @brief Move assignment operator.
         * @param other The XLEmpty object to move.
         * @return A reference to the object with the moved value.
         * @note The default implementation has been used.
         */
        XLValueEmpty &operator=(XLValueEmpty &&other) noexcept = default;

        /**
         * @brief Creates a polymorphic clone of the object.
         * @param parent The parent XLCell object of the clone.
         * @return A unique_ptr with the clone.
         */
        std::unique_ptr<XLValue> Clone(XLCell &parent) override;

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
         * @brief Get the value as a std::string, regardless of the type.
         * @return A std::string with the string value.
         */
        std::string AsString() const override;

    };

}


#endif //OPENXLEXE_XLEMPTY_H
