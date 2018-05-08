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

#ifndef OPENXLEXE_XLNUMBER_H
#define OPENXLEXE_XLNUMBER_H

#include <string>
#include "XLCellType.h"
#include "XLValue.h"
#include "Utilities/XML/XML.h"

namespace OpenXLSX
{
    class XLCell;

//======================================================================================================================
//========== XLValueNumber Class =======================================================================================
//======================================================================================================================
    /**
     * @brief
     */
    class XLValueNumber: public XLValue
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief
         * @param parent
         */
        explicit XLValueNumber(XLCellValue &parent);

        /**
         * @brief
         * @param numberValue
         * @param parent
         */
        XLValueNumber(long long int numberValue, XLCellValue &parent);

        /**
         * @brief
         * @param numberValue
         * @param parent
         */
        XLValueNumber(long double numberValue, XLCellValue &parent);

        /**
         * @brief
         * @param other
         * @todo A copy constructor does not really make sense, as an XLCellValueType object will belong to a specific
         * XLCell object. However, a copy constructor is required in order to use the std::variant data structure. A
         * solution should be found, that does not require a copy constructor.
         */
        XLValueNumber(const XLValueNumber &other) = default;

        /**
         * @brief
         * @param other
         * @todo A copy constructor does not really make sense, as an XLCellValueType object will belong to a specific
         * XLCell object. However, a copy constructor is required in order to use the std::variant data structure. A
         * solution should be found, that does not require a copy constructor.
         */
        XLValueNumber(XLValueNumber &&other) = default;

        /**
         * @brief
         */
        ~XLValueNumber() override = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLValueNumber &operator=(const XLValueNumber &other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLValueNumber &operator=(XLValueNumber &&other) noexcept;

        /**
         * @brief
         * @param numberValue
         * @return
         */
        XLValueNumber &operator=(long long int numberValue);

        /**
         * @brief
         * @param numberValue
         * @return
         */
        XLValueNumber &operator=(long double numberValue);

        /**
         * @brief
         * @param parent
         * @return
         */
        std::unique_ptr<XLValue> Clone(XLCell &parent) override;

        /**
         * @brief
         * @param numberValue
         */
        void Set(long long int numberValue);

        /**
         * @brief
         * @param numberValue
         */
        void Set(long double numberValue);

        /**
         * @brief
         * @return
         */
        long long int Integer() const;

        /**
         * @brief
         * @return
         */
        long double Float() const;

        /**
         * @brief
         * @return
         */
        long long int AsInteger() const;

        /**
         * @brief
         * @return
         */
        long double AsFloat() const;

        /**
         * @brief
         * @return
         */
        std::string AsString() const override;

        /**
         * @brief
         * @return
         */
        XLNumberType NumberType() const;

        /**
         * @brief
         * @return
         */
        XLValueType ValueType() const override;

        /**
         * @brief
         * @return
         */
        XLCellType CellType() const override;

        /**
         * @brief
         * @return
         */
        std::string TypeString() const override;

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Functions
//----------------------------------------------------------------------------------------------------------------------

    private:

        /**
         * @brief
         * @param numberString
         * @return
         */
        XLNumberType DetermineNumberType(const std::string &numberString) const;

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:
        XLNumberType m_numberType; /**<  */
        long long int m_integer;
        long double m_float;
    };

}

#endif //OPENXLEXE_XLNUMBER_H
