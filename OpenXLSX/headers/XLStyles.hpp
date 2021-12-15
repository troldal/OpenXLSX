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

#ifndef OPENXLSX_XLSYLES_HPP
#define OPENXLSX_XLSYLES_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

#include <string>
#include <vector>
#include <map>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlFile.hpp"

namespace OpenXLSX
{
    class XLStyle;

    class OPENXLSX_EXPORT XLStyles : public XLXmlFile
    {
        friend class XLCell;
        friend XLDocument;

    public:
        XLStyles() = default;
        ~XLStyles();
        XLStyles(XLXmlData* xmlData);
        XLStyle style(const XLCell& cell) const;
        std::string formatString(int numFmtId) const;

    private:
        void init(const XLXmlData* stylesData);

    private:
        std::vector<XLStyle> m_VecStyle;
    };

    class OPENXLSX_EXPORT XLNumberFormat
    {
    public:
        enum class XLNumberType { kUnkown, kDate, kPercent, kCurrency };

    public:
        explicit XLNumberFormat(const XLStyle& style);
        ~XLNumberFormat() = default;

        XLNumberType type();
        std::string currencySumbol() const;

    private:
        XLNumberType tryFindType();
        XLNumberType tryBuiltinType();

    private:
        std::reference_wrapper<const XLStyle> m_style;
        std::string                           m_currencySumbol;
        std::string                           m_fmtLocal;
    };

    class OPENXLSX_EXPORT XLStyle
    {
        friend class XLCell;
        friend class XLStyles;

    public:
        explicit XLStyle(const XLDocument& styles);
        ~XLStyle() = default;

        XLNumberFormat numberFormat() const;
        std::string    formatString() const;
        int            numFmtId() const;

    private:
        int                                      m_numFmtId = -1;
        int                                      m_fontId   = -1;
        int                                      m_fillId   = -1;
        int                                      m_borderId = -1;
        int                                      m_xfId     = -1;
        std::reference_wrapper<const XLDocument> m_doc;
    };

    
}// namespace OpenXLSX



#pragma warning(pop)
#endif    // OPENXLSX_XLSYLES_HPP