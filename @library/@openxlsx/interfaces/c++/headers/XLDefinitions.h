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

#ifndef OPENXLSX_XLDEFINITIONS_H
#define OPENXLSX_XLDEFINITIONS_H

#include <cstdint>
#include "config.h"

namespace OpenXLSX {

    //======================================================================================================================
    //========== Global Constants ==========================================================================================
    //======================================================================================================================

    constexpr uint32_t maxRows = 1'048'576;

    constexpr uint16_t maxCols = 16'384;

    constexpr uint8_t alphabetSize = 26;

    constexpr uint8_t asciiOffset = 64;

    //======================================================================================================================
    //========== XLDocumentProperties Enum =================================================================================
    //======================================================================================================================

    /**
     * @brief The XLDocumentProperties class is an enumeration of the possible properties (metadata) that can be set
     * for a XLDocument object (and .xlsx file)
     */
    enum class XLProperty {
        Title,
        Subject,
        Creator,
        Keywords,
        Description,
        LastModifiedBy,
        LastPrinted,
        CreationDate,
        ModificationDate,
        Category,
        Application,
        DocSecurity,
        ScaleCrop,
        Manager,
        Company,
        LinksUpToDate,
        SharedDoc,
        HyperlinkBase,
        HyperlinksChanged,
        AppVersion
    };

    //======================================================================================================================
    //========== XLSheetType Enum ==========================================================================================
    //======================================================================================================================

    /**
     * @brief The XLSheetType class is an enumeration of the available sheet types, e.g. Worksheet (ordinary
     * spreadsheets), and Chartsheet (sheets with only a chart).
     */
    enum class XLSheetType {
        WorkSheet, ChartSheet, DialogSheet, MacroSheet
    };


    //======================================================================================================================
    //========== XLSheetState Enum =========================================================================================
    //======================================================================================================================

    /**
     * @brief The XLSheetState is an enumeration of the possible (visibility) states, e.g. Visible or Hidden.
     */
    enum class XLSheetState {
        Visible, Hidden, VeryHidden
    };


    //======================================================================================================================
    //========== XLValue Enum ==========================================================================================
    //======================================================================================================================

    /**
     * @brief The XLValueType class is an enumeration of the possible cell value types.
     */
    enum class XLValueType {
        Empty, Boolean, Integer, Float, Error, String
    };

}  // namespace OpenXLSX

#endif //OPENXLSX_XLDEFINITIONS_H
