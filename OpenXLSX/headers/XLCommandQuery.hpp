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

#ifndef OPENXLSX_XLCOMMANDQUERY_HPP
#define OPENXLSX_XLCOMMANDQUERY_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <map>
#include <string>
#include <variant>

// ===== OpenXLSX Includes ===== //
#include "XLXmlData.hpp"

namespace OpenXLSX
{
    class XLSharedStrings;

    class XLCommandSetSheetName
    {
    public:
        XLCommandSetSheetName(const std::string& sheetID, const std::string& sheetName, const std::string& newName)
            : m_sheetID(sheetID),
              m_sheetName(sheetName),
              m_newName(newName)
        {}

        const std::string& sheetID() const
        {
            return m_sheetID;
        }

        const std::string& sheetName() const
        {
            return m_sheetName;
        }

        const std::string& newName() const
        {
            return m_newName;
        }

    private:
        std::string m_sheetID {};
        std::string m_sheetName {};
        std::string m_newName {};
    };

    class XLCommandSetSheetVisibility
    {
    public:
        XLCommandSetSheetVisibility(const std::string& sheetID, const std::string& sheetName, const std::string& sheetVisibility)
            : m_sheetID(sheetID),
              m_sheetName(sheetName),
              m_sheetVisibility(sheetVisibility)
        {}

        const std::string& sheetID() const
        {
            return m_sheetID;
        }
        const std::string& sheetName() const
        {
            return m_sheetName;
        }
        const std::string& sheetVisibility() const
        {
            return m_sheetVisibility;
        }

    private:
        std::string m_sheetID {};
        std::string m_sheetName {};
        std::string m_sheetVisibility {};
    };

    class XLCommandSetSheetColor
    {
    public:
        XLCommandSetSheetColor(const std::string& sheetID, const std::string& sheetName, const std::string& sheetColor)
            : m_sheetID(sheetID),
              m_sheetName(sheetName),
              m_sheetColor(sheetColor)
        {}

        const std::string& sheetID() const
        {
            return m_sheetID;
        }
        const std::string& sheetName() const
        {
            return m_sheetName;
        }
        const std::string& sheetColor() const
        {
            return m_sheetColor;
        }

    private:
        std::string m_sheetID {};
        std::string m_sheetName {};
        std::string m_sheetColor {};
    };

    class XLCommandSetSheetIndex
    {
    public:
        XLCommandSetSheetIndex(const std::string& sheetID, uint16_t sheetIndex) : m_sheetID(sheetID), m_sheetIndex(sheetIndex) {}

        const std::string& sheetID() const
        {
            return m_sheetID;
        }
        uint16_t sheetIndex() const
        {
            return m_sheetIndex;
        }

    private:
        std::string m_sheetID {};
        uint16_t    m_sheetIndex {};
    };

    class XLCommandResetCalcChain
    {
    };

    class XLCommandAddSharedStrings
    {
    };

    class XLCommandAddWorksheet
    {
    public:
        XLCommandAddWorksheet(const std::string& sheetName, const std::string& sheetPath) : m_sheetName(sheetName), m_sheetPath(sheetPath)
        {}

        const std::string& sheetName() const
        {
            return m_sheetName;
        }

        const std::string& sheetPath() const
        {
            return m_sheetPath;
        }

    private:
        std::string m_sheetName {};
        std::string m_sheetPath {};
    };

    class XLCommandAddChartsheet
    {
    public:
        XLCommandAddChartsheet(const std::string& sheetName, uint16_t sheetIndex) : m_sheetName(sheetName), m_sheetIndex(sheetIndex) {}

        const std::string& sheetName() const
        {
            return m_sheetName;
        }
        uint16_t sheetIndex() const
        {
            return m_sheetIndex;
        }

    private:
        std::string m_sheetName {};
        uint16_t    m_sheetIndex {};
    };

    class XLCommandDeleteSheet
    {
    public:
        XLCommandDeleteSheet(const std::string& sheetID, const std::string& sheetName) : m_sheetID(sheetID), m_sheetName(sheetName) {}

        const std::string& sheetID() const
        {
            return m_sheetID;
        }
        const std::string& sheetName() const
        {
            return m_sheetName;
        }

    private:
        std::string m_sheetID {};
        std::string m_sheetName {};
    };

    class XLCommandCloneSheet
    {
    public:
        XLCommandCloneSheet(const std::string& sheetID, const std::string& cloneName) : m_sheetID(sheetID), m_cloneName(cloneName) {}

        const std::string& sheetID() const
        {
            return m_sheetID;
        }

        const std::string& cloneName() const
        {
            return m_cloneName;
        }

    private:
        std::string m_sheetID {};   /**< ID of the sheet to clone. */
        std::string m_cloneName {}; /**< Name of the cloned sheet. */
    };

    class XLQuerySheetName
    {
    public:
        explicit XLQuerySheetName(const std::string& sheetID) : m_sheetID(sheetID) {}

        const std::string& sheetID() const
        {
            return m_sheetID;
        }

        const std::string& sheetName() const
        {
            return m_sheetName;
        }

        void setSheetName(const std::string& sheetName)
        {
            m_sheetName = sheetName;
        }

    private:
        std::string m_sheetID {};
        std::string m_sheetName {};
    };

    class XLQuerySheetIndex
    {
    public:
        explicit XLQuerySheetIndex(const std::string& sheetID) : m_sheetID(sheetID) {}

        const std::string& sheetID() const
        {
            return m_sheetID;
        }

        uint16_t sheetIndex() const
        {
            return m_sheetIndex;
        }

        void setSheetIndex(uint16_t sheetIndex)
        {
            m_sheetIndex = sheetIndex;
        }

    private:
        std::string m_sheetID {};
        uint16_t    m_sheetIndex {};
    };

    class XLQuerySheetVisibility
    {
    public:
        explicit XLQuerySheetVisibility(const std::string& sheetID) : m_sheetID(sheetID) {}

        const std::string& sheetID() const
        {
            return m_sheetID;
        }

        const std::string& sheetVisibility() const
        {
            return m_sheetVisibility;
        }

        void setSheetVisibility(const std::string& sheetVisibility)
        {
            m_sheetVisibility = sheetVisibility;
        }

    private:
        std::string m_sheetID {};
        std::string m_sheetVisibility {};
    };

    class XLQuerySheetType
    {
    public:
        explicit XLQuerySheetType(const std::string& sheetID) : m_sheetID(sheetID) {}

        const std::string& sheetID() const
        {
            return m_sheetID;
        }

        XLContentType sheetType() const
        {
            return m_sheetType;
        }

        void setSheetType(XLContentType sheetType)
        {
            m_sheetType = sheetType;
        }

    private:
        std::string   m_sheetID {};
        XLContentType m_sheetType {};
    };

    class XLQuerySheetID
    {
    public:
        explicit XLQuerySheetID(const std::string& sheetName) : m_sheetName(sheetName) {}

        const std::string& sheetName() const
        {
            return m_sheetName;
        }

        const std::string& sheetID() const
        {
            return m_sheetID;
        }

        void setSheetID(const std::string& sheetID)
        {
            m_sheetID = sheetID;
        }

    private:
        std::string m_sheetName {};
        std::string m_sheetID {};
    };

    class XLQuerySheetRelsID
    {
    public:
        explicit XLQuerySheetRelsID(const std::string& sheetPath) : m_sheetPath(sheetPath) {}

        const std::string& sheetPath() const
        {
            return m_sheetPath;
        }

        const std::string& sheetID() const
        {
            return m_sheetID;
        }

        void setSheetID(const std::string& sheetID)
        {
            m_sheetID = sheetID;
        }

    private:
        std::string m_sheetPath {};
        std::string m_sheetID {};
    };

    class XLQuerySheetRelsTarget
    {
    public:
        explicit XLQuerySheetRelsTarget(const std::string& sheetID) : m_sheetID(sheetID) {}

        const std::string& sheetID() const
        {
            return m_sheetID;
        }

        const std::string& sheetTarget() const
        {
            return m_sheetTarget;
        }

        void setSheetTarget(const std::string& sheetTarget)
        {
            m_sheetTarget = sheetTarget;
        }

    private:
        std::string m_sheetID {};
        std::string m_sheetTarget {};
    };

    class XLQuerySharedStrings
    {
    public:
        explicit XLQuerySharedStrings() = default;

        XLSharedStrings* sharedStrings() const
        {
            return m_sharedStrings;
        }

        void setSharedStrings(XLSharedStrings* sharedStrings)
        {
            m_sharedStrings = sharedStrings;
        }

    private:
        XLSharedStrings* m_sharedStrings { nullptr };
    };

    class XLQueryXmlData
    {
    public:
        explicit XLQueryXmlData(const std::string& xmlPath) : m_xmlPath(xmlPath) {}

        const std::string& xmlPath() const
        {
            return m_xmlPath;
        }

        XLXmlData* xmlData() const
        {
            return m_xmlData;
        }

        void setXmlData(XLXmlData* xmlData)
        {
            m_xmlData = xmlData;
        }

    private:
        std::string m_xmlPath {};
        XLXmlData*  m_xmlData { nullptr };
    };

}    // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_XLCOMMANDQUERY_HPP
