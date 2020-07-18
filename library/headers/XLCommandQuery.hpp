//
// Created by Kenneth Balslev on 30/06/2020.
//

#ifndef OPENXLSX_XLCOMMANDQUERY_HPP
#define OPENXLSX_XLCOMMANDQUERY_HPP

#include "XLEnums.hpp"

#include <map>
#include <string>
#include <variant>

namespace OpenXLSX
{
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

    class XLCommandAddWorksheet
    {
    public:
        XLCommandAddWorksheet(const std::string& sheetName, const std::string& sheetPath, uint16_t sheetIndex)
            : m_sheetName(sheetName),
              m_sheetPath(sheetPath),
              m_sheetIndex(sheetIndex)
        {}

        const std::string& sheetName() const
        {
            return m_sheetName;
        }

        const std::string& sheetPath() const
        {
            return m_sheetPath;
        }

        uint16_t sheetIndex() const
        {
            return m_sheetIndex;
        }

    private:
        std::string m_sheetName {};
        std::string m_sheetPath {};
        uint16_t    m_sheetIndex {};
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
        XLCommandCloneSheet(const std::string& sheetID, const std::string& sheetName) : m_sheetID(sheetID), m_sheetName(sheetName) {}

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

    using XLCommand = std::variant<XLCommandSetSheetName,
                                   XLCommandSetSheetVisibility,
                                   XLCommandSetSheetColor,
                                   XLCommandAddWorksheet,
                                   XLCommandAddChartsheet,
                                   XLCommandDeleteSheet,
                                   XLCommandCloneSheet>;

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

    template<class... Ts>
    struct overloaded : Ts...
    {
        using Ts::operator()...;
    };
    template<class... Ts>
    overloaded(Ts...) -> overloaded<Ts...>;

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLCOMMANDQUERY_HPP
