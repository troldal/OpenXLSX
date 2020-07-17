//
// Created by Kenneth Balslev on 30/06/2020.
//

#ifndef OPENXLSX_XLCOMMAND_HPP
#define OPENXLSX_XLCOMMAND_HPP

#include <map>
#include <string>
#include <variant>

namespace OpenXLSX
{
    /**
     * @brief
     */
    using XLCommandParams = std::map<std::string, std::string>;

    /**
     * @brief
     */
    enum class XLCommandType {
        None,
        SetSheetName,
        SetSheetVisibility,
        SetSheetColor,
        AddWorksheet,
        AddChartsheet,
        DeleteSheet,
        CloneSheet
    };

    /**
     * @brief
     */
    //    class XLCommand
    //    {
    //    public:
    //        /**
    //         * @brief
    //         */
    //        XLCommand() = default;
    //
    //        /**
    //         * @brief
    //         * @param commandType
    //         * @param parameters
    //         * @param sender
    //         */
    //        XLCommand(XLCommandType commandType, const XLCommandParams& parameters, const std::string& sender)
    //            : m_commandType(commandType),
    //              m_parameters(parameters),
    //              m_sender(sender)
    //        {}
    //
    //        /**
    //         * @brief
    //         * @param other
    //         */
    //        XLCommand(const XLCommand& other) = default;
    //
    //        /**
    //         * @brief
    //         * @param other
    //         */
    //        XLCommand(XLCommand&& other) noexcept = default;
    //
    //        /**
    //         * @brief
    //         */
    //        ~XLCommand() = default;
    //
    //        /**
    //         * @brief
    //         * @param other
    //         * @return
    //         */
    //        XLCommand& operator=(const XLCommand& other) = default;
    //
    //        /**
    //         * @brief
    //         * @param other
    //         * @return
    //         */
    //        XLCommand& operator=(XLCommand&& other) noexcept = default;
    //
    //        /**
    //         * @brief
    //         * @return
    //         */
    //        XLCommandType commandType() const
    //        {
    //            return m_commandType;
    //        }
    //
    //        /**
    //         * @brief
    //         * @return
    //         */
    //        const std::string& sender() const
    //        {
    //            return m_sender;
    //        }
    //
    //        /**
    //         * @brief
    //         * @return
    //         */
    //        const XLCommandParams& parameters() const
    //        {
    //            return m_parameters;
    //        }
    //
    //    private:
    //        XLCommandType   m_commandType { XLCommandType::None }; /**< */
    //        XLCommandParams m_parameters;                          /**< */
    //        std::string     m_sender;                              /**< */
    //    };

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

    template<class... Ts>
    struct overloaded : Ts...
    {
        using Ts::operator()...;
    };
    template<class... Ts>
    overloaded(Ts...) -> overloaded<Ts...>;

}    // namespace OpenXLSX

#endif    // OPENXLSX_XLCOMMAND_HPP
