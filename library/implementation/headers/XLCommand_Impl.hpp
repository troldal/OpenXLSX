//
// Created by Kenneth Balslev on 30/06/2020.
//

#ifndef OPENXLSX_XLCOMMAND_IMPL_HPP
#define OPENXLSX_XLCOMMAND_IMPL_HPP

#include <map>
#include <string>

namespace OpenXLSX::Impl
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
        SetSheetColor,
        AddWorksheet,
        AddChartsheet,
        DeleteSheet,
        CloneSheet,
        SetSheetVisibility
    };

    /**
     * @brief
     */
    class XLCommand
    {
    public:
        /**
         * @brief
         */
        XLCommand() = default;

        /**
         * @brief
         * @param commandType
         * @param parameters
         * @param sender
         */
        XLCommand(XLCommandType commandType, const XLCommandParams& parameters, const std::string& sender)
            : m_commandType(commandType),
              m_parameters(parameters),
              m_sender(sender)
        {}

        /**
         * @brief
         * @param other
         */
        XLCommand(const XLCommand& other) = default;

        /**
         * @brief
         * @param other
         */
        XLCommand(XLCommand&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLCommand() = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLCommand& operator=(const XLCommand& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLCommand& operator=(XLCommand&& other) noexcept = default;

        /**
         * @brief
         * @return
         */
        XLCommandType commandType() const
        {
            return m_commandType;
        }

        /**
         * @brief
         * @return
         */
        const std::string& sender() const
        {
            return m_sender;
        }

        /**
         * @brief
         * @return
         */
        const XLCommandParams& parameters() const
        {
            return m_parameters;
        }

    private:
        XLCommandType   m_commandType { XLCommandType::None }; /**< */
        XLCommandParams m_parameters;                          /**< */
        std::string     m_sender;                              /**< */
    };

}    // namespace OpenXLSX::Impl

#endif    // OPENXLSX_XLCOMMAND_IMPL_HPP
