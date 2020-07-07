//
// Created by Kenneth Balslev on 30/06/2020.
//

#ifndef OPENXLSX_XLCOMMAND_IMPL_HPP
#define OPENXLSX_XLCOMMAND_IMPL_HPP

#include <string>

namespace OpenXLSX::Impl
{
    /**
     * @brief
     */
    enum class XLCommandType { None, SetSheetName, SetSheetColor, DeleteSheet, CloneSheet, SetSheetVisibility };

    /**
     * @brief
     */
    class XLCommand
    {
    public:
        explicit XLCommand() = default;
        explicit XLCommand(XLCommandType commandType, const std::string& sender, const std::string& parameter)
                : m_commandType(commandType),
                  m_sender(sender),
                  m_parameter(parameter) {}

        XLCommand(const XLCommand& other) = default;
        XLCommand(XLCommand&& other) noexcept = default;
        ~XLCommand() = default;
        XLCommand& operator=(const XLCommand& other) = default;
        XLCommand& operator=(XLCommand&& other) noexcept = default;

        XLCommandType commandType() const {
            return m_commandType;
        }

        const std::string& sender() const {
            return m_sender;
        }

        const std::string& parameter() const {
            return m_parameter;
        }

    private:

        XLCommandType m_commandType = XLCommandType::None;
        std::string m_sender;
        std::string m_parameter;

    };

}

#endif //OPENXLSX_XLCOMMAND_IMPL_HPP
