//
// Created by Kenneth Balslev on 30/06/2020.
//

#ifndef OPENXLSX_XLQUERY_IMPL_HPP
#define OPENXLSX_XLQUERY_IMPL_HPP

#include <string>

namespace OpenXLSX::Impl
{
    /**
     * @brief
     */
    enum class XLQueryType { None, GetSheetName, GetSheetIndex, GetSheetVisibility };

    /**
     * @brief
     */
    class XLQuery
    {
    public:
        explicit XLQuery() = default;
        explicit XLQuery(XLQueryType queryType, const std::string& sender, const std::string& parameter = "")
                : m_queryType(queryType),
                  m_subject(sender),
                  m_parameter(parameter) {}

        XLQuery(const XLQuery& other) = default;
        XLQuery(XLQuery&& other) noexcept = default;
        ~XLQuery() = default;
        XLQuery& operator=(const XLQuery& other) = default;
        XLQuery& operator=(XLQuery&& other) noexcept = default;

        XLQueryType queryType() const {
            return m_queryType;
        }

        const std::string& subject() const {
            return m_subject;
        }

        const std::string& parameter() const {
            return m_parameter;
        }

    private:

        XLQueryType m_queryType = XLQueryType::None;
        std::string m_subject;
        std::string m_parameter;

    };

}

#endif //OPENXLSX_XLQUERY_IMPL_HPP
