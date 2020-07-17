//
// Created by Kenneth Balslev on 30/06/2020.
//

#ifndef OPENXLSX_XLQUERY_HPP
#define OPENXLSX_XLQUERY_HPP

#include <map>
#include <string>

namespace OpenXLSX::Impl
{
    using XLQueryParams = std::map<std::string, std::string>;

    /**
     * @brief
     */
    enum class XLQueryType { None, GetSheetName, GetSheetIndex, GetSheetVisibility, GetSheetType, GetSheetID };

    /**
     * @brief
     */
    class XLQuery
    {
    public:
        /**
         * @brief
         */
        explicit XLQuery() = default;

        /**
         * @brief
         * @param queryType
         * @param sender
         * @param parameter
         */
        explicit XLQuery(XLQueryType queryType, const XLQueryParams& parameters = {}) : m_queryType(queryType), m_parameters(parameters) {}

        /**
         * @brief
         * @param other
         */
        XLQuery(const XLQuery& other) = default;

        /**
         * @brief
         * @param other
         */
        XLQuery(XLQuery&& other) noexcept = default;

        /**
         * @brief
         */
        ~XLQuery() = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLQuery& operator=(const XLQuery& other) = default;

        /**
         * @brief
         * @param other
         * @return
         */
        XLQuery& operator=(XLQuery&& other) noexcept = default;

        /**
         * @brief
         * @return
         */
        XLQueryType queryType() const
        {
            return m_queryType;
        }

        /**
         * @brief
         * @return
         */
        const XLQueryParams& parameters() const
        {
            return m_parameters;
        }

    private:
        XLQueryType   m_queryType { XLQueryType::None }; /**< */
        XLQueryParams m_parameters;                      /**< */
    };

}    // namespace OpenXLSX::Impl

#endif    // OPENXLSX_XLQUERY_HPP
