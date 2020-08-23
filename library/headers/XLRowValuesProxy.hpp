//
// Created by Kenneth Balslev on 23/08/2020.
//

#ifndef OPENXLSX_XLROWVALUESPROXY_HPP
#define OPENXLSX_XLROWVALUESPROXY_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <type_traits>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlParser.hpp"
#include "XLCellValue.hpp"

// ========== CLASS AND ENUM TYPE DEFINITIONS ========== //
namespace OpenXLSX
{
    class XLRow;

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLRowValuesProxy
    {
        friend class XLRow;

    public:

        /**
         * @brief
         */
        ~XLRowValuesProxy();

        /**
         * @brief
         * @param other
         * @return
         */
        XLRowValuesProxy& operator=(const XLRowValuesProxy& other);

        /**
         * @brief
         * @param values
         * @return
         */
        XLRowValuesProxy& operator=(const std::vector<XLCellValue>& values);

        /**
         * @brief
         * @return
         */
        operator std::vector<XLCellValue>(); // NOLINT

    private:
        //---------- Private Member Functions ---------- //

        /**
         * @brief
         * @param row
         * @param rowNode
         */
        XLRowValuesProxy(XLRow* row, XMLNode* rowNode);

        /**
         * @brief
         * @param other
         */
        XLRowValuesProxy(const XLRowValuesProxy& other);

        /**
         * @brief
         * @param other
         */
        XLRowValuesProxy(XLRowValuesProxy&& other) noexcept;

        /**
         * @brief
         * @param other
         * @return
         */
        XLRowValuesProxy& operator=(XLRowValuesProxy&& other) noexcept;

        /**
         * @brief
         * @return
         */
        std::vector<XLCellValue> getValues() const;

        //---------- Private Member Variables ---------- //

        XLRow* m_row; /**< */
        XMLNode* m_rowNode; /**< */

    };

}

#pragma warning(pop)
#endif    // OPENXLSX_XLROWVALUESPROXY_HPP
