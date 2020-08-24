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
#include <iterator>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLXmlParser.hpp"
#include "XLCellValue.hpp"
#include "XLRowDataRange.hpp"

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

        template<typename T,
                 typename std::enable_if<!std::is_same_v<T, XLRowValuesProxy> && std::is_base_of_v<typename std::forward_iterator_tag,
                                                                                                   typename std::iterator_traits<typename T::iterator>::iterator_category>, T>::type* = nullptr>
        XLRowValuesProxy& operator=(const T& values);

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

}  // namespace OpenXLSX

namespace OpenXLSX {

    template<typename T,
        typename std::enable_if<!std::is_same_v<T, XLRowValuesProxy> && std::is_base_of_v<typename std::forward_iterator_tag,
                                                                                          typename std::iterator_traits<typename T::iterator>::iterator_category>, T>::type*>
    XLRowValuesProxy& XLRowValuesProxy::operator=(const T& values) {

        auto range = XLRowDataRange(*m_rowNode,1,16384,nullptr);

        auto dst = range.begin();
        auto src = values.begin();

        while (true) {
            dst->value() = *src;
            ++src;
            if (src == values.end()) break;
            ++dst;
        }

        return *this;
    }
}  // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_XLROWVALUESPROXY_HPP
