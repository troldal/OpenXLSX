//
// Created by Kenneth Balslev on 24/08/2020.
//

#ifndef OPENXLSX_XLROWDATARANGE_HPP
#define OPENXLSX_XLROWDATARANGE_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <memory>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLRowDataIterator.hpp"

namespace OpenXLSX
{
    class OPENXLSX_EXPORT XLRowDataRange
    {
        friend class XLRowDataIterator;

    public:
        /**
         * @brief
         * @param dataNode
         * @param first
         * @param last
         * @param sharedStrings
         */
        explicit XLRowDataRange(const XMLNode& rowNode, uint16_t firstColumn, uint16_t lastColumn, XLSharedStrings* sharedStrings);

        /**
         * @brief
         * @param other
         */
        XLRowDataRange(const XLRowDataRange& other);

        /**
         * @brief
         * @param other
         */
        XLRowDataRange(XLRowDataRange&& other) noexcept;

        /**
         * @brief
         */
        ~XLRowDataRange();

        /**
         * @brief
         * @param other
         * @return
         */
        XLRowDataRange& operator=(const XLRowDataRange& other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLRowDataRange& operator=(XLRowDataRange&& other) noexcept;

        /**
         * @brief
         * @return
         */
        uint16_t cellCount() const;

        /**
         * @brief
         * @return
         */
        XLRowDataIterator begin();

        /**
         * @brief
         * @return
         */
        XLRowDataIterator end();

        /**
         * @brief
         */
        void clear();

    private:
        std::unique_ptr<XMLNode> m_rowNode;                  /**< */
        uint16_t                 m_firstCol;                  /**< The cell reference of the first cell in the range */
        uint16_t                 m_lastCol;                   /**< The cell reference of the last cell in the range */
        XLSharedStrings*         m_sharedStrings { nullptr }; /**< */

    };

} // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_XLROWDATARANGE_HPP
