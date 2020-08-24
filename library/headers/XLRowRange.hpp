//
// Created by Kenneth Balslev on 22/08/2020.
//

#ifndef OPENXLSX_XLROWRANGE_HPP
#define OPENXLSX_XLROWRANGE_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <memory>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLRowIterator.hpp"

namespace OpenXLSX
{
    class XLSharedStrings;

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLRowRange
    {
        friend class XLRowIterator;

        //----------------------------------------------------------------------------------------------------------------------
        //           Public Member Functions
        //----------------------------------------------------------------------------------------------------------------------

    public:
        /**
         * @brief
         * @param dataNode
         * @param first
         * @param last
         * @param sharedStrings
         */
        explicit XLRowRange(const XMLNode& dataNode, uint32_t first, uint32_t last, XLSharedStrings* sharedStrings);

        /**
         * @brief
         * @param other
         */
        XLRowRange(const XLRowRange& other);

        /**
         * @brief
         * @param other
         */
        XLRowRange(XLRowRange&& other) noexcept;

        /**
         * @brief
         */
        ~XLRowRange();

        /**
         * @brief
         * @param other
         * @return
         */
        XLRowRange& operator=(const XLRowRange& other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLRowRange& operator=(XLRowRange&& other) noexcept;

        /**
         * @brief
         * @return
         */
        uint32_t rowCount() const;

        /**
         * @brief
         * @return
         */
        XLRowIterator begin();

        /**
         * @brief
         * @return
         */
        XLRowIterator end();

        /**
         * @brief
         */
        void clear();

        //----------------------------------------------------------------------------------------------------------------------
        //           Private Member Variables
        //----------------------------------------------------------------------------------------------------------------------

    private:
        std::unique_ptr<XMLNode> m_dataNode;                  /**< */
        uint32_t                 m_firstRow;                  /**< The cell reference of the first cell in the range */
        uint32_t                 m_lastRow;                   /**< The cell reference of the last cell in the range */
        XLSharedStrings*         m_sharedStrings { nullptr }; /**< */
    };
}    // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_XLROWRANGE_HPP
