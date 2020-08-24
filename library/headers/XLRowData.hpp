//
// Created by Kenneth Balslev on 24/08/2020.
//

#ifndef OPENXLSX_XLROWDATA_HPP
#define OPENXLSX_XLROWDATA_HPP

#pragma warning(push)
#pragma warning(disable : 4251)
#pragma warning(disable : 4275)

// ===== External Includes ===== //
#include <memory>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLCell.hpp"
#include "XLIterator.hpp"
#include "XLXmlParser.hpp"

// ========== CLASS AND ENUM TYPE DEFINITIONS ========== //
namespace OpenXLSX
{
    class XLRowDataRange;

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLRowDataIterator
    {
    public:
        using iterator_category = std::forward_iterator_tag;
        using value_type        = XLCell;
        using difference_type   = int64_t;
        using pointer           = XLCell*;
        using reference         = XLCell&;

        /**
         * @brief
         * @param cellRange
         * @param loc
         */
        XLRowDataIterator(const XLRowDataRange& rowDataRange, XLIteratorLocation loc);

        /**
         * @brief
         */
        ~XLRowDataIterator();

        /**
         * @brief
         * @param other
         */
        XLRowDataIterator(const XLRowDataIterator& other);

        /**
         * @brief
         * @param other
         */
        XLRowDataIterator(XLRowDataIterator&& other) noexcept;

        /**
         * @brief
         * @param other
         * @return
         */
        XLRowDataIterator& operator=(const XLRowDataIterator& other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLRowDataIterator& operator=(XLRowDataIterator&& other) noexcept;

        /**
         * @brief
         * @return
         */
        XLRowDataIterator& operator++();

        /**
         * @brief
         * @return
         */
        XLRowDataIterator operator++(int);    // NOLINT

        /**
         * @brief
         * @return
         */
        reference operator*();

        /**
         * @brief
         * @return
         */
        pointer operator->();

        /**
         * @brief
         * @param rhs
         * @return
         */
        bool operator==(const XLRowDataIterator& rhs);

        /**
         * @brief
         * @param rhs
         * @return
         */
        bool operator!=(const XLRowDataIterator& rhs);

    private:
        std::unique_ptr<XMLNode> m_rowNode;                   /**< */
        uint16_t                 m_firstCol { 1 };            /**< The cell reference of the first cell in the range */
        uint16_t                 m_lastCol { 16384 };         /**< The cell reference of the last cell in the range */
        XLCell                   m_currentCell;               /**< */
        XLSharedStrings*         m_sharedStrings { nullptr }; /**< */
    };

    /**
     * @brief
     */
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
        std::unique_ptr<XMLNode> m_rowNode;                   /**< */
        uint16_t                 m_firstCol;                  /**< The cell reference of the first cell in the range */
        uint16_t                 m_lastCol;                   /**< The cell reference of the last cell in the range */
        XLSharedStrings*         m_sharedStrings { nullptr }; /**< */
    };

}    // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_XLROWDATA_HPP
