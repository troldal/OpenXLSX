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
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "OpenXLSX-Exports.hpp"
#include "XLCell.hpp"
#include "XLIterator.hpp"
#include "XLXmlParser.hpp"

// ========== CLASS AND ENUM TYPE DEFINITIONS ========== //
namespace OpenXLSX
{
    class XLRow;
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

            /**
             * @brief
             * @return
             */
            explicit operator bool() const;

        private:
            std::unique_ptr<XLRowDataRange> m_dataRange;   /**< */
            XLCell                          m_currentCell; /**< */
            uint16_t                        m_currentCol;  /**< */
            std::unique_ptr<XMLNode>        m_cellNode;    /**< */
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

        private:
            std::unique_ptr<XMLNode> m_rowNode;                   /**< */
            uint16_t                 m_firstCol { 1 };            /**< The cell reference of the first cell in the range */
            uint16_t                 m_lastCol { 1 };             /**< The cell reference of the last cell in the range */
            XLSharedStrings*         m_sharedStrings { nullptr }; /**< */
        };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLRowDataProxy
        {
        friend class XLRow;

        public:
            /**
             * @brief
             */
            ~XLRowDataProxy();

            /**
             * @brief
             * @param other
             * @return
             */
            XLRowDataProxy& operator=(const XLRowDataProxy& other);

            /**
             * @brief
             * @param values
             * @return
             */
            XLRowDataProxy& operator=(const std::vector<XLCellValue>& values);

            /**
             * @brief
             * @tparam T
             * @param values
             * @return
             */
            template<typename T,
                typename std::enable_if<!std::is_same_v<T, XLRowDataProxy> &&
                std::is_base_of_v<typename std::forward_iterator_tag,
                typename std::iterator_traits<typename T::iterator>::iterator_category>,
                T>::type* = nullptr>
                    XLRowDataProxy& operator=(const T& values);

            /**
             * @brief Implicit conversion to std::vector of XLCellValues.
             * @return A std::vector of XLCellValues.
             */
            operator std::vector<XLCellValue>();    // NOLINT

            /**
             * @brief Clears all values for the current row.
             */
            void clear();

        private:
            //---------- Private Member Functions ---------- //

            /**
             * @brief
             * @param row
             * @param rowNode
             */
            XLRowDataProxy(XLRow* row, XMLNode* rowNode);

            /**
             * @brief
             * @param other
             */
            XLRowDataProxy(const XLRowDataProxy& other);

            /**
             * @brief
             * @param other
             */
            XLRowDataProxy(XLRowDataProxy&& other) noexcept;

            /**
             * @brief
             * @param other
             * @return
             */
            XLRowDataProxy& operator=(XLRowDataProxy&& other) noexcept;

            /**
             * @brief
             * @return
             */
            std::vector<XLCellValue> getValues() const;

            //---------- Private Member Variables ---------- //

            XLRow*   m_row { nullptr };     /**< */
            XMLNode* m_rowNode { nullptr }; /**< */
        };

}    // namespace OpenXLSX

// ========== TEMPLATE MEMBER IMPLEMENTATIONS ========== //
namespace OpenXLSX
{
    template<typename T,
        typename std::enable_if<!std::is_same_v<T, XLRowDataProxy> &&
        std::is_base_of_v<typename std::forward_iterator_tag,
        typename std::iterator_traits<typename T::iterator>::iterator_category>,
        T>::type*>
        XLRowDataProxy& XLRowDataProxy::operator=(const T& values)
            {
            auto range = XLRowDataRange(*m_rowNode, 1, 16384, nullptr);

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
}    // namespace OpenXLSX

#pragma warning(pop)
#endif    // OPENXLSX_XLROWDATA_HPP
