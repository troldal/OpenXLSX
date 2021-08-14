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
     * @brief This class encapsilates a (non-const) iterator, for iterating over the cells in a row.
     * @todo Consider implementing a const iterator also
     */
    class OPENXLSX_EXPORT XLRowDataIterator
        {
            friend class XLRowDataRange;

        public:
            using iterator_category = std::forward_iterator_tag;
            using value_type        = XLCell;
            using difference_type   = int64_t;
            using pointer           = XLCell*;
            using reference         = XLCell&;

            /**
             * @brief Destructor.
             */
            ~XLRowDataIterator();

            /**
             * @brief Copy constructor.
             * @param other Object to be copied.
             */
            XLRowDataIterator(const XLRowDataIterator& other);

            /**
             * @brief Move constructor.
             * @param other Object to be moved.
             */
            XLRowDataIterator(XLRowDataIterator&& other) noexcept;

            /**
             * @brief Copy assignment operator.
             * @param other Object to be copied.
             * @return Reference to the copied-to object.
             */
            XLRowDataIterator& operator=(const XLRowDataIterator& other);

            /**
             * @brief Move assignment operator.
             * @param other Object to be moved.
             * @return Reference to the moved-to object.
             */
            XLRowDataIterator& operator=(XLRowDataIterator&& other) noexcept;

            /**
             * @brief Pre-increment of the iterator.
             * @return Reference to the iterator object.
             */
            XLRowDataIterator& operator++();

            /**
             * @brief Post-increment of the iterator.
             * @return Reference to the iterator object.
             */
            XLRowDataIterator operator++(int);    // NOLINT

            /**
             * @brief Dereferencing operator.
             * @return Reference to the object pointed to by the iterator.
             */
            reference operator*();

            /**
             * @brief Arrow operator.
             * @return Pointer to the object pointed to by the iterator.
             */
            pointer operator->();

            /**
             * @brief Equality operator.
             * @param rhs XLRowDataIterator to compare to.
             * @return true if equal, otherwise false.
             */
            bool operator==(const XLRowDataIterator& rhs);

            /**
             * @brief Non-equality operator.
             * @param rhs XLRowDataIterator to compare to.
             * @return false if equal, otherwise true.
             */
            bool operator!=(const XLRowDataIterator& rhs);

        private:

            /**
             * @brief Constructor.
             * @param cellRange The range to iterate over.
             * @param loc The location of the iterator (begin or end).
             */
            XLRowDataIterator(const XLRowDataRange& rowDataRange, XLIteratorLocation loc);

            std::unique_ptr<XLRowDataRange> m_dataRange;   /**< A pointer to the range to iterate over. */
            std::unique_ptr<XMLNode>        m_cellNode;    /**< The XML node representing the cell currently pointed at. */
            XLCell                          m_currentCell; /**< The XLCell currently pointed at. */
        };

    /**
     * @brief
     */
    class OPENXLSX_EXPORT XLRowDataRange
        {
        friend class XLRowDataIterator;
        friend class XLRow;

        public:

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
            uint16_t size() const;

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

            /**
             * @brief
             * @param rowNode
             * @param firstColumn
             * @param lastColumn
             * @param sharedStrings
             */
            explicit XLRowDataRange(const XMLNode& rowNode, uint16_t firstColumn, uint16_t lastColumn, XLSharedStrings* sharedStrings);


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
