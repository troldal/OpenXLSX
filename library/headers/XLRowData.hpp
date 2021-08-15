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
#include "XLConstants.hpp"
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
     * @brief This class encapsulates the concept of a contiguous range of cells in a row.
     */
    class OPENXLSX_EXPORT XLRowDataRange
    {
        friend class XLRowDataIterator;
        friend class XLRowDataProxy;
        friend class XLRow;

    public:
        /**
         * @brief Copy constructor.
         * @param other Object to be copied.
         */
        XLRowDataRange(const XLRowDataRange& other);

        /**
         * @brief Move constructor.
         * @param other Object to be moved.
         */
        XLRowDataRange(XLRowDataRange&& other) noexcept;

        /**
         * @brief Destructor.
         */
        ~XLRowDataRange();

        /**
         * @brief Copy assignment operator.
         * @param other Object to be copied.
         * @return A reference to the copied-to object.
         */
        XLRowDataRange& operator=(const XLRowDataRange& other);

        /**
         * @brief Move assignment operator.
         * @param other Object to be moved.
         * @return A reference to the moved-to object.
         */
        XLRowDataRange& operator=(XLRowDataRange&& other) noexcept;

        /**
         * @brief Get the size (length) of the range.
         * @return The size of the range.
         */
        uint16_t size() const;

        /**
         * @brief Get an iterator to the first element.
         * @return An XLRowDataIterator pointing to the first element.
         */
        XLRowDataIterator begin();

        /**
         * @brief Get an iterator to (one-past) the last element.
         * @return An XLRowDataIterator pointing to (one past) the last element.
         */
        XLRowDataIterator end();

    private:
        /**
         * @brief Constructor.
         * @param rowNode XMLNode representing the row of the range.
         * @param firstColumn The index of the first column.
         * @param lastColumn The index of the last column.
         * @param sharedStrings A pointer to the shared strings repository.
         */
        explicit XLRowDataRange(const XMLNode& rowNode, uint16_t firstColumn, uint16_t lastColumn, XLSharedStrings* sharedStrings);

        std::unique_ptr<XMLNode> m_rowNode;                   /**< */
        uint16_t                 m_firstCol { 1 };            /**< The cell reference of the first cell in the range */
        uint16_t                 m_lastCol { 1 };             /**< The cell reference of the last cell in the range */
        XLSharedStrings*         m_sharedStrings { nullptr }; /**< */
    };

    /**
     * @brief The XLRowDataProxy is used as a proxy object when getting or setting row data. The class facilitates easy conversion
     * to/from containers.
     */
    class OPENXLSX_EXPORT XLRowDataProxy
    {
        friend class XLRow;

    public:
        /**
         * @brief Destructor
         */
        ~XLRowDataProxy();

        /**
         * @brief Copy assignment operator.
         * @param other Object to be copied.
         * @return A reference to the copied-to object.
         */
        XLRowDataProxy& operator=(const XLRowDataProxy& other);

        /**
         * @brief Assignment operator taking a std::vector of XLCellValues as an argument.
         * @param values A std::vector of XLCellValues representing the values to be assigned.
         * @return A reference to the copied-to object.
         */
        XLRowDataProxy& operator=(const std::vector<XLCellValue>& values);

        /**
         * @brief Templated assignment operator taking any container supporting forward iterators, holding XLCellValues.
         * @tparam T The container type (will be auto deducted by the compiler).
         * @param values The container of XLCellValues representing the values to be assigned.
         * @return A reference to the copied-to object.
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
         * @todo Consider providing implicit conversion to other containers.
         */
        operator std::vector<XLCellValue>();    // NOLINT

        /**
         * @brief Clears all values for the current row.
         */
        void clear();

    private:
        //---------- Private Member Functions ---------- //

        /**
         * @brief Constructor.
         * @param row Pointer to the parent XLRow object.
         * @param rowNode Pointer to the underlying XML node representing the row.
         */
        XLRowDataProxy(XLRow* row, XMLNode* rowNode);

        /**
         * @brief Copy constructor.
         * @param other Object to be copied.
         * @note The copy constructor is private in order to prevent use of the auto keyword in client code.
         */
        XLRowDataProxy(const XLRowDataProxy& other);

        /**
         * @brief Move constructor.
         * @param other Object to be moved.
         * @note Made private, as move construction should only be allowed when the parent object is moved. Disallowed for client code.
         */
        XLRowDataProxy(XLRowDataProxy&& other) noexcept;

        /**
         * @brief Move assignment operator.
         * @param other Object to be moved.
         * @return Reference to the moved-to object.
         * @note Made private, as move assignment is disallowed for client code.
         */
        XLRowDataProxy& operator=(XLRowDataProxy&& other) noexcept;

        /**
         * @brief Get the cell values for the row.
         * @return A std::vector of XLCellValues.
         */
        std::vector<XLCellValue> getValues() const;

        //---------- Private Member Variables ---------- //

        XLRow*   m_row { nullptr };     /**< Pointer to the parent XLRow object. */
        XMLNode* m_rowNode { nullptr }; /**< Pointer the the XML node representing the row. */
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
        auto range = XLRowDataRange(*m_rowNode, 1, MAX_COLS, nullptr);

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
