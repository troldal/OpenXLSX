//
// Created by Kenneth Balslev on 24/08/2020.
//

#include "XLRowData.hpp"
#include "XLCell.hpp"
#include "XLRow.hpp"

#include <pugixml.hpp>

namespace OpenXLSX
{
    XMLNode getCellNode(XMLNode rowNode, uint16_t columnNumber);
}

namespace
{
    using namespace OpenXLSX;
    template<typename T>
    void setValues(const T& values, uint64_t rowNumber, XMLNode* rowNode, XLSharedStrings* sharedStrings)
    {
        // ===== Find or create the first cell node
        auto curNode = rowNode->first_child();
        if (!curNode || XLCellReference(curNode.attribute("r").value()).column() != 1) {
            curNode = rowNode->append_child("c");
            curNode.append_attribute("r").set_value(XLCellReference(rowNumber, 1).address().c_str());
        }

        auto     prevNode = XMLNode();
        uint16_t col      = 1;
        for (const auto& value : values) {
            if (!curNode || XLCellReference(curNode.attribute("r").value()).column() != col) {
                curNode = rowNode->insert_child_after("c", prevNode);
                curNode.append_attribute("r").set_value(XLCellReference(rowNumber, col).address().c_str());
            }

            XLCell(curNode, sharedStrings).value() = value;

            prevNode = curNode;
            curNode  = curNode.next_sibling();
            ++col;
        }
    }
}    // namespace

// ========== XLRowDataIterator  ============================================ //
namespace OpenXLSX
{
    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator::XLRowDataIterator(const XLRowDataRange& rowDataRange, XLIteratorLocation loc)
        : m_rowNode(std::make_unique<XMLNode>(*rowDataRange.m_rowNode)),
          m_firstCol(rowDataRange.m_firstCol),
          m_lastCol(rowDataRange.m_lastCol),
          m_currentCell(),
          m_sharedStrings(rowDataRange.m_sharedStrings)
    {
        if (loc == XLIteratorLocation::End)
            m_currentCell = XLCell();
        else {
            m_currentCell = XLCell(getCellNode(*m_rowNode, m_firstCol), m_sharedStrings);
        }
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator::~XLRowDataIterator() = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator::XLRowDataIterator(const XLRowDataIterator& other)
        : m_rowNode(std::make_unique<XMLNode>(*other.m_rowNode)),
          m_firstCol(other.m_firstCol),
          m_lastCol(other.m_lastCol),
          m_currentCell(other.m_currentCell),
          m_sharedStrings(other.m_sharedStrings)
    {}

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator::XLRowDataIterator(XLRowDataIterator&& other) noexcept = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator& XLRowDataIterator::operator=(const XLRowDataIterator& other)
    {
        if (&other != this) {
            *m_rowNode      = *other.m_rowNode;
            m_firstCol      = other.m_firstCol;
            m_lastCol       = other.m_lastCol;
            m_currentCell   = other.m_currentCell;
            m_sharedStrings = other.m_sharedStrings;
        }

        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator& XLRowDataIterator::operator=(XLRowDataIterator&& other) noexcept = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator& XLRowDataIterator::operator++()
    {
        auto cellNumber = m_currentCell.cellReference().column() + 1;
        auto cellNode   = m_currentCell.m_cellNode->next_sibling();

        if (cellNumber > m_lastCol)
            m_currentCell = XLCell();

        else if (!cellNode || XLCellReference(cellNode.attribute("r").value()).column() != cellNumber) {
            cellNode = m_rowNode->insert_child_after("c", *m_currentCell.m_cellNode);
            cellNode.append_attribute("r").set_value(XLCellReference(m_rowNode->attribute("r").as_ullong(), cellNumber).address().c_str());
            m_currentCell = XLCell(cellNode, m_sharedStrings);
        }

        else
            m_currentCell = XLCell(cellNode, m_sharedStrings);

        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator XLRowDataIterator::operator++(int)
    {
        auto oldIter(*this);
        ++(*this);
        return oldIter;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLCell& XLRowDataIterator::operator*()
    {
        return m_currentCell;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator::pointer XLRowDataIterator::operator->()
    {
        return &m_currentCell;
    }

    /**
     * @details
     * @pre
     * @post
     */
    bool XLRowDataIterator::operator==(const XLRowDataIterator& rhs)
    {
        return m_currentCell == rhs.m_currentCell;
    }

    /**
     * @details
     * @pre
     * @post
     */
    bool XLRowDataIterator::operator!=(const XLRowDataIterator& rhs)
    {
        return !(m_currentCell == rhs.m_currentCell);
    }
}    // namespace OpenXLSX

// ========== XLRowDataRange  =============================================== //
namespace OpenXLSX
{
    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataRange::XLRowDataRange(const XMLNode& rowNode, uint16_t firstColumn, uint16_t lastColumn, XLSharedStrings* sharedStrings)
        : m_rowNode(std::make_unique<XMLNode>(rowNode)),
          m_firstCol(firstColumn),
          m_lastCol(lastColumn),
          m_sharedStrings(sharedStrings)
    {}

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataRange::XLRowDataRange(const XLRowDataRange& other)
        : m_rowNode(std::make_unique<XMLNode>(*other.m_rowNode)),
          m_firstCol(other.m_firstCol),
          m_lastCol(other.m_lastCol),
          m_sharedStrings(other.m_sharedStrings)

    {}

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataRange::XLRowDataRange(XLRowDataRange&& other) noexcept = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataRange::~XLRowDataRange() = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataRange& XLRowDataRange::operator=(const XLRowDataRange& other)
    {
        if (&other != this) {
            *m_rowNode      = *other.m_rowNode;
            m_firstCol      = other.m_firstCol;
            m_lastCol       = other.m_lastCol;
            m_sharedStrings = other.m_sharedStrings;
        }

        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataRange& XLRowDataRange::operator=(XLRowDataRange&& other) noexcept
    {
        if (&other != this) {
            *m_rowNode      = *other.m_rowNode;
            m_firstCol      = other.m_firstCol;
            m_lastCol       = other.m_lastCol;
            m_sharedStrings = other.m_sharedStrings;
        }

        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    uint16_t XLRowDataRange::cellCount() const
    {
        return m_lastCol - m_firstCol + 1;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator XLRowDataRange::begin()
    {
        return XLRowDataIterator(*this, XLIteratorLocation::Begin);
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataIterator XLRowDataRange::end()
    {
        return XLRowDataIterator(*this, XLIteratorLocation::End);
    }

    /**
     * @details
     * @pre
     * @post
     */
    void XLRowDataRange::clear() {}
}    // namespace OpenXLSX

// ========== XLRowDataProxy  =============================================== //
namespace OpenXLSX
{
    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy::~XLRowDataProxy() = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy& XLRowDataProxy::operator=(const XLRowDataProxy& other)
    {
        if (&other != this) {
            *this = other.getValues();
        }

        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy::XLRowDataProxy(XLRow* row, XMLNode* rowNode) : m_row(row), m_rowNode(rowNode) {}

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy::XLRowDataProxy(const XLRowDataProxy& other) = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy::XLRowDataProxy(XLRowDataProxy&& other) noexcept = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy& XLRowDataProxy::operator=(XLRowDataProxy&& other) noexcept = default;

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy& XLRowDataProxy::operator=(const std::vector<XLCellValue>& values)
    {
        setValues(values, m_row->rowNumber(), m_rowNode, m_row->m_sharedStrings);
        return *this;
    }

    /**
     * @details
     * @pre
     * @post
     */
    XLRowDataProxy::operator std::vector<XLCellValue>()
    {
        return getValues();
    }

    /**
     * @details
     * @pre
     * @post
     */
    std::vector<XLCellValue> XLRowDataProxy::getValues() const
    {
        auto                     numCells = XLCellReference(m_rowNode->last_child().attribute("r").value()).column();
        std::vector<XLCellValue> result(numCells);

        for (auto& node : m_rowNode->children()) {
            if (XLCellReference(node.attribute("r").value()).column() > numCells) break;
            result[XLCellReference(node.attribute("r").value()).column() - 1] = XLCell(node, m_row->m_sharedStrings).value();
        }

        return result;
    }

}    // namespace OpenXLSX