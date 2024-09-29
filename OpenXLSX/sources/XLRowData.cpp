/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.
 8b    d8   MM.  ,M9 YM    d9 MM    MM  d'    `MM.   MM    / L    ,M9  d'    `MM.
  YMMMM9    MMYMMM9   YMMMM9 _MM_  _MM_M(_    _)MM_ _MMMMMMM MYMMMM9 _M(_    _)MM_
            MM
            MM
           _MM_

  Copyright (c) 2018, Kenneth Troldal Balslev

  All rights reserved.

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  - Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.
  - Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.
  - Neither the name of the author nor the
    names of any contributors may be used to endorse or promote products
    derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

// ===== External Includes ===== //
#include <algorithm>
#include <cassert>

// ===== OpenXLSX Includes ===== //
#include "XLCell.hpp"
#include "XLRow.hpp"
#include "XLRowData.hpp"
#include "utilities/XLUtilities.hpp"

// ========== XLRowDataIterator  ============================================ //
namespace OpenXLSX
{
    /**
     * @details Constructor.
     * @pre The given range and location are both valid.
     * @post
     * @note 2024-05-28: added support for constructing with an empty m_cellNode from an empty rowDataRange which will allow obtaining
     *       an XLIteratorLocation::End for such a range so that iterations can fail in a controlled manner
     */
    XLRowDataIterator::XLRowDataIterator(const XLRowDataRange& rowDataRange, XLIteratorLocation loc)
        : m_dataRange(std::make_unique<XLRowDataRange>(rowDataRange)),
          m_cellNode(std::make_unique<XMLNode>(
              getCellNode((m_dataRange->size() ? *m_dataRange->m_rowNode : XMLNode {}), m_dataRange->m_firstCol))),
          m_currentCell(loc == XLIteratorLocation::End ? XLCell() : XLCell(*m_cellNode, m_dataRange->m_sharedStrings))
    {}

    /**
     * @details Destructor. Default implementation.
     * @pre
     * @post
     */
    XLRowDataIterator::~XLRowDataIterator() = default;

    /**
     * @details Copy constructor. Trivial implementation with deep copy of pointer members.
     * @pre
     * @post
     */
    XLRowDataIterator::XLRowDataIterator(const XLRowDataIterator& other)
        : m_dataRange(std::make_unique<XLRowDataRange>(*other.m_dataRange)),
          m_cellNode(std::make_unique<XMLNode>(*other.m_cellNode)),
          m_currentCell(other.m_currentCell)
    {}

    /**
     * @details Move constructor. Default implementation.
     * @pre
     * @post
     */
    XLRowDataIterator::XLRowDataIterator(XLRowDataIterator&& other) noexcept = default;    // NOLINT

    /**
     * @details Copy assignment operator. Implemented using copy-and-swap idiom.
     * @pre
     * @post
     */
    XLRowDataIterator& XLRowDataIterator::operator=(const XLRowDataIterator& other)
    {
        if (&other != this) {
            XLRowDataIterator temp = other;
            std::swap(temp, *this);
        }

        return *this;
    }

    /**
     * @details Move assignment operator. Default implementation.
     * @pre
     * @post
     */
    XLRowDataIterator& XLRowDataIterator::operator=(XLRowDataIterator&& other) noexcept = default;

    /**
     * @details Pre-increment operator. Advances the iterator one element.
     * @pre
     * @post
     */
    XLRowDataIterator& XLRowDataIterator::operator++()
    {
        // ===== Compute the column number, and move the m_cellNode to the next sibling.
        const uint16_t cellNumber = m_currentCell.cellReference().column() + 1;
        XMLNode        cellNode   = m_currentCell.m_cellNode->next_sibling_of_type(pugi::node_element);

        // ===== If the cellNumber exceeds the last column in the range,
        // ===== m_currentCell is set to an empty XLCell, indicating the end of the range has been reached.
        if (cellNumber > m_dataRange->m_lastCol) m_currentCell = XLCell();

        // ====== If the cellNode is empty (i.e. no more children in the current row node) or the column number of the cell node
        // ====== is higher than the computed column number, then insert the node.
        // BUG BUGFIX 2024-04-26: check was for m_cellNode->empty(), allowing an invalid test for the attribute r, discovered
        //       because the modified XLCellReference throws an exception on invalid parameter
        else if (cellNode.empty() || XLCellReference(cellNode.attribute("r").value()).column() > cellNumber) {
            cellNode = m_dataRange->m_rowNode->insert_child_after("c", *m_currentCell.m_cellNode);
            setDefaultCellAttributes(cellNode, XLCellReference(
            /**/                                   static_cast<uint32_t>(m_dataRange->m_rowNode->attribute("r").as_ullong()), cellNumber
            /**/                               ).address(),
            /**/                               *m_dataRange->m_rowNode, cellNumber);
            m_currentCell = XLCell(cellNode, m_dataRange->m_sharedStrings);
        }

        // ===== Otherwise, the cell node and the column number match.
        else {
            assert(XLCellReference(cellNode.attribute("r").value()).column() == cellNumber);
            m_currentCell = XLCell(cellNode, m_dataRange->m_sharedStrings);
        }

        return *this;
    }

    /**
     * @details Post-increment operator. Implemented in terms of the pre-increment operator.
     * @pre
     * @post
     */
    XLRowDataIterator XLRowDataIterator::operator++(int)    // NOLINT
    {
        auto oldIter(*this);
        ++(*this);
        return oldIter;
    }

    /**
     * @details Dereferencing operator.
     * @pre
     * @post
     */
    XLCell& XLRowDataIterator::operator*() { return m_currentCell; }

    /**
     * @details Arrow operator.
     * @pre
     * @post
     */
    XLRowDataIterator::pointer XLRowDataIterator::operator->() { return &m_currentCell; }

    /**
     * @details Equality comparison operator.
     * @pre
     * @post
     */
    bool XLRowDataIterator::operator==(const XLRowDataIterator& rhs) const
    {
        // 2024-05-28 BUGFIX: (!m_currentCell && rhs.m_currentCell) was not evaluated, triggering a segmentation fault on dereferencing
        if (static_cast<bool>(m_currentCell) != static_cast<bool>(rhs.m_currentCell)) return false;
        // ===== If execution gets here, current cells are BOTH valid or BOTH invalid / empty
        if (not m_currentCell) return true;    // checking one for being empty is enough to know both are empty
        return m_currentCell == rhs.m_currentCell;
    }

    /**
     * @details Non-equality comparison operator.
     * @pre
     * @post
     */
    bool XLRowDataIterator::operator!=(const XLRowDataIterator& rhs) const { return !(*this == rhs); }

}    // namespace OpenXLSX

// ========== XLRowDataRange  =============================================== //
namespace OpenXLSX
{
    /**
     * @details Constructor. Trivial implementation.
     * @throws If firstColumn > than lastColumn, an XLOverflowError will be thrown.
     * @pre
     * @post
     */
    XLRowDataRange::XLRowDataRange(const XMLNode& rowNode, uint16_t firstColumn, uint16_t lastColumn, const XLSharedStrings& sharedStrings)
        : m_rowNode(std::make_unique<XMLNode>(rowNode)),
          m_firstCol(firstColumn),
          m_lastCol(lastColumn),
          m_sharedStrings(sharedStrings)
    {
        if (lastColumn < firstColumn) {
            m_firstCol = 1;
            m_lastCol  = 1;
            throw XLOverflowError("lastColumn is less than firstColumn.");
        }
    }

    /**
     * @details Constructs an empty XLDataRange, whose size() will return 0. To be used as return value in functions that shall fail without
     * exception.
     */
    XLRowDataRange::XLRowDataRange()
        : m_rowNode(nullptr),
          m_firstCol(1),    // first col of 1
          m_lastCol(0),     // and last col of 0 will ensure that size returns 0
          m_sharedStrings()
    {
        // nothing to do
    }

    /**
     * @details Copy constructor. Trivial implementation.
     * @pre
     * @post
     */
    XLRowDataRange::XLRowDataRange(const XLRowDataRange& other)
        : m_rowNode((other.m_rowNode != nullptr) ? std::make_unique<XMLNode>(*other.m_rowNode) : nullptr),    // 2024-05-28: support for copy-construction from an empty XLDataRange
          m_firstCol(other.m_firstCol),
          m_lastCol(other.m_lastCol),
          m_sharedStrings(other.m_sharedStrings)

    {}

    /**
     * @details Move constructor. Default implementation.
     * @pre
     * @post
     */
    XLRowDataRange::XLRowDataRange(XLRowDataRange&& other) noexcept = default;

    /**
     * @details Destructor. Default implementation.
     * @pre
     * @post
     */
    XLRowDataRange::~XLRowDataRange() = default;

    /**
     * @details Copy assignment operator. Implemented in terms of copy-and-swap.
     * @pre
     * @post
     */
    XLRowDataRange& XLRowDataRange::operator=(const XLRowDataRange& other)
    {
        if (&other != this) {
            XLRowDataRange temp(other);
            std::swap(temp, *this);
        }

        return *this;
    }

    /**
     * @details Move assignment operator. Default implementation.
     * @pre
     * @post
     */
    XLRowDataRange& XLRowDataRange::operator=(XLRowDataRange&& other) noexcept = default;

    /**
     * @details Calculates the size (number of cells) in the range.
     * @pre
     * @post
     */
    uint16_t XLRowDataRange::size() const { return m_lastCol - m_firstCol + 1; }

    /**
     * @details Get an iterator to the first cell in the range.
     * @pre
     * @post
     * @note 2024-05-28: enhanced ::begin() to return an end iterator for an empty range
     */
    XLRowDataIterator XLRowDataRange::begin() { return XLRowDataIterator { *this, (size() > 0 ? XLIteratorLocation::Begin : XLIteratorLocation::End) }; }

    /**
     * @details Get an iterator to (one past) the last cell in the range.
     * @pre
     * @post
     */
    XLRowDataIterator XLRowDataRange::end() { return XLRowDataIterator { *this, XLIteratorLocation::End }; }

}    // namespace OpenXLSX

// ========== XLRowDataProxy  =============================================== //
namespace OpenXLSX
{
    /**
     * @details Destructor. Default implementation.
     * @pre
     * @post
     */
    XLRowDataProxy::~XLRowDataProxy() = default;

    /**
     * @details Copy constructor. This is not a 'true' copy constructor, as it is the row values that will
     * be copied, not the XLRowDataProxy member variables (pointers to the XLRow and row node objects).
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
     * @details Constructor
     * @pre
     * @post
     */
    XLRowDataProxy::XLRowDataProxy(XLRow* row, XMLNode* rowNode) : m_row(row), m_rowNode(rowNode) {}

    /**
     * @details Copy constructor. Default implementation.
     * @pre
     * @post
     */
    XLRowDataProxy::XLRowDataProxy(const XLRowDataProxy& other) = default;

    /**
     * @details Move constructor. Default implementation.
     * @pre
     * @post
     */
    XLRowDataProxy::XLRowDataProxy(XLRowDataProxy&& other) noexcept = default;

    /**
     * @details Move assignment operator. Default implementation.
     * @pre
     * @post
     */
    XLRowDataProxy& XLRowDataProxy::operator=(XLRowDataProxy&& other) noexcept = default;

    /**
     * @details Assignment operator taking a std::vector of XLCellValue objects as an argument. Other container types
     * and/or value types will be handled by the templated operator=. However, because assigning a std::vector of
     * XLCellValue object is the most common case, this case is handled separately for higher performance.
     * @pre
     * @post
     */
    XLRowDataProxy& XLRowDataProxy::operator=(const std::vector<XLCellValue>& values)    // 2024-04-30: whitespace support
    {
        if (values.size() > MAX_COLS) throw XLOverflowError("vector<XLCellValue> size exceeds maximum number of columns.");
        if (values.empty()) return *this;

        deleteCellValues(static_cast<uint16_t>(values.size()));    // 2024-04-30: whitespace support

        // ===== prepend new cell nodes to current row node
        XMLNode curNode{};
        uint16_t colNo = values.size();
        for (auto value = values.rbegin(); value != values.rend(); ++value) {    // NOLINT
            curNode = m_rowNode->prepend_child("c");
            setDefaultCellAttributes(curNode, XLCellReference(static_cast<uint32_t>(m_row->rowNumber()), colNo).address(), *m_rowNode, colNo);
            XLCell(curNode, m_row->m_sharedStrings).value() = *value;
            --colNo;
        }

        return *this;
    }

    /**
     * @details Assignment operator taking a std::vector of bool values as an argument. Under most circumstances,
     * one of the templated assignment operators should do the job. However, because std::vector<bool> is handled
     * differently than other value types, some compilers don't play well with using std::vector<bool> in a
     * template function. Therefore this edge case is handled separately.
     * @pre
     * @post
     */
    XLRowDataProxy& XLRowDataProxy::operator=(const std::vector<bool>& values)    // 2024-04-30: whitespace support
    {
        if (values.size() > MAX_COLS) throw XLOverflowError("vector<bool> size exceeds maximum number of columns.");
        if (values.empty()) return *this;

        auto range = XLRowDataRange(*m_rowNode, 1, static_cast<uint16_t>(values.size()), getSharedStrings());
        auto dst   = range.begin();    // 2024-04-30: whitespace support: safe because XLRowDataRange::begin invokes whitespace-safe
                                       // getCellNode for column 1
        auto src = values.begin();

        while (true) {
            dst->value() = static_cast<bool>(*src);
            ++src;
            if (src == values.end()) break;
            ++dst;    // 2024-04-30: whitespace support: XLRowDataIterator::operator++ is whitespace-safe
        }

        return *this;
    }

    /**
     * @details This function simply calls the getValues() function, which returns a std::vector of XLCellValues as required.
     * @pre
     * @post
     */
    XLRowDataProxy::operator std::vector<XLCellValue>() const { return getValues(); }

    /**
     * @details Calls the convertContainer convenience function with a std::deque of XLCellValues as an argument.
     * @pre
     * @post
     */
    XLRowDataProxy::operator std::deque<XLCellValue>() const { return convertContainer<std::deque<XLCellValue>>(); }

    /**
     * @details Calls the convertContainer convenience function with a std::list of XLCellValues as an argument.
     * @pre
     * @post
     */
    XLRowDataProxy::operator std::list<XLCellValue>() const { return convertContainer<std::list<XLCellValue>>(); }

    /**
     * @details Iterates through the cell values (if any) for the current row, and copies them to an output std::vector of XLCellValues.
     * @pre
     * @post
     */
    std::vector<XLCellValue> XLRowDataProxy::getValues() const    // 2024-04-30: whitespace support
    {
        // ===== Determine the number of cells in the current row. Create a std::vector of the same size.
        const XMLNode  lastElementChild = m_rowNode->last_child_of_type(pugi::node_element);
        const uint16_t numCells = (lastElementChild.empty() ? 0 : XLCellReference(lastElementChild.attribute("r").value()).column());
        std::vector<XLCellValue> result(static_cast<uint64_t>(numCells));

        // ===== If there are one or more cells in the current row, iterate through them and add the value to the container.
        if (numCells > 0) {
            XMLNode node = lastElementChild;    // avoid unneeded call to first_child_of_type by iterating backwards, vector is random
                                                // access so it doesn't matter
            while (not node.empty()) {
                result[XLCellReference(node.attribute("r").value()).column() - 1] = XLCell(node, m_row->m_sharedStrings).value();
                node                                                              = node.previous_sibling_of_type(pugi::node_element);
            }
        }

        // ===== Return the resulting container.
        return result;
    }

    /**
     * @details The function returns a pointer to an XLSharedStrings object embedded in the m_row member.
     * This is required because the XLRow class internals is not visible in the header file.
     * @pre
     * @post
     */
    XLSharedStrings XLRowDataProxy::getSharedStrings() const { return m_row->m_sharedStrings; }

    /**
     * @details The deleteCellValues is a convenience function used solely by the templated operator= function.
     * The purpose of a separate function is to keep details of xml_node out of the header file.
     * @pre
     * @post
     */
    void XLRowDataProxy::deleteCellValues(uint16_t count)    // NOLINT   // 2024-04-30: whitespace support
    {
        // ===== Mark cell nodes for deletion
        std::vector<XMLNode> toBeDeleted;
        XMLNode              cellNode = m_rowNode->first_child_of_type(pugi::node_element);
        while (not cellNode.empty()) {
            if (XLCellReference(cellNode.attribute("r").value()).column() <= count) {
                toBeDeleted.emplace_back(cellNode);
                XMLNode nextNode = cellNode.next_sibling();    // get next "regular" sibling (any type) before advancing cellNode
                cellNode         = cellNode.next_sibling_of_type(pugi::node_element);
                // ===== Iterate over non-element nodes and mark them for deletion
                while (nextNode != cellNode)
                {    // this also works with the empty node returned past last sibling, as for XMLNode a{}, b{}, ( a == b ) is true
                    toBeDeleted.emplace_back(nextNode);
                    nextNode = nextNode.next_sibling();
                }
            }
            else
                cellNode = cellNode.next_sibling_of_type(pugi::node_element);
        }

        // ===== Delete selected cell nodes
        for (auto cellNodeToDelete : toBeDeleted) m_rowNode->remove_child(cellNodeToDelete);
    }

    /**
     * @details The prependCellValue is a convenience function used solely by the templated operator= function.
     * The purpose of a separate function is to keep details of xml_node out of the header file.
     * Note that no checking on the column number is made.
     * @pre
     * @post
     */
    void XLRowDataProxy::prependCellValue(const XLCellValue& value, uint16_t col)    // NOLINT   // 2024-04-30: whitespace support
    {
        // ===== (disabled) Pretty formatting by inserting before an existing first child
        // XMLNode first_child = m_rowNode->first_child_of_type(pugi::node_element);
        // XMLNode curNode{};
        // if (first_child.empty())
        //     curNode = m_rowNode->prepend_child("c");
        // else
        //     curNode = m_rowNode->insert_child_before("c", first_child);

        XMLNode curNode = m_rowNode->prepend_child("c");    // this will correctly insert a new cell directly at the beginning of the row
        setDefaultCellAttributes(curNode, XLCellReference(static_cast<uint32_t>(m_row->rowNumber()), col).address(), *m_rowNode, col);
        XLCell(curNode, m_row->m_sharedStrings).value() = value;
    }

    /**
     * @details
     * @pre
     * @post
     */
    void XLRowDataProxy::clear() { m_rowNode->remove_children(); }    // NOLINT

}    // namespace OpenXLSX
