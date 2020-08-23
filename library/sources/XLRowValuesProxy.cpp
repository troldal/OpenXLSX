//
// Created by Kenneth Balslev on 23/08/2020.
//

// ===== External Includes ===== //
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLRowValuesProxy.hpp"
#include "XLSharedStrings.hpp"
#include "XLCellReference.hpp"
#include "XLCell.hpp"
#include "XLRow.hpp"

using namespace OpenXLSX;

namespace
{
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

/**
 * @details
 * @pre
 * @post
 */
XLRowValuesProxy::~XLRowValuesProxy() = default;

/**
 * @details
 * @pre
 * @post
 */
XLRowValuesProxy& XLRowValuesProxy::operator=(const XLRowValuesProxy& other)
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
XLRowValuesProxy::XLRowValuesProxy(XLRow* row, XMLNode* rowNode) : m_row(row), m_rowNode(rowNode){}

/**
 * @details
 * @pre
 * @post
 */
XLRowValuesProxy::XLRowValuesProxy(const XLRowValuesProxy& other) = default;

/**
 * @details
 * @pre
 * @post
 */
XLRowValuesProxy::XLRowValuesProxy(XLRowValuesProxy&& other) noexcept = default;

/**
 * @details
 * @pre
 * @post
 */
XLRowValuesProxy& XLRowValuesProxy::operator=(XLRowValuesProxy&& other) noexcept = default;

/**
 * @details
 * @pre
 * @post
 */
XLRowValuesProxy& XLRowValuesProxy::operator=(const std::vector<XLCellValue>& values)
{
    setValues(values, m_row->rowNumber(), m_rowNode, m_row->m_sharedStrings);
    return *this;
}

/**
 * @details
 * @pre
 * @post
 */
XLRowValuesProxy::operator std::vector<XLCellValue>()
{
    return getValues();
}

/**
 * @details
 * @pre
 * @post
 */
std::vector<XLCellValue> XLRowValuesProxy::getValues() const
{
    auto numCells = XLCellReference(m_rowNode->last_child().attribute("r").value()).column();
    std::vector<XLCellValue> result(numCells);

    for (auto& node : m_rowNode->children()) {
        if (XLCellReference(node.attribute("r").value()).column() > numCells) break;
        result[XLCellReference(node.attribute("r").value()).column() - 1] = XLCell(node, m_row->m_sharedStrings).value();
    }

    return result;
}
