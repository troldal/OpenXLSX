//
// Created by Kenneth Balslev on 22/08/2020.
//

#include "XLRowRange.hpp"
#include <pugixml.hpp>

using namespace OpenXLSX;

/**
 * @details
 * @pre
 * @post
 */
XLRowRange::XLRowRange(const XMLNode& dataNode, uint32_t first, uint32_t last, OpenXLSX::XLSharedStrings* sharedStrings)
    : m_dataNode(std::make_unique<XMLNode>(dataNode)),
      m_firstRow(first),
      m_lastRow(last),
      m_sharedStrings(sharedStrings)
{}

/**
 * @details
 * @pre
 * @post
 */
XLRowRange::XLRowRange(const XLRowRange& other)
    : m_dataNode(std::make_unique<XMLNode>(*other.m_dataNode)),
      m_firstRow(other.m_firstRow),
      m_lastRow(other.m_lastRow),
      m_sharedStrings(other.m_sharedStrings)
{}

/**
 * @details
 * @pre
 * @post
 */
XLRowRange::XLRowRange(XLRowRange&& other) noexcept = default;

/**
 * @details
 * @pre
 * @post
 */
XLRowRange::~XLRowRange() = default;

/**
 * @details
 * @pre
 * @post
 */
XLRowRange& XLRowRange::operator=(const XLRowRange& other)
{
    if (&other != this) {
        *m_dataNode     = *other.m_dataNode;
        m_firstRow      = other.m_firstRow;
        m_lastRow       = other.m_lastRow;
        m_sharedStrings = other.m_sharedStrings;
    }

    return *this;
}

/**
 * @details
 * @pre
 * @post
 */
XLRowRange& XLRowRange::operator=(XLRowRange&& other) noexcept
{
    if (&other != this) {
        *m_dataNode     = *other.m_dataNode;
        m_firstRow      = other.m_firstRow;
        m_lastRow       = other.m_lastRow;
        m_sharedStrings = other.m_sharedStrings;
    }

    return *this;
}

/**
 * @details
 * @pre
 * @post
 */
uint32_t XLRowRange::rowCount() const
{
    return m_lastRow - m_firstRow + 1;
}

/**
 * @details
 * @pre
 * @post
 */
XLRowIterator XLRowRange::begin()
{
    return XLRowIterator(*this, XLIteratorLocation::Begin);
}

/**
 * @details
 * @pre
 * @post
 */
XLRowIterator XLRowRange::end()
{
    return XLRowIterator(*this, XLIteratorLocation::End);
}

/**
 * @details
 * @pre
 * @post
 */
void XLRowRange::clear() {}
