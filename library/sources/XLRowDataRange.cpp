//
// Created by Kenneth Balslev on 24/08/2020.
//

#include "XLRowDataRange.hpp"

#include <pugixml.hpp>
using namespace OpenXLSX;

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
        *m_rowNode     = *other.m_rowNode;
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
        *m_rowNode     = *other.m_rowNode;
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
