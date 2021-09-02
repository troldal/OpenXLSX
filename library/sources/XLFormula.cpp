//
// Created by Kenneth Balslev on 27/08/2021.
//

// ===== OpenXLSX Includes ===== //
#include "XLFormula.hpp"
#include <pugixml.hpp>

#include <cassert>

using namespace OpenXLSX;

XLFormula::XLFormula() = default;

XLFormula::XLFormula(const XLFormula& other) = default;

XLFormula::XLFormula(XLFormula&& other) noexcept = default;

XLFormula::~XLFormula() = default;

XLFormula& XLFormula::operator=(const XLFormula& other) = default;

XLFormula& XLFormula::operator=(XLFormula&& other) noexcept = default;

std::string XLFormula::get() const
{
    return m_formulaString;
}

XLFormula& XLFormula::clear()
{
    m_formulaString = "";
    return *this;
}

XLFormulaProxy::XLFormulaProxy(XLCell* cell, XMLNode* cellNode) : m_cell(cell), m_cellNode(cellNode)
{
    assert(cell);                  // NOLINT
    //    assert(cellNode);              // NOLINT
    //    assert(!cellNode->empty());    // NOLINT
}

XLFormulaProxy::~XLFormulaProxy() = default;

XLFormulaProxy::XLFormulaProxy(const XLFormulaProxy& other) = default;

XLFormulaProxy::XLFormulaProxy(XLFormulaProxy&& other) noexcept = default;

XLFormulaProxy& XLFormulaProxy::operator=(const XLFormulaProxy& other)
{
    if (&other != this) {
        *this = other.getFormula();
    }

    return *this;
}

XLFormulaProxy& XLFormulaProxy::operator=(XLFormulaProxy&& other) noexcept = default;

XLFormulaProxy::operator XLFormula()
{
    return getFormula();
}

std::string XLFormulaProxy::get() const
{
    return getFormula().get();
}

XLFormulaProxy& XLFormulaProxy::clear()
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode);              // NOLINT
    assert(!m_cellNode->empty());    // NOLINT

    // ===== Remove the value node.
    if (m_cellNode->child("v")) m_cellNode->remove_child("v");
    return *this;
}

void XLFormulaProxy::setFormulaString(const char* formulaString) {
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode);              // NOLINT
    assert(!m_cellNode->empty());    // NOLINT

    // ===== If the cell node doesn't have a value child node, create it.
    if (!m_cellNode->child("f")) m_cellNode->append_child("f");

    // ===== Set the text of the value node.
    m_cellNode->child("f").text().set(formulaString);
}

XLFormula XLFormulaProxy::getFormula() const
{
    assert(m_cellNode);              // NOLINT
    assert(!m_cellNode->empty());    // NOLINT

    auto formulaNode = m_cellNode->child("f");

    if (!formulaNode)
        return XLFormula();
    return XLFormula(formulaNode.text().get());
}
