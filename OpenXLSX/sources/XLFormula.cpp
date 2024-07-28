//
// Created by Kenneth Balslev on 27/08/2021.
//

// ===== External Includes ===== //
#include <cassert>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLFormula.hpp"
#include <XLException.hpp>

using namespace OpenXLSX;

/**
 * @details Constructor. Default implementation.
 */
XLFormula::XLFormula() = default;

/**
 * @details Copy constructor. Default implementation.
 */
XLFormula::XLFormula(const XLFormula& other) = default;

/**
 * @details Move constructor. Default implementation.
 */
XLFormula::XLFormula(XLFormula&& other) noexcept = default;

/**
 * @details Destructor. Default implementation.
 */
XLFormula::~XLFormula() = default;

/**
 * @details Copy assignment operator. Default implementation.
 */
XLFormula& XLFormula::operator=(const XLFormula& other) = default;

/**
 * @details Move assignment operator. Default implementation.
 */
XLFormula& XLFormula::operator=(XLFormula&& other) noexcept = default;

/**
 * @details Return the m_formulaString member variable.
 */
std::string XLFormula::get() const { return m_formulaString; }

/**
 * @details Set the m_formulaString member to an empty string.
 */
XLFormula& XLFormula::clear()
{
    m_formulaString = "";
    return *this;
}

/**
 * @details
 */
XLFormula::operator std::string() const { return get(); }

/**
 * @details Constructor. Set the m_cell and m_cellNode objects.
 */
XLFormulaProxy::XLFormulaProxy(XLCell* cell, XMLNode* cellNode) : m_cell(cell), m_cellNode(cellNode)
{
    assert(cell);    // NOLINT
}

/**
 * @details Destructor. Default implementation.
 */
XLFormulaProxy::~XLFormulaProxy() = default;

/**
 * @details Copy constructor. default implementation.
 */
XLFormulaProxy::XLFormulaProxy(const XLFormulaProxy& other) = default;

/**
 * @details Move constructor. Default implementation.
 */
XLFormulaProxy::XLFormulaProxy(XLFormulaProxy&& other) noexcept = default;

/**
 * @details Calls the templated string assignment operator.
 */
XLFormulaProxy& XLFormulaProxy::operator=(const XLFormulaProxy& other)
{
    if (&other != this) {
        *this = other.getFormula();
    }

    return *this;
}

/**
 * @details Move assignment operator. Default implementation.
 */
XLFormulaProxy& XLFormulaProxy::operator=(XLFormulaProxy&& other) noexcept = default;

/**
 * @details
 */
XLFormulaProxy::operator std::string() const { return get(); }

/**
 * @details Returns the underlying XLFormula object, by calling getFormula().
 */
XLFormulaProxy::operator XLFormula() const { return getFormula(); }

/**
 * @details Call the .get() function in the underlying XLFormula object.
 */
std::string XLFormulaProxy::get() const { return getFormula().get(); }

/**
 * @details If a formula node exists, it will be erased.
 */
XLFormulaProxy& XLFormulaProxy::clear()
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    // ===== Remove the value node.
    if (not m_cellNode->child("f").empty()) m_cellNode->remove_child("f");
    return *this;
}

/**
 * @details Convenience function for setting the formula. This method is called from the templated
 * string assignment operator.
 */
void XLFormulaProxy::setFormulaString(const char* formulaString, bool resetValue) // NOLINT
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    // ===== If the cell node doesn't have a value child node, create it.
    if (m_cellNode->child("f").empty()) m_cellNode->append_child("f");
    if (m_cellNode->child("v").empty()) m_cellNode->append_child("v");

    // ===== Remove the formula type and shared index attributes, if they exist.
    m_cellNode->child("f").remove_attribute("t");
    m_cellNode->child("f").remove_attribute("si");

    // ===== Set the text of the value node.
    m_cellNode->child("f").text().set(formulaString);
    if (resetValue) m_cellNode->child("v").text().set(0);

    // BEGIN pull request #189
    // ===== Remove cell type attribute so that it can be determined by Office Suite when next calculating the formula.
    m_cellNode->remove_attribute("t");

    // ===== Remove inline string <is> tag, in case previous type was "inlineStr".
    m_cellNode->remove_child("is");

    // ===== Ensure that the formula node <f> is the first child, listed before the value <v> node.
    m_cellNode->prepend_move(m_cellNode->child("f"));
    // END pull request #189
}

/**
 * @details Creates and returns an XLFormula object, based on the formula string in the underlying
 * XML document.
 */
XLFormula XLFormulaProxy::getFormula() const
{
    assert(m_cellNode != nullptr);      // NOLINT
    assert(not m_cellNode->empty());    // NOLINT

    const auto formulaNode = m_cellNode->child("f");

    // ===== If the formula node doesn't exist, return an empty XLFormula object.
    if (formulaNode.empty()) return XLFormula();

    // ===== If the formula type is 'shared' or 'array', throw an exception.
    if (not formulaNode.attribute("t").empty() ) {    // 2024-05-28: de-duplicated check (only relevant for performance,
                                                      //  xml_attribute::value() returns an empty string for empty attributes)
        if (std::string(formulaNode.attribute("t").value()) == "shared")
            throw XLFormulaError("Shared formulas not supported.");
        if (std::string(formulaNode.attribute("t").value()) == "array")
            throw XLFormulaError("Array formulas not supported.");
    }

    return XLFormula(formulaNode.text().get());
}
