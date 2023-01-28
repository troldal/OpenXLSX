//
// Created by Kenneth Balslev on 27/08/2021.
//

// ===== External Includes ===== //
#include <regex>
#include <cassert>

// ===== OpenXLSX Includes ===== //
#include "XLFormula.hpp"
#include <pugixml.hpp>

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
std::string XLFormula::get() const
{
    return m_formulaString;
}

/**
 * @details Set the m_formulaString member to an empty string.
 */
XLFormula& XLFormula::clear()
{
    m_formulaString = "";
    m_isError = false;
    return *this;
}

/**
 * @details use a regex to replace the variable to be deleted by #REF!
 */
XLFormula XLFormula::updateDeleting(const std::string& toBeDeleted)
{
    
    //std::string s = toBeDeleted + R"aka((?=([^"]*"[^"]*")*[^"]*$))aka";
    // Regex that will find all the occurence of toBeDeleted excepted inside quote
    // TODO deal with nested quotes 
    const std::regex re(toBeDeleted + R"&((?=([^"]*"[^"]*")*[^"]*$))&");

    std::string newFormula = std::regex_replace(m_formulaString, re, "#REF!");

    if(m_formulaString == newFormula)
        m_isError = false;
    else
        m_isError = true;

    m_formulaString = newFormula;

    return *this;
}

bool XLFormula::hasError() const
{
    return m_isError;
}

void XLFormula::checkIfError()
{
    //Check if Formula has error
    // TODO deal with nested quotes
    const std::regex re(R"&(#REF!(?=([^"]*"[^"]*")*[^"]*$))&");

    std::smatch m; // unused

    if (std::regex_search(m_formulaString, m, re))
        m_isError = true;

}

/**
 * @details
 */
XLFormula::operator std::string() const
{
    return get();
}

/**
 * @details Constructor. Set the m_cell and m_cellNode objects.
 */
XLFormulaProxy::XLFormulaProxy(XLCell* cell, std::shared_ptr<XMLNode> cellNode) : m_cell(cell), m_cellNode(cellNode)
{
    assert(cell); // NOLINT
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
XLFormulaProxy::operator std::string() const
{
    return get();
}

/**
 * @details Returns the underlying XLFormula object, by calling getFormula().
 */
XLFormulaProxy::operator XLFormula() const
{
    return getFormula();
}

/**
 * @details Call the .get() function in the underlying XLFormula object.
 */
std::string XLFormulaProxy::get() const
{
    return getFormula().get();
}

/**
 * @details If a formula node exists, it will be erased.
 */
XLFormulaProxy& XLFormulaProxy::clear()
{
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode);              // NOLINT
    assert(!m_cellNode->empty());    // NOLINT

    // ===== Remove the value node.
    if (m_cellNode->child("f")) m_cellNode->remove_child("f");
    return *this;
}

/**
 * @details Set the m_formulaString member to an empty string.
 */
XLFormula XLFormulaProxy::updateDeleting(const std::string& toBeDeleted)
{ 
    return getFormula().updateDeleting(toBeDeleted);
}

bool XLFormulaProxy::hasError() const
{
     return getFormula().hasError();
}

/**
 * @details Convenience function for setting the formula. This method is called from the templated
 * string assignment operator.
 */
void XLFormulaProxy::setFormulaString(const char* formulaString) {
    // ===== Check that the m_cellNode is valid.
    assert(m_cellNode);              // NOLINT
    assert(!m_cellNode->empty());    // NOLINT

    // ===== If the cell node doesn't have a value child node, create it.
    if (!m_cellNode->child("f")) m_cellNode->append_child("f");
    if (!m_cellNode->child("v")) m_cellNode->append_child("v");

    // ===== Remove the type and shared index attributes, if they exists.
    m_cellNode->child("f").remove_attribute("t");
    m_cellNode->child("f").remove_attribute("si");

    // ===== Set the text of the value node.
    m_cellNode->child("f").text().set(formulaString);
    m_cellNode->child("v").text().set(0);

    // ===== Remove "t" attribute so that the cell shows the number we've just set
    m_cellNode->remove_attribute("t");

    // ===== Remove <is> tag in case previous type was "inlineStr"
    m_cellNode->remove_child("is");   

    // ===== Excel fails to load documents where <f> comes after <v> so make sure it is the first child    
    m_cellNode->prepend_move(m_cellNode->child("f"));  
}

/**
 * @details Creates and returns an XLFormula object, based on the formula string in the underlying
 * XML document.
 */
XLFormula XLFormulaProxy::getFormula() const
{
    assert(m_cellNode);              // NOLINT
    assert(!m_cellNode->empty());    // NOLINT

    auto formulaNode = m_cellNode->child("f");

    // ===== If the formula node doesn't exist, return an empty XLFormula object.
    if (!formulaNode)
        return XLFormula();

    // ===== If the formula type is 'shared' or 'array', throw an exception.
    if (formulaNode.attribute("t") && std::string(formulaNode.attribute("t").value()) == "shared")
        throw XLFormulaError("Shared formulas not supported.");
    if (formulaNode.attribute("t") && std::string(formulaNode.attribute("t").value()) == "array")
        throw XLFormulaError("Array formulas not supported.");

    return XLFormula(formulaNode.text().get());
}
