//
// Created by Kenneth Balslev on 27/08/2021.
//

// ===== OpenXLSX Includes ===== //
#include "XLFormula.hpp"

using namespace OpenXLSX;

XLFormula::XLFormula() = default;

XLFormula::XLFormula(const XLFormula& other) = default;

XLFormula::XLFormula(XLFormula&& other) noexcept = default;

XLFormula::~XLFormula() = default;

XLFormula& XLFormula::operator=(const XLFormula& other) = default;

XLFormula& XLFormula::operator=(XLFormula&& other) noexcept = default;

const std::string& XLFormula::get() const
{
    return m_formulaString;
}

XLFormula& XLFormula::clear()
{
    m_formulaString = "";
    return *this;
}
