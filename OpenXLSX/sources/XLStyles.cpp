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
#include <pugixml.hpp>
#include <sstream>
#include <string>
#include <string_view>


// ===== OpenXLSX Includes ===== //
#include "XLCell.hpp"
#include "XLDocument.hpp"
#include "XLStyles.hpp"
#include "XLXmlData.hpp"

using namespace OpenXLSX;

// ================================================================================
// Helpers
// ================================================================================
template<typename Out>
inline void splitA(const std::string& s, char delim, Out result)
{
    std::istringstream iss(s);
    std::string        item;
    while (std::getline(iss, item, delim)) {
        *result++ = item;
    }
}

inline void splitA(const std::string& s, char delim, std::vector<std::string>& elems)
{
    splitA(s, delim, std::back_inserter(elems));
}

static pugi::xml_node findFontNode(const XMLDocument& xmlDoc, int fontId)
{
    //xmlDoc.find_child() ? how to use
    pugi::xml_node    node;
    constexpr std::string_view nodeFontsName { "fonts" };
    constexpr std::string_view nodeFontName { "font" };
    for (const auto& nodeItem : xmlDoc.document_element().children()) {
        if (nodeFontsName == nodeItem.name()) {
            for (auto& xfNode : nodeItem.children()) {
                if (nodeFontName == xfNode.name()) {
                    if (fontId == 0) {
                        return xfNode;
                    }
                    else {
                        int idx = 0;
                        while (idx < fontId) {
                            xfNode = xfNode.next_sibling();
                            idx++;
                        }
                        return xfNode;
                    }
                }
            }
        }
    }
    return node;
}

// ================================================================================
// XLStyles Class
// ================================================================================

XLStyles::XLStyles(XLXmlData* xmlData) : XLXmlFile(xmlData)
{
    init(xmlData);
}

XLStyles::~XLStyles() = default;

XLStyle XLStyles::style(const XLCell& cell) const
{
    XLStyle res(this->parentDoc());
    if (!cell) return res;
    if (const auto* val = cell.m_cellNode->attribute("s").value(); val != nullptr) {
        if (const size_t pos = std::atoi(val); m_VecStyle.size() > pos) {
            res = m_VecStyle.at(pos);
        }
    }
    return res;
}

std::string XLStyles::formatString(int numFmtId) const
{
    constexpr const auto* snumFmts    = "numFmts";
    constexpr const auto* snumFmtId   = "numFmtId";
    constexpr const auto* sformatCode = "formatCode";

    return xmlDocument()
        .document_element()
        .child(snumFmts)
        .find_child_by_attribute(snumFmtId, std::to_string(numFmtId).c_str())
        .attribute(sformatCode)
        .value();
}

void XLStyles::init(const XLXmlData* stylesData)
{
    // cache cellXfs
    if (stylesData) {
        constexpr const auto* numFmtId = "numFmtId";
        constexpr const auto* fontId   = "fontId";
        constexpr const auto* fillId   = "fillId";
        constexpr const auto* borderId = "borderId";
        constexpr const auto* xfId     = "xfId";

        constexpr std::string_view nodeNameCellXfs { "cellXfs" };
        constexpr std::string_view nodeNameXf { "xf" };

        // TODO: check applyNumberFormat, reserve size from cellXfs count;
        for (const auto& node : stylesData->getXmlDocument()->document_element().children()) {
            if (nodeNameCellXfs == node.name()) {
                for (const auto& xfNode : node.children()) {
                    if (nodeNameXf == xfNode.name()) {
                        XLStyle item(this->parentDoc());
                        if (const auto* val = xfNode.attribute(numFmtId).value(); val != nullptr) item.m_numFmtId = std::stoi(val);
                        if (const auto* val = xfNode.attribute(fontId).value(); val != nullptr) item.m_fontId = std::stoi(val);
                        if (const auto* val = xfNode.attribute(fillId).value(); val != nullptr) item.m_fillId = std::stoi(val);
                        if (const auto* val = xfNode.attribute(borderId).value(); val != nullptr) item.m_borderId = std::stoi(val);
                        if (const auto* val = xfNode.attribute(xfId).value(); val != nullptr) item.m_xfId = std::stoi(val);
                        m_VecStyle.emplace_back(item);
                    }
                }
                break;
            }
            // TODO: Fonts
        }
    }
}

// ================================================================================
// XLStyle Class
// ================================================================================

XLStyle::XLStyle(const XLDocument& doc) : m_doc(doc) {}

XLNumberFormat XLStyle::numberFormat() const
{
    XLNumberFormat fmt(*this);
    return fmt;
}

std::string XLStyle::formatString() const
{
    return m_doc.get().styles().formatString(m_numFmtId);
}

int XLStyle::numFmtId() const
{
    return m_numFmtId;
}

XLFont XLStyle::font() const
{
    return XLFont(*this, findFontNode(m_doc.get().styles().xmlDocument(), m_fontId));
}

// ================================================================================
// XLFont Class
// ================================================================================
// XLQuery?
XLFont::XLFont(const XLStyle& style, const XMLNode& node) : 
    m_style(style), m_node(new XMLNode(node)) {}

XLFont::~XLFont() {}

std::string XLFont::name() const
{
    std::string _name;
    if (isValid()) {
        _name = m_node->child("name").attribute("val").value();
    }
    return _name;
}

double XLFont::size() const
{
    double result = 0.0;
    if (isValid()) {
        result = std::stod(m_node->child("sz").attribute("val").value());
    }
    return result;
}

XLColor XLFont::color() const
{
    XLColor clr;
    if (isValid()) {
        const std::string val = m_node->child("color").attribute("rgb").value();
        if (val.size() != 0) {
            clr.set(val);
        }
    }
    return clr;
}

int XLFont::colorIndex() const
{
    constexpr int xlColorIndexNone = -4142;
    if (isValid()) {
        const std::string val = m_node->child("color").attribute("theme").value();
        if (val.size() != 0) {
            return std::atoi(val.c_str());
        }
    }
    return xlColorIndexNone;
}


int XLFont::underline() const {
    //https://docs.microsoft.com/en-us/office/vba/api/excel.xlunderlinestyle
    if (isValid()) {
        constexpr std::string_view nodeName { "u" };
        constexpr std::string_view attDouble { "double" };
        for (const auto& node : m_node->children()) {

            if (nodeName == node.name()) {
                if (auto attribute = node.attribute("val"); !attribute.empty() && attribute.value() == attDouble) {
                    return -4119;
                }
                else {
                    return 2;
                }

            }

        }
    }
    return -4142;
}

bool XLFont::strikethrough() const {
    if (isValid()) {
        constexpr std::string_view nodeName { "strike" };
        for (const auto& node : m_node->children()) {
            if (nodeName == node.name()) {
                return true;
            }
        }
    }
    return false;
}


bool XLFont::bold() const {
    if (isValid()) {
        constexpr std::string_view nodeName { "b" };
        for (const auto& node : m_node->children()) {
            if (nodeName == node.name()) {
                return true;
            }
        }
    }
    return false;
}


bool XLFont::italic() const{
    if (isValid()) {
        constexpr std::string_view nodeName { "i" };
        for (const auto& node : m_node->children()) {
            if (nodeName == node.name()) {
                return true;
            }
        }
    }
    return false;
}

bool XLFont::isValid() const 
{
    constexpr std::string_view nodeFontName { "font" };
    if (m_style.get().m_fontId != -1 && m_node != nullptr && m_node->type() != pugi::xml_node_type::node_null && m_node->name() == nodeFontName){
        return true;
    }
    return false;
}

// ================================================================================
// XLNumberFormat Class
// ================================================================================
XLNumberFormat::XLNumberFormat(const XLStyle& style) : m_style { style } {}

XLNumberFormat::XLNumberType XLNumberFormat::type()
{
    return tryFindType();
}

std::string XLNumberFormat::currencySymbol() const
{
    return m_currencySymbol;
}

// TODO: it may be possible to get precision and forrmatting in a single pass.
XLNumberFormat::XLNumberType XLNumberFormat::tryFindType()
{
    const auto _type = tryBuiltinType();
    if (_type != XLNumberType::kUnkown) return _type;

    const std::string fmt = m_style.get().formatString();
    if (fmt.size() == 0) return XLNumberType::kUnkown;

    std::vector<std::string> bracketTextArray;
    std::lconv*              lc = std::localeconv();
    std::string              bracketText;
    bool                     isInBracket = false;

    auto foundType = XLNumberType::kUnkown;

    if (fmt.find("AM/PM") != std::string::npos) return XLNumberType::kDate;

    for (size_t idx = 0; idx < fmt.size(); idx++) {
        if (fmt[idx] == '[') {
            isInBracket = true;
        }
        else if (fmt[idx] == ']') {
            isInBracket = false;
            bracketTextArray.push_back(bracketText);
            bracketText.clear();
        }
        else if (isInBracket) {
            bracketText += fmt[idx];
        }
        else {
            const char chr = fmt[idx];
            if (chr == '$') {
                m_currencySymbol = chr;
                return XLNumberType::kCurrency;
            }
            else if (std::string lc_curr = lc->currency_symbol; lc_curr.find(chr) != std::string::npos) {
                m_currencySymbol = lc->currency_symbol;
                return XLNumberType::kCurrency;
            }
            else if (chr == '%') {
                return XLNumberType::kPercent;
            }
            else if (chr == 'y' || chr == 'M' || chr == 'm' || chr == 'd') {
                return XLNumberType::kDate;
            }
            else if (chr == 'h' || chr == 's' || chr == ':') {
                return XLNumberType::kDate;
            }
            else if (chr == '0' || chr == '#' || chr == '.') {
                foundType = XLNumberType::kUnkown;
            }
            // use lc->int_curr_symbol ? how to get en_US?
            //(iCompare(m_slocale, L"ja_jp") || iCompare(m_slocale, L"zh_tw")) && wchr == 'e' || wchr == 'g' || wchr == 'r')
        }
    }
    for (const auto& item : bracketTextArray) {
        if (item.find("F800") != std::string::npos) {
            return XLNumberType::kDate;
        }
        else if (item.find("F400") != std::string::npos) {
            return XLNumberType::kDate;
        }
        else {
            if (size_t pos = item.find('-'); pos != std::wstring::npos) {
                std::vector<std::string> parts;
                splitA(item, '-', parts);
                if (parts.size() == 2) {
                    if (parts[0].size() > 1) {
                        m_currencySymbol = parts[0].substr(1, parts[0].size() - 1);
                        m_fmtLocal       = parts[1];
                        return XLNumberType::kCurrency;
                    }
                }
            }
            else {
                if (item[0] == '$') {
                    m_currencySymbol = item.substr(1, item.size() - 1);
                    return XLNumberType::kCurrency;
                }
            }
        }
    }
    return foundType;
}

XLNumberFormat::XLNumberType XLNumberFormat::tryBuiltinType()
{
    std::setlocale(LC_MONETARY, "");
    std::lconv* lc = std::localeconv();
    auto n = m_style.get().numFmtId();
    switch (m_style.get().numFmtId()) {
        case 0:
        case 1:
        case 3:
            return XLNumberType::kUnkown;
        case 2:
        case 4:
            return XLNumberType::kUnkown;
        case 5:
        case 6:
        case 7:
        case 8:
            m_currencySymbol = lc->currency_symbol;
            return XLNumberType::kCurrency;
        case 9:
        case 10:
            return XLNumberType::kPercent;
        case 11:
        case 12:
        case 13:
            return XLNumberType::kUnkown;
        case 14:
        case 15:
        case 16:
        case 17:
        case 18:
        case 19:
        case 20:
        case 21:
        case 22:
            return XLNumberType::kDate;
        case 37:
        case 38:
        case 39:
        case 40:    // accounting
            return XLNumberType::kUnkown;
        case 45:
        case 46:
        case 47:
        case 48:
            return XLNumberType::kDate;
    }
    return XLNumberType::kUnkown;
}