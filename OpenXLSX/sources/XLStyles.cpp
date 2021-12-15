// ===== External Includes ===== //
#include <algorithm>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLStyles.hpp"
#include "XLXmlData.hpp"
#include "XLCell.hpp"
#include "XLDocument.hpp"

using namespace OpenXLSX;


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

// DRM
XLStyles::~XLStyles() = default;

XLStyles::XLStyles(XLXmlData* xmlData) 
    : XLXmlFile(xmlData) {
    init(xmlData);
}



OpenXLSX::XLStyle OpenXLSX::XLStyles::style(const XLCell& cell) const
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

std::string OpenXLSX::XLStyles::formatString(int numFmtId) const
{
    constexpr const auto* snumFmts = "numFmts";
    constexpr const auto* snumFmtId = "numFmtId";
    constexpr const auto* sformatCode = "formatCode";

    return xmlDocument()
        .document_element()
        .child(snumFmts)
        .find_child_by_attribute(snumFmtId, std::to_string(numFmtId).c_str())
        .attribute(sformatCode)
        .value();
}

void OpenXLSX::XLStyles::init(const XLXmlData* stylesData)
{
    //cache cellXfs
    if (stylesData) {
        constexpr const auto* numFmtId   = "numFmtId";
        constexpr const auto* fontId     = "fontId";
        constexpr const auto* fillId     = "fillId";
        constexpr const auto* borderId   = "borderId";
        constexpr const auto* xfId       = "xfId";

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

//
OpenXLSX::XLStyle::XLStyle(const XLDocument& doc) 
    : m_doc( doc ) 
{
}

OpenXLSX::XLNumberFormat OpenXLSX::XLStyle::numberFormat() const { 
    XLNumberFormat fmt(*this);
    return fmt;
}

std::string OpenXLSX::XLStyle::formatString() const {
    return m_doc.get().styles().formatString(m_numFmtId);
}

int OpenXLSX::XLStyle::numFmtId() const {
    return m_numFmtId;
}

//
OpenXLSX::XLNumberFormat::XLNumberFormat(const XLStyle& style) 
    : m_style { style }
{
}

OpenXLSX::XLNumberFormat::XLNumberType OpenXLSX::XLNumberFormat::type() {
    return tryFindType();
}


std::string OpenXLSX::XLNumberFormat::currencySumbol() const{
    return m_currencySumbol;
}

OpenXLSX::XLNumberFormat::XLNumberType OpenXLSX::XLNumberFormat::tryFindType()
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
        else 
        {
            const char chr = fmt[idx];
            if (chr == lc->currency_symbol[0] || chr == '$') {
                m_currencySumbol = chr;
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
            //(iCompare(m_slocale, L"ja-jp") || iCompare(m_slocale, L"zh-tw")) && wchr == 'e' || wchr == 'g' || wchr == 'r')
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
                        m_currencySumbol = parts[0].substr(1, parts[0].size() - 1);
                        m_fmtLocal       = parts[1];
                        return XLNumberType::kCurrency;
                    }
                }
            }
            else {
                if (item[0] == '$') {
                    m_currencySumbol = item.substr(1, item.size() - 1);
                    return XLNumberType::kCurrency;
                }
            }
            break;
        }
    }
    return foundType;
}

OpenXLSX::XLNumberFormat::XLNumberType OpenXLSX::XLNumberFormat::tryBuiltinType()
{
    std::setlocale(LC_MONETARY,"");
    std::lconv* lc = std::localeconv();
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
            m_currencySumbol = lc->currency_symbol[0];
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
        case 40: //accounting 
            return XLNumberType::kUnkown;
        case 45: 
        case 46: 
        case 47: 
        case 48: 
            return XLNumberType::kDate;
    }
    return XLNumberType::kUnkown;
}
