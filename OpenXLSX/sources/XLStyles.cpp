// ===== External Includes ===== //
#include <algorithm>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLStyles.hpp"
#include "XLXmlData.hpp"
#include "XLCell.hpp"

using namespace OpenXLSX;

// DRM
XLStyles::~XLStyles() = default;

XLStyles::XLStyles(XLXmlData* xmlData) 
    : XLXmlFile(xmlData) {
    init(xmlData);
}

size_t OpenXLSX::XLStyles::cellXfsSize() const {
    return m_VecXLCellXfs.size();
}

OpenXLSX::XLCellXfs OpenXLSX::XLStyles::cellXfsFor(const XLCell& cell) const
{
    XLCellXfs res;
    if (!cell) return res;
    if (const auto* val = cell.m_cellNode->attribute("s").value(); val != nullptr) {
        if (const size_t pos = std::atoi(val); m_VecXLCellXfs.size() > pos) {
            return m_VecXLCellXfs.at(pos);
        }
    }
    return XLCellXfs();
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
                        XLCellXfs item;
                        if (const auto* val = xfNode.attribute(numFmtId).value(); val != nullptr) item.numFmtId = std::stoi(val);
                        if (const auto* val = xfNode.attribute(fontId).value(); val != nullptr) item.fontId = std::stoi(val);
                        if (const auto* val = xfNode.attribute(fillId).value(); val != nullptr) item.fillId = std::stoi(val);
                        if (const auto* val = xfNode.attribute(borderId).value(); val != nullptr) item.borderId = std::stoi(val);
                        if (const auto* val = xfNode.attribute(xfId).value(); val != nullptr) item.xfId = std::stoi(val);
                        m_VecXLCellXfs.push_back(item);
                    }
                }
                break;
            }
            // TODO: Fonts
        }
    }
}
