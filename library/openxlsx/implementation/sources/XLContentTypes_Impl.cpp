#include <utility>

//
// Created by Troldal on 13/08/16.
//

#include <cstring>

#include "XLContentTypes_Impl.h"
#include "XLDocument_Impl.h"

using namespace std;
using namespace OpenXLSX;

/**
 * @details
 */
Impl::XLContentItem::XLContentItem(XMLNode node)
        : m_contentNode(node) {
}

/**
 * @details
 */
Impl::XLContentType Impl::XLContentItem::Type() const {

    return XLContentItem::GetTypeFromString(m_contentNode.attribute("ContentType").value());
}

/**
 * @details
 */
const string Impl::XLContentItem::Path() const {

    return m_contentNode.attribute("PartName").value();
}

/**
 * @details
 */
Impl::XLContentTypes::XLContentTypes(XLDocument& parent, const string& filePath)
        : XLAbstractXMLFile(parent, filePath),
          m_defaults(),
          m_overrides() {

    ParseXMLData();
}

/**
 * @details
 */
bool Impl::XLContentTypes::ParseXMLData() {

    m_defaults.clear();
    m_overrides.clear();

    std::string strOverride = "Override";
    std::string strDefault = "Default";

    for (auto& node : XmlDocument()->first_child().children()) {
        if (string(node.name()) == "Default") {
            std::string extension;
            std::string contentType;

            if (node.attribute("Extension") != nullptr)
                extension = node.attribute("Extension").value();
            if (node.attribute("ContentType") != nullptr)
                contentType = node.attribute("ContentType").value();
            m_defaults.insert({extension, node});
        }
        else if (string(node.name()) == "Override") {
            string path = node.attribute("PartName").value();
            XLContentType type = XLContentItem::GetTypeFromString(node.attribute("ContentType").value());

            m_overrides.emplace(path, XLContentItem(node));
        }
    }

    return true;
}

/**
 * @details
 */
void Impl::XLContentTypes::AddDefault(const string& key, XMLNode node) {

    m_defaults.emplace(key, node);
}

/**
 * @details
 */
void Impl::XLContentTypes::AddOverride(const string& path, XLContentType type) {

    string typeString = XLContentItem::GetStringFromType(type);

    auto node = XmlDocument()->first_child().append_child("Override");
    node.append_attribute("PartName").set_value(path.c_str());
    node.append_attribute("ContentType").set_value(typeString.c_str());

    m_overrides.emplace(path, XLContentItem(node));
}

/**
 * @details
 */
void Impl::XLContentTypes::DeleteOverride(XLContentItem& item) {

    m_overrides.erase(item.Path());
    XmlDocument()->first_child().remove_child(XmlDocument()->first_child().find_child_by_attribute("PartName",
                                                                                                   item.Path()
                                                                                                       .c_str()));

}

/**
 * @details
 */
Impl::XLContentItem Impl::XLContentTypes::ContentItem(const std::string& path) {

    return m_overrides.at(path);
}

/**
 * @details
 */
Impl::XLContentType Impl::XLContentItem::GetTypeFromString(const std::string& typeString) {

    XLContentType type;

    if (typeString == "application/vnd.ms-excel.Sheet.macroEnabled.main+xml")
        type = XLContentType::WorkbookMacroEnabled;
    else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.Sheet.main+xml")
        type = XLContentType::Workbook;
    else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.Worksheet+xml")
        type = XLContentType::Worksheet;
    else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.Chartsheet+xml")
        type = XLContentType::Chartsheet;
    else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml")
        type = XLContentType::ExternalLink;
    else if (typeString == "application/vnd.openxmlformats-officedocument.theme+xml")
        type = XLContentType::Theme;
    else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
        type = XLContentType::Styles;
    else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.SharedStrings+xml")
        type = XLContentType::SharedStrings;
    else if (typeString == "application/vnd.openxmlformats-officedocument.drawing+xml")
        type = XLContentType::Drawing;
    else if (typeString == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
        type = XLContentType::Chart;
    else if (typeString == "application/vnd.ms-office.chartstyle+xml")
        type = XLContentType::ChartStyle;
    else if (typeString == "application/vnd.ms-office.chartcolorstyle+xml")
        type = XLContentType::ChartColorStyle;
    else if (typeString == "application/vnd.ms-excel.controlproperties+xml")
        type = XLContentType::ControlProperties;
    else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml")
        type = XLContentType::CalculationChain;
    else if (typeString == "application/vnd.ms-office.vbaProject")
        type = XLContentType::VBAProject;
    else if (typeString == "application/vnd.openxmlformats-package.core-properties+xml")
        type = XLContentType::CoreProperties;
    else if (typeString == "application/vnd.openxmlformats-officedocument.extended-properties+xml")
        type = XLContentType::ExtendedProperties;
    else if (typeString == "application/vnd.openxmlformats-officedocument.custom-properties+xml")
        type = XLContentType::CustomProperties;
    else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
        type = XLContentType::Comments;
    else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
        type = XLContentType::Table;
    else if (typeString == "application/vnd.openxmlformats-officedocument.vmlDrawing")
        type = XLContentType::VMLDrawing;
    else
        type = XLContentType::Unknown;

    return type;
}

/**
 * @details
 */
std::string Impl::XLContentItem::GetStringFromType(Impl::XLContentType type) {

    string typeString;

    if (type == XLContentType::WorkbookMacroEnabled)
        typeString = "application/vnd.ms-excel.Sheet.macroEnabled.main+xml";
    else if (type == XLContentType::Workbook)
        typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.Sheet.main+xml";
    else if (type == XLContentType::Worksheet)
        typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.Worksheet+xml";
    else if (type == XLContentType::Chartsheet)
        typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.Chartsheet+xml";
    else if (type == XLContentType::ExternalLink)
        typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
    else if (type == XLContentType::Theme)
        typeString = "application/vnd.openxmlformats-officedocument.theme+xml";
    else if (type == XLContentType::Styles)
        typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
    else if (type == XLContentType::SharedStrings)
        typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.SharedStrings+xml";
    else if (type == XLContentType::Drawing)
        typeString = "application/vnd.openxmlformats-officedocument.drawing+xml";
    else if (type == XLContentType::Chart)
        typeString = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    else if (type == XLContentType::ChartStyle)
        typeString = "application/vnd.ms-office.chartstyle+xml";
    else if (type == XLContentType::ChartColorStyle)
        typeString = "application/vnd.ms-office.chartcolorstyle+xml";
    else if (type == XLContentType::ControlProperties)
        typeString = "application/vnd.ms-excel.controlproperties+xml";
    else if (type == XLContentType::CalculationChain)
        typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
    else if (type == XLContentType::VBAProject)
        typeString = "application/vnd.ms-office.vbaProject";
    else if (type == XLContentType::CoreProperties)
        typeString = "application/vnd.openxmlformats-package.core-properties+xml";
    else if (type == XLContentType::ExtendedProperties)
        typeString = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
    else if (type == XLContentType::CustomProperties)
        typeString = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
    else if (type == XLContentType::Comments)
        typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
    else if (type == XLContentType::Table)
        typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
    else if (type == XLContentType::VMLDrawing)
        typeString = "application/vnd.openxmlformats-officedocument.vmlDrawing";
    else
        throw XLException("Unknown ContentType");

    return typeString;
}
