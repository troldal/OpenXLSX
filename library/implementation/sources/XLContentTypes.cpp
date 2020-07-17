#include <utility>

//
// Created by Troldal on 13/08/16.
//

#include "XLContentTypes.hpp"
#include "XLDocument.hpp"

#include <cstring>

using namespace std;
using namespace OpenXLSX;

namespace
{
    /**
     * @details
     */
    Impl::XLContentType GetTypeFromString(const std::string& typeString)
    {
        Impl::XLContentType type;

        if (typeString == "application/vnd.ms-excel.Sheet.macroEnabled.main+xml")
            type = Impl::XLContentType::WorkbookMacroEnabled;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
            type = Impl::XLContentType::Workbook;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
            type = Impl::XLContentType::Worksheet;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml")
            type = Impl::XLContentType::Chartsheet;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml")
            type = Impl::XLContentType::ExternalLink;
        else if (typeString == "application/vnd.openxmlformats-officedocument.theme+xml")
            type = Impl::XLContentType::Theme;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
            type = Impl::XLContentType::Styles;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
            type = Impl::XLContentType::SharedStrings;
        else if (typeString == "application/vnd.openxmlformats-officedocument.drawing+xml")
            type = Impl::XLContentType::Drawing;
        else if (typeString == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
            type = Impl::XLContentType::Chart;
        else if (typeString == "application/vnd.ms-office.chartstyle+xml")
            type = Impl::XLContentType::ChartStyle;
        else if (typeString == "application/vnd.ms-office.chartcolorstyle+xml")
            type = Impl::XLContentType::ChartColorStyle;
        else if (typeString == "application/vnd.ms-excel.controlproperties+xml")
            type = Impl::XLContentType::ControlProperties;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml")
            type = Impl::XLContentType::CalculationChain;
        else if (typeString == "application/vnd.ms-office.vbaProject")
            type = Impl::XLContentType::VBAProject;
        else if (typeString == "application/vnd.openxmlformats-package.core-properties+xml")
            type = Impl::XLContentType::CoreProperties;
        else if (typeString == "application/vnd.openxmlformats-officedocument.extended-properties+xml")
            type = Impl::XLContentType::ExtendedProperties;
        else if (typeString == "application/vnd.openxmlformats-officedocument.custom-properties+xml")
            type = Impl::XLContentType::CustomProperties;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
            type = Impl::XLContentType::Comments;
        else if (typeString == "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
            type = Impl::XLContentType::Table;
        else if (typeString == "application/vnd.openxmlformats-officedocument.vmlDrawing")
            type = Impl::XLContentType::VMLDrawing;
        else
            type = Impl::XLContentType::Unknown;

        return type;
    }

    /**
     * @details
     */
    std::string GetStringFromType(Impl::XLContentType type)
    {
        string typeString;

        if (type == Impl::XLContentType::WorkbookMacroEnabled)
            typeString = "application/vnd.ms-excel.Sheet.macroEnabled.main+xml";
        else if (type == Impl::XLContentType::Workbook)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.Sheet.main+xml";
        else if (type == Impl::XLContentType::Worksheet)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.Worksheet+xml";
        else if (type == Impl::XLContentType::Chartsheet)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.Chartsheet+xml";
        else if (type == Impl::XLContentType::ExternalLink)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml";
        else if (type == Impl::XLContentType::Theme)
            typeString = "application/vnd.openxmlformats-officedocument.theme+xml";
        else if (type == Impl::XLContentType::Styles)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
        else if (type == Impl::XLContentType::SharedStrings)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.SharedStrings+xml";
        else if (type == Impl::XLContentType::Drawing)
            typeString = "application/vnd.openxmlformats-officedocument.drawing+xml";
        else if (type == Impl::XLContentType::Chart)
            typeString = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
        else if (type == Impl::XLContentType::ChartStyle)
            typeString = "application/vnd.ms-office.chartstyle+xml";
        else if (type == Impl::XLContentType::ChartColorStyle)
            typeString = "application/vnd.ms-office.chartcolorstyle+xml";
        else if (type == Impl::XLContentType::ControlProperties)
            typeString = "application/vnd.ms-excel.controlproperties+xml";
        else if (type == Impl::XLContentType::CalculationChain)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";
        else if (type == Impl::XLContentType::VBAProject)
            typeString = "application/vnd.ms-office.vbaProject";
        else if (type == Impl::XLContentType::CoreProperties)
            typeString = "application/vnd.openxmlformats-package.core-properties+xml";
        else if (type == Impl::XLContentType::ExtendedProperties)
            typeString = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
        else if (type == Impl::XLContentType::CustomProperties)
            typeString = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
        else if (type == Impl::XLContentType::Comments)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
        else if (type == Impl::XLContentType::Table)
            typeString = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
        else if (type == Impl::XLContentType::VMLDrawing)
            typeString = "application/vnd.openxmlformats-officedocument.vmlDrawing";
        else
            throw XLException("Unknown ContentType");

        return typeString;
    }
}    // namespace

/**
 * @details
 */
Impl::XLContentItem::XLContentItem(XMLNode node) : m_contentNode(node) {}

/**
 * @details
 */
Impl::XLContentType Impl::XLContentItem::Type() const
{
    return GetTypeFromString(m_contentNode.attribute("ContentType").value());
}

/**
 * @details
 */
string Impl::XLContentItem::Path() const
{
    return m_contentNode.attribute("PartName").value();
}

/**
 * @details
 */
Impl::XLContentTypes::XLContentTypes(XLXmlData* xmlData) : XLAbstractXMLFile(xmlData)
{
    ParseXMLData();
}

/**
 * @details
 */
bool Impl::XLContentTypes::ParseXMLData()
{
    return true;
}

/**
 * @details
 */
void Impl::XLContentTypes::AddOverride(const string& path, XLContentType type)
{
    string typeString = GetStringFromType(type);

    auto node = XmlDocument().first_child().append_child("Override");
    node.append_attribute("PartName").set_value(path.c_str());
    node.append_attribute("ContentType").set_value(typeString.c_str());
}

/**
 * @details
 */
void Impl::XLContentTypes::DeleteOverride(const string& path)
{
    XmlDocument().document_element().remove_child(XmlDocument().document_element().find_child_by_attribute("PartName", path.c_str()));
}

/**
 * @details
 */
void Impl::XLContentTypes::DeleteOverride(XLContentItem& item)
{
    DeleteOverride(item.Path());
}

/**
 * @details
 */
Impl::XLContentItem Impl::XLContentTypes::ContentItem(const std::string& path)
{
    return XLContentItem(XmlDocument().document_element().find_child_by_attribute("PartName", path.c_str()));
}

vector<Impl::XLContentItem> Impl::XLContentTypes::getContentItems()
{
    std::vector<Impl::XLContentItem> result;
    for (auto item : XmlDocument().document_element().children()) {
        if (strcmp(item.name(), "Override") == 0) result.emplace_back(item);
    }

    return result;
}
