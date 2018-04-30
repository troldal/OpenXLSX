//
// Created by Troldal on 13/08/16.
//

#include <cstring>

#include "XLContentTypes.h"
#include "XLDocument.h"

using namespace std;
using namespace OpenXLSX;

/**
 * @details
 */
XLContentItem::XLContentItem(XMLNode &node,
                             const std::string &path,
                             XLContentType type)
    : m_contentNode(make_unique<XMLNode>(node)),
      m_contentPath(path),
      m_contentType(type)
{

}

/**
 * @details
 */
XLContentItem::XLContentItem(const XLContentItem &other)
    : m_contentNode(make_unique<XMLNode>(*other.m_contentNode)),
      m_contentPath(other.m_contentPath),
      m_contentType(other.m_contentType)
{
    //todo: copying of base class members...?
}

/**
 * @details
 */
XLContentItem::XLContentItem(XLContentItem &&other)
    : m_contentNode(move(other.m_contentNode)),
      m_contentPath(move(other.m_contentPath)),
      m_contentType(move(other.m_contentType))
{

}

/**
 * @details
 */
XLContentItem &XLContentItem::operator=(const XLContentItem &other)
{
    m_contentNode = make_unique<XMLNode>(*other.m_contentNode);
    m_contentPath = other.m_contentPath;
    m_contentType = other.m_contentType;

    return *this;
}

/**
 * @details
 */
XLContentItem &XLContentItem::operator=(XLContentItem &&other)
{
    m_contentNode = move(other.m_contentNode);
    m_contentPath = move(other.m_contentPath);
    m_contentType = move(other.m_contentType);

    return *this;
}

/**
 * @details
 */
XLContentType XLContentItem::Type() const
{
    return m_contentType;
}

/**
 * @details
 */
const string &XLContentItem::Path() const
{
    return m_contentPath;
}

/**
 * @details
 */
void XLContentItem::DeleteItem()
{
    if (m_contentNode) m_contentNode->parent().remove_child(*m_contentNode);

    m_contentNode = nullptr;
    m_contentPath = "";
    m_contentType = XLContentType::Unknown;

}

/**
 * @details
 */
XLContentTypes::XLContentTypes(XLDocument &parent,
                               const string &filePath)
    : XLAbstractXMLFile(parent, filePath),
      XLSpreadsheetElement(parent),
      m_defaults(),
      m_overrides()
{
    ParseXMLData();
}

/**
 * @details
 */
bool XLContentTypes::ParseXMLData()
{
    m_defaults.clear();
    m_overrides.clear();

    std::string strOverride = "Override";
    std::string strDefault = "Default";

    auto node = XmlDocument()->first_child();

    while (node) {
        if (strcmp(node.name(), "Default") == 0) {
            std::string extension = "";
            std::string contentType = "";

            if (node.attribute("Extension")) extension = node.attribute("Extension").value();
            if (node.attribute("ContentType")) contentType = node.attribute("ContentType").value();
            m_defaults.insert_or_assign(extension, make_unique<XMLNode>(node));
        }
        else if (strcmp(node.name(), "Override") == 0) {
            string path = node.attribute("PartName").value();
            XLContentType type;
            string typeString = node.attribute("ContentType").value();

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

            unique_ptr<XLContentItem> item(new XLContentItem(node, path, type));
            m_overrides.insert_or_assign(path, move(item));
        }

        node = node.next_sibling();
    }

    return true;
}

/**
 * @details
 */
void XLContentTypes::AddDefault(const string &key,
                                XMLNode &node)
{
    m_defaults.insert_or_assign(key, make_unique<XMLNode>(node));
    SetModified();
}

/**
 * @details
 */
void XLContentTypes::addOverride(const string &path,
                                 XLContentType type)
{
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
        return;

    auto node = XmlDocument()->root().append_child("override");
    node.append_attribute("PartName").set_value(path.c_str());
    node.append_attribute("ContentType").set_value(typeString.c_str());

    unique_ptr<XLContentItem> item(new XLContentItem(node, path, type));
    m_overrides.insert_or_assign(path, move(item));
    SetModified();
}

/**
 * @details
 */
void XLContentTypes::ClearOverrides()
{
    m_overrides.clear();
    SetModified();
}

/**
 * @details
 */
const XLContentItemMap *XLContentTypes::contentItems() const
{
    return &m_overrides;
}

/**
 * @details
 */
XLContentItem *XLContentTypes::ContentItem(const std::string &path)
{
    return m_overrides.at(path).get();
}
