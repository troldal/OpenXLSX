//
// Created by Troldal on 25/07/16.
//

#include "XLRelationships_Impl.hpp"

#include "XLDocument_Impl.hpp"

#include <algorithm>

using namespace std;
using namespace OpenXLSX;

namespace
{
    Impl::XLRelationshipType GetTypeFromString(const std::string& typeString)
    {
        Impl::XLRelationshipType type;

        if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
            type = Impl::XLRelationshipType::ExtendedProperties;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties")
            type = Impl::XLRelationshipType::CustomProperties;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
            type = Impl::XLRelationshipType::Workbook;
        else if (typeString == "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
            type = Impl::XLRelationshipType::CoreProperties;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
            type = Impl::XLRelationshipType::Worksheet;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
            type = Impl::XLRelationshipType::Styles;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
            type = Impl::XLRelationshipType::SharedStrings;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain")
            type = Impl::XLRelationshipType::CalculationChain;
        else if (typeString == "http://schemas.microsoft.com/office/2006/relationships/vbaProject")
            type = Impl::XLRelationshipType::VBAProject;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink")
            type = Impl::XLRelationshipType::ExternalLink;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
            type = Impl::XLRelationshipType::Theme;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet")
            type = Impl::XLRelationshipType::ChartSheet;
        else if (typeString == "http://schemas.microsoft.com/office/2011/relationships/chartStyle")
            type = Impl::XLRelationshipType::ChartStyle;
        else if (typeString == "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle")
            type = Impl::XLRelationshipType::ChartColorStyle;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing")
            type = Impl::XLRelationshipType::Drawing;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
            type = Impl::XLRelationshipType::Image;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart")
            type = Impl::XLRelationshipType::Chart;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath")
            type = Impl::XLRelationshipType::ExternalLinkPath;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings")
            type = Impl::XLRelationshipType::PrinterSettings;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing")
            type = Impl::XLRelationshipType::VMLDrawing;
        else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp")
            type = Impl::XLRelationshipType::ControlProperties;
        else
            type = Impl::XLRelationshipType::Unknown;

        return type;
    }

    std::string GetStringFromType(Impl::XLRelationshipType type)
    {
        string typeString;

        if (type == Impl::XLRelationshipType::ExtendedProperties)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        else if (type == Impl::XLRelationshipType::CustomProperties)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
        else if (type == Impl::XLRelationshipType::Workbook)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        else if (type == Impl::XLRelationshipType::CoreProperties)
            typeString = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        else if (type == Impl::XLRelationshipType::Worksheet)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        else if (type == Impl::XLRelationshipType::Styles)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        else if (type == Impl::XLRelationshipType::SharedStrings)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
        else if (type == Impl::XLRelationshipType::CalculationChain)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
        else if (type == Impl::XLRelationshipType::VBAProject)
            typeString = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
        else if (type == Impl::XLRelationshipType::ExternalLink)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
        else if (type == Impl::XLRelationshipType::Theme)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
        else if (type == Impl::XLRelationshipType::ChartSheet)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
        else if (type == Impl::XLRelationshipType::ChartStyle)
            typeString = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
        else if (type == Impl::XLRelationshipType::ChartColorStyle)
            typeString = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
        else if (type == Impl::XLRelationshipType::Drawing)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
        else if (type == Impl::XLRelationshipType::Image)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        else if (type == Impl::XLRelationshipType::Chart)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
        else if (type == Impl::XLRelationshipType::ExternalLinkPath)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
        else if (type == Impl::XLRelationshipType::PrinterSettings)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
        else if (type == Impl::XLRelationshipType::VMLDrawing)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
        else if (type == Impl::XLRelationshipType::ControlProperties)
            typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";
        else
            throw XLException("RelationshipType not recognized!");

        return typeString;
    }

    uint32_t GetNewRelsID(XMLNode relationshipsNode)
    {
        return stoi(string(max_element(relationshipsNode.children().begin(),
                                       relationshipsNode.children().end(),
                                       [](XMLNode a, XMLNode b) {
                                           return stoi(string(a.attribute("Id").value()).substr(3)) <
                                                  stoi(string(b.attribute("Id").value()).substr(3));
                                       })
                               ->attribute("Id")
                               .value())
                        .substr(3)) +
               1;
    }
}    // namespace

/**
 * @details Constructor. Initializes the member variables for the new XLRelationshipItem object.
 */
Impl::XLRelationshipItem::XLRelationshipItem(XMLNode node) : m_relationshipNode(node) {}

/**
 * @details Returns the m_relationshipType member variable by value.
 */
Impl::XLRelationshipType Impl::XLRelationshipItem::Type() const
{
    return GetTypeFromString(m_relationshipNode.attribute("Type").value());
}

/**
 * @details Returns the m_relationshipTarget member variable by value.
 */
std::string Impl::XLRelationshipItem::Target() const
{
    return m_relationshipNode.attribute("Target").value();
}

/**
 * @details Returns the m_relationshipId member variable by value.
 */
std::string Impl::XLRelationshipItem::Id() const
{
    return m_relationshipNode.attribute("Id").value();
}

/**
 * @details Creates a XLRelationships object, which will read the XML file with the given path
 */
Impl::XLRelationships::XLRelationships(XLXmlData* xmlData) : XLAbstractXMLFile(xmlData)
{
    ParseXMLData();
}

/**
 * @details Returns the XLRelationshipItem with the given ID, by looking it up in the m_relationships map.
 */
Impl::XLRelationshipItem Impl::XLRelationships::RelationshipByID(const std::string& id) const
{
    return XLRelationshipItem(XmlDocument().document_element().find_child_by_attribute("Id", id.c_str()));
}

/**
 * @details Returns the XLRelationshipItem with the requested target, by iterating through the items.
 */
Impl::XLRelationshipItem Impl::XLRelationships::RelationshipByTarget(const std::string& target) const
{
    return XLRelationshipItem(XmlDocument().document_element().find_child_by_attribute("Target", target.c_str()));
}

/**
 * @details Returns a const reference to the internal datastructure (std::map)
 */
std::vector<Impl::XLRelationshipItem> Impl::XLRelationships::Relationships() const
{
    auto result = std::vector<XLRelationshipItem>();
    for (const auto& item : XmlDocument().document_element().children()) result.emplace_back(XLRelationshipItem(item));

    return result;
}

void Impl::XLRelationships::DeleteRelationship(const string& relID)
{
    XmlDocument().document_element().remove_child(XmlDocument().document_element().find_child_by_attribute("Id", relID.c_str()));
}

void Impl::XLRelationships::DeleteRelationship(Impl::XLRelationshipItem& item)
{
    DeleteRelationship(item.Id());
}

/**
 * @details Adds a new relationship by creating new XML node in the .rels file and creating a new XLRelationshipItem
 * based on the newly created node.
 */
Impl::XLRelationshipItem Impl::XLRelationships::AddRelationship(XLRelationshipType type, const std::string& target)
{
    string typeString = GetStringFromType(type);

    string id = "rId" + to_string(GetNewRelsID(XmlDocument().document_element()));

    // Create new node in the .rels file
    auto node = XmlDocument().document_element().append_child("Relationship");
    node.append_attribute("Id").set_value(id.c_str());
    node.append_attribute("Type").set_value(typeString.c_str());
    node.append_attribute("Target").set_value(target.c_str());

    if (type == XLRelationshipType::ExternalLinkPath) {
        node.append_attribute("TargetMode").set_value("External");
    }

    return XLRelationshipItem(node);
}

/**
 * @details Each item in the XML file is read and a corresponding XLRelationshipItem object is created and added
 * to the current XLRelationship object.
 */
bool Impl::XLRelationships::ParseXMLData()
{
    return true;
}

/**
 * @details
 */
bool Impl::XLRelationships::TargetExists(const std::string& target) const
{
    return XmlDocument().document_element().find_child_by_attribute("Target", target.c_str()) != nullptr;
}

/**
 * @details
 */
bool Impl::XLRelationships::IdExists(const std::string& id) const
{
    return XmlDocument().document_element().find_child_by_attribute("Id", id.c_str()) != nullptr;
}
