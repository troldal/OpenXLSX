//
// Created by Troldal on 25/07/16.
//

#include "XLRelationships_Impl.h"
#include "XLDocument_Impl.h"

#include <pugixml.hpp>
#include <algorithm>

using namespace std;
using namespace OpenXLSX;

/**
 * @details Constructor. Initializes the member variables for the new XLRelationshipItem object.
 */
Impl::XLRelationshipItem::XLRelationshipItem(XMLNode node)
        : m_relationshipNode(node) {
}

/**
 * @details Returns the m_relationshipType member variable by value.
 */
Impl::XLRelationshipType Impl::XLRelationshipItem::Type() const {

    return XLRelationshipItem::GetTypeFromString(m_relationshipNode.attribute("Type").value());
}

/**
 * @details Returns the m_relationshipTarget member variable by value.
 */
XMLAttribute Impl::XLRelationshipItem::Target() const {

    return m_relationshipNode.attribute("Target");
}

/**
 * @details Returns the m_relationshipId member variable by value.
 */
XMLAttribute Impl::XLRelationshipItem::Id() const {

    return m_relationshipNode.attribute("Id");
}

/**
 * @details Creates a XLRelationships object, which will read the XML file with the given path
 */
Impl::XLRelationships::XLRelationships(XLDocument& parent, const std::string& filePath)
        : XLAbstractXMLFile(parent, filePath),
          m_relationships() {

    ParseXMLData();
}

/**
 * @details Returns the XLRelationshipItem with the given ID, by looking it up in the m_relationships map.
 */
Impl::XLRelationshipItem Impl::XLRelationships::RelationshipByID(const std::string& id) const {

    return *find_if(m_relationships.begin(), m_relationships.end(), [&](const XLRelationshipItem& item) {
        return id == item.Id().value();
    });
}

/**
 * @details Returns the XLRelationshipItem with the requested target, by iterating through the items.
 */
Impl::XLRelationshipItem Impl::XLRelationships::RelationshipByTarget(const std::string& target) const {

    return *find_if(m_relationships.begin(), m_relationships.end(), [&](const XLRelationshipItem& item) {
        return target == item.Target().value();
    });
}

/**
 * @details Returns a const reference to the internal datastructure (std::map)
 */
std::vector<const Impl::XLRelationshipItem*> Impl::XLRelationships::Relationships() const {

    auto result = std::vector<const XLRelationshipItem*>();
    for (const auto& item : m_relationships)
        result.emplace_back(&item);

    return result;
}

void Impl::XLRelationships::DeleteRelationship(Impl::XLRelationshipItem& item) {

    remove_if(m_relationships.begin(), m_relationships.end(), [&](const XLRelationshipItem& theitem) {
        if (item.m_relationshipNode == theitem.m_relationshipNode) {
            theitem.m_relationshipNode.parent().remove_child(theitem.m_relationshipNode);
            return true;
        }

        return false;
    });
    WriteXMLData(); //TODO: is this really required?

}

/**
 * @details Adds a new relationship by creating new XML node in the .rels file and creating a new XLRelationshipItem
 * based on the newly created node.
 */
Impl::XLRelationshipItem* Impl::XLRelationships::AddRelationship(XLRelationshipType type, const std::string& target) {

    string typeString = XLRelationshipItem::GetStringFromType(type);

    string id = "rId" + to_string(GetNewRelsID());

    // Create new node in the .rels file
    auto node = XmlDocument()->first_child().append_child("Relationship");
    node.append_attribute("Id").set_value(id.c_str());
    node.append_attribute("Type").set_value(typeString.c_str());
    node.append_attribute("Target").set_value(target.c_str());

    if (type == XLRelationshipType::ExternalLinkPath) {
        node.append_attribute("TargetMode").set_value("External");
    }

    WriteXMLData(); //TODO: Is this really required?

    return &m_relationships.emplace_back(XLRelationshipItem(node));
}

/**
 * @details Each item in the XML file is read and a corresponding XLRelationshipItem object is created and added
 * to the current XLRelationship object.
 */
bool Impl::XLRelationships::ParseXMLData() {

    for (auto& theNode : XmlDocument()->first_child().children()) {
        string typeString = theNode.attribute("Type").value();
        XLRelationshipType type = XLRelationshipItem::GetTypeFromString(typeString);

        m_relationships.emplace_back(XLRelationshipItem(theNode));
    }

    return true;
}

/**
 * @details
 */
unsigned long Impl::XLRelationships::GetNewRelsID() const {

    return stoi(string(max_element(m_relationships.begin(),
                                   m_relationships.end(),
                                   [](const XLRelationshipItem& a, const XLRelationshipItem& b) {

                                       return stoi(string(a.Id().value()).substr(3))
                                               < stoi(string(b.Id().value()).substr(3));
                                   })->m_relationshipNode.attribute("Id").value()).substr(3)) + 1;
}

/**
 * @details
 */
bool Impl::XLRelationships::TargetExists(const std::string& target) const {

    return find_if(m_relationships.begin(), m_relationships.end(), [&](const XLRelationshipItem& item) {
        return (target == item.Target().value());
    }) != m_relationships.end();
}

/**
 * @details
 */
bool Impl::XLRelationships::IdExists(const std::string& id) const {

    return find_if(m_relationships.begin(), m_relationships.end(), [&](const XLRelationshipItem& item) {
        return (id == item.Id().value());
    }) != m_relationships.end();
}

/**
 * @details
 */
Impl::XLRelationshipType Impl::XLRelationshipItem::GetTypeFromString(const std::string& typeString) {

    XLRelationshipType type;

    if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
        type = XLRelationshipType::ExtendedProperties;
    else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties")
        type = XLRelationshipType::CustomProperties;
    else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        type = XLRelationshipType::Workbook;
    else if (typeString == "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
        type = XLRelationshipType::CoreProperties;
    else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
        type = XLRelationshipType::Worksheet;
    else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
        type = XLRelationshipType::Styles;
    else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
        type = XLRelationshipType::SharedStrings;
    else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain")
        type = XLRelationshipType::CalculationChain;
    else if (typeString == "http://schemas.microsoft.com/office/2006/relationships/vbaProject")
        type = XLRelationshipType::VBAProject;
    else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink")
        type = XLRelationshipType::ExternalLink;
    else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        type = XLRelationshipType::Theme;
    else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet")
        type = XLRelationshipType::ChartSheet;
    else if (typeString == "http://schemas.microsoft.com/office/2011/relationships/chartStyle")
        type = XLRelationshipType::ChartStyle;
    else if (typeString == "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle")
        type = XLRelationshipType::ChartColorStyle;
    else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing")
        type = XLRelationshipType::Drawing;
    else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
        type = XLRelationshipType::Image;
    else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart")
        type = XLRelationshipType::Chart;
    else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath")
        type = XLRelationshipType::ExternalLinkPath;
    else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings")
        type = XLRelationshipType::PrinterSettings;
    else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing")
        type = XLRelationshipType::VMLDrawing;
    else if (typeString == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp")
        type = XLRelationshipType::ControlProperties;
    else
        type = XLRelationshipType::Unknown;

    return type;
}

/**
 * @details
 */
std::string Impl::XLRelationshipItem::GetStringFromType(Impl::XLRelationshipType type) {

    string typeString;

    if (type == XLRelationshipType::ExtendedProperties)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    else if (type == XLRelationshipType::CustomProperties)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
    else if (type == XLRelationshipType::Workbook)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    else if (type == XLRelationshipType::CoreProperties)
        typeString = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    else if (type == XLRelationshipType::Worksheet)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    else if (type == XLRelationshipType::Styles)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    else if (type == XLRelationshipType::SharedStrings)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    else if (type == XLRelationshipType::CalculationChain)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
    else if (type == XLRelationshipType::VBAProject)
        typeString = "http://schemas.microsoft.com/office/2006/relationships/vbaProject";
    else if (type == XLRelationshipType::ExternalLink)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink";
    else if (type == XLRelationshipType::Theme)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    else if (type == XLRelationshipType::ChartSheet)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
    else if (type == XLRelationshipType::ChartStyle)
        typeString = "http://schemas.microsoft.com/office/2011/relationships/chartStyle";
    else if (type == XLRelationshipType::ChartColorStyle)
        typeString = "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle";
    else if (type == XLRelationshipType::Drawing)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    else if (type == XLRelationshipType::Image)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
    else if (type == XLRelationshipType::Chart)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    else if (type == XLRelationshipType::ExternalLinkPath)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";
    else if (type == XLRelationshipType::PrinterSettings)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
    else if (type == XLRelationshipType::VMLDrawing)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
    else if (type == XLRelationshipType::ControlProperties)
        typeString = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp";
    else
        throw XLException("RelationshipType not recognized!");

    return typeString;
}