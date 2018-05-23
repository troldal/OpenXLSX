//
// Created by Troldal on 25/07/16.
//

#include "XLRelationships.h"
#include "XLDocument.h"

using namespace std;
using namespace OpenXLSX;

/**
 * @details Constructor. Initializes the member variables for the new XLRelationshipItem object.
 */
XLRelationshipItem::XLRelationshipItem(XMLNode node,
                                       XLRelationshipType type,
                                       const std::string &target,
                                       const std::string &id)
    : m_relationshipNode(node),
      m_relationshipType(type),
      m_relationshipTarget(target),
      m_relationshipId(id)
{

}

/**
 * @details Returns the m_relationshipType member variable by value.
 */
XLRelationshipType XLRelationshipItem::Type() const
{
    return m_relationshipType;
}

/**
 * @details Returns the m_relationshipTarget member variable by value.
 */
const std::string &XLRelationshipItem::Target() const
{
    return m_relationshipTarget;
}

/**
 * @details Returns the m_relationshipId member variable by value.
 */
const std::string &XLRelationshipItem::Id() const
{
    return m_relationshipId;
}

/**
 * @details Deletes the underlying XML node and sets the object to a safe state. Deletion of the object can only be done
 * from the parent XLRelationships object.
 */
void XLRelationshipItem::Delete()
{
    // Delete the XML node
    if (m_relationshipNode) m_relationshipNode.parent().remove_child(m_relationshipNode);

    // Set the object to a safe State
    m_relationshipNode = XMLNode();
    m_relationshipType = XLRelationshipType::Unknown;
    m_relationshipTarget = "";
    m_relationshipId = "";

}

/**
 * @details Creates a XLRelationships object, which will read the XML file with the given path
 */
XLRelationships::XLRelationships(XLDocument &parent,
                                 const std::string &filePath)
    : XLAbstractXMLFile(parent, filePath),
      XLSpreadsheetElement(parent),
      m_relationships(),
      m_relationshipCount(0)
{
    ParseXMLData(); // This will call the ParseXMLData method.
}

/**
 * @details Returns the XLRelationshipItem with the given ID, by looking it up in the m_relationships map.
 */
const XLRelationshipItem *XLRelationships::RelationshipByID(const std::string &id) const
{
    return Relationships()->at(id).get();
}

/**
 * @details Returns the XLRelationshipItem with the given ID, by looking it up in the m_relationships map.
 */
XLRelationshipItem * XLRelationships::RelationshipByID(const std::string &id)
{
    return Relationships()->at(id).get();
}

/**
 * @details Returns the XLRelationshipItem with the requested target, by iterating through the items.
 */
const XLRelationshipItem *XLRelationships::RelationshipByTarget(const std::string &target) const
{

    for (auto const &item : *Relationships()) {
        if (item.second->Target() == target) return item.second.get();
    }

    return nullptr;
}

/**
 * @details Returns the XLRelationshipItem with the requested target, by iterating through the items.
 */
XLRelationshipItem *XLRelationships::RelationshipByTarget(const std::string &target)
{

    for (auto &item : *Relationships()) {
        if (item.second->Target() == target) return item.second.get();
    }

    return nullptr;
}

/**
 * @details Returns a const reference to the internal datastructure (std::map)
 */
const XLRelationshipMap *XLRelationships::Relationships() const
{
    return &m_relationships;
}

/**
 * @details Returns a mutable reference to the internal datastructure (std::map)
 * @todo Consider if there is a more elegant way of doing this.
 */
XLRelationshipMap *XLRelationships::relationshipsMutable()
{
    return &m_relationships;
}

/**
 * @details
 */
void XLRelationships::DeleteRelationship(const std::string &id)
{

    Relationships()->at(id)->Delete();
    relationshipsMutable()->erase(id); // Delete item from the Relationships map
    CommitXMLData(); //TODO: is this really required?
    SetModified();
}

/**
 * @details Adds a new relationship by creating new XML node in the .rels file and creating a new XLRelationshipItem
 * based on the newly created node.
 */
XLRelationshipItem * XLRelationships::AddRelationship(XLRelationshipType type, const std::string &target)
{
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
    else throw runtime_error("RelationshipType not recognized!");

    m_relationshipCount++;

    string id = "rId" + to_string(m_relationshipCount);

    // Create new node in the .rels file
    auto node = XmlDocument()->first_child().append_child("Relationship");
    node.append_attribute("Id").set_value(id.c_str());
    node.append_attribute("Type").set_value(typeString.c_str());
    node.append_attribute("Target").set_value(target.c_str());

    if (type == XLRelationshipType::ExternalLinkPath) {
        node.append_attribute("TargetMode").set_value("External");
    }

    // Create new XLRelationshipItem object and add to internal datastructure.
    unique_ptr<XLRelationshipItem> rShip(new XLRelationshipItem(node, type, target, id));
    XLRelationshipItem *result = rShip.get();
    relationshipsMutable()->insert_or_assign(id, move(rShip));
    SetModified();

    CommitXMLData(); //TODO: Is this really required?

    return result;
}

/**
 * @details Each item in the XML file is read and a corresponding XLRelationshipItem object is created and added
 * to the current XLRelationship object.
 */
bool XLRelationships::ParseXMLData()
{
    for (auto &theNode : XmlDocument()->first_child().children()) {
        XLRelationshipType type;
        string typeString = theNode.attribute("Type").value();

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

        unique_ptr<XLRelationshipItem> rShip(new XLRelationshipItem(theNode,
                                                                    type,
                                                                    theNode.attribute("Target").value(),
                                                                    theNode.attribute("Id").value()));

        relationshipsMutable()->insert_or_assign(theNode.attribute("Id").value(), move(rShip));
    }

    m_relationshipCount = m_relationships.size();

    return true;
}

bool XLRelationships::TargetExists(const std::string &target) const
{
    if (RelationshipByTarget(target))
        return true;
    else
        return false;
}

bool XLRelationships::IdExists(const std::string &id) const
{
    if (RelationshipByID(id))
        return true;
    else
        return false;
}
