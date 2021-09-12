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

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLRelationships.hpp"

using namespace OpenXLSX;

namespace
{
    XLRelationshipType GetTypeFromString(const std::string& typeString)
    {
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
            type = XLRelationshipType::Chartsheet;
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

    std::string GetStringFromType(XLRelationshipType type)
    {
        std::string typeString;

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
        else if (type == XLRelationshipType::Chartsheet)
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
            throw XLInternalError("RelationshipType not recognized!");

        return typeString;
    }

    uint32_t GetNewRelsID(XMLNode relationshipsNode)
    {
        return static_cast<uint32_t>(stoi(std::string(std::max_element(relationshipsNode.children().begin(),
                                                                       relationshipsNode.children().end(),
                                                                       [](XMLNode a, XMLNode b) {
                                                                           return stoi(std::string(a.attribute("Id").value()).substr(3)) <
                                                                                  stoi(std::string(b.attribute("Id").value()).substr(3));
                                                                       })
                                                          ->attribute("Id")
                                                          .value())
                                              .substr(3)) +
                                     1);
    }
}    // namespace

XLRelationshipItem::XLRelationshipItem() : m_relationshipNode(std::make_unique<XMLNode>()) {}

/**
 * @details Constructor. Initializes the member variables for the new XLRelationshipItem object.
 */
XLRelationshipItem::XLRelationshipItem(const XMLNode& node) : m_relationshipNode(std::make_unique<XMLNode>(node)) {}

XLRelationshipItem::~XLRelationshipItem() = default;

XLRelationshipItem::XLRelationshipItem(const XLRelationshipItem& other)
    : m_relationshipNode(std::make_unique<XMLNode>(*other.m_relationshipNode))
{}

XLRelationshipItem& XLRelationshipItem::operator=(const XLRelationshipItem& other)
{
    if (&other != this) *m_relationshipNode = *other.m_relationshipNode;
    return *this;
}

/**
 * @details Returns the m_relationshipType member variable by getValue.
 */
XLRelationshipType XLRelationshipItem::type() const
{
    return GetTypeFromString(m_relationshipNode->attribute("Type").value());
}

/**
 * @details Returns the m_relationshipTarget member variable by getValue.
 */
std::string XLRelationshipItem::target() const
{
    return m_relationshipNode->attribute("Target").value();
}

/**
 * @details Returns the m_relationshipId member variable by getValue.
 */
std::string XLRelationshipItem::id() const
{
    return m_relationshipNode->attribute("Id").value();
}

/**
 * @details Creates a XLRelationships object, which will read the XML file with the given path
 */
XLRelationships::XLRelationships(XLXmlData* xmlData) : XLXmlFile(xmlData) {}

XLRelationships::~XLRelationships() = default;

/**
 * @details Returns the XLRelationshipItem with the given ID, by looking it up in the m_relationships map.
 */
XLRelationshipItem XLRelationships::relationshipById(const std::string& id) const
{
    return XLRelationshipItem(xmlDocument().document_element().find_child_by_attribute("Id", id.c_str()));
}

/**
 * @details Returns the XLRelationshipItem with the requested target, by iterating through the items.
 */
XLRelationshipItem XLRelationships::relationshipByTarget(const std::string& target) const
{
    return XLRelationshipItem(xmlDocument().document_element().find_child_by_attribute("Target", target.c_str()));
}

/**
 * @details Returns a const reference to the internal datastructure (std::map)
 */
std::vector<XLRelationshipItem> XLRelationships::relationships() const
{
    auto result = std::vector<XLRelationshipItem>();
    for (const auto& item : xmlDocument().document_element().children()) result.emplace_back(XLRelationshipItem(item));

    return result;
}

void XLRelationships::deleteRelationship(const std::string& relID)
{
    xmlDocument().document_element().remove_child(xmlDocument().document_element().find_child_by_attribute("Id", relID.c_str()));
}

void XLRelationships::deleteRelationship(const XLRelationshipItem& item)
{
    deleteRelationship(item.id());
}

/**
 * @details Adds a new relationship by creating new XML node in the .rels file and creating a new XLRelationshipItem
 * based on the newly created node.
 */
XLRelationshipItem XLRelationships::addRelationship(XLRelationshipType type, const std::string& target)
{
    std::string typeString = GetStringFromType(type);

    std::string id = "rId" + std::to_string(GetNewRelsID(xmlDocument().document_element()));

    // Create new node in the .rels file
    auto node = xmlDocument().document_element().append_child("Relationship");
    node.append_attribute("Id").set_value(id.c_str());
    node.append_attribute("Type").set_value(typeString.c_str());
    node.append_attribute("Target").set_value(target.c_str());

    if (type == XLRelationshipType::ExternalLinkPath) {
        node.append_attribute("TargetMode").set_value("External");
    }

    return XLRelationshipItem(node);
}

/**
 * @details
 */
bool XLRelationships::targetExists(const std::string& target) const
{
    return xmlDocument().document_element().find_child_by_attribute("Target", target.c_str()) != nullptr;
}

/**
 * @details
 */
bool XLRelationships::idExists(const std::string& id) const
{
    return xmlDocument().document_element().find_child_by_attribute("Id", id.c_str()) != nullptr;
}
