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
#include <cstdint>      // uint32_t
#include <memory>       // std::make_unique
#include <pugixml.hpp>
#include <stdexcept>    // std::invalid_argument
#include <string>       // std::stoi, std::literals::string_literals
#include <vector>       // std::vector

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLRelationships.hpp"

#include <XLException.hpp>

using namespace OpenXLSX;

namespace
{
    const std::string relationshipDomainOpenXml2006          = "http://schemas.openxmlformats.org/officeDocument/2006";
    const std::string relationshipDomainOpenXml2006CoreProps = "http://schemas.openxmlformats.org/package/2006";
    const std::string relationshipDomainMicrosoft2006        = "http://schemas.microsoft.com/office/2006";
    const std::string relationshipDomainMicrosoft2011        = "http://schemas.microsoft.com/office/2011";

    /**
     * @note 2024-08-31: Included a "dumb" fallback solution in relationship tests to support
     *          previously unknown relationship domains, e.g. type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet"
     */
    XLRelationshipType GetTypeFromString(const std::string& typeString)
    {
        // TODO 2024-08-09: support dumb applications that implemented relationship Type in different case (e.g. vmldrawing instead of vmlDrawing)
        //                  easy approach: convert typestring and comparison string to all lower characters
        size_t comparePos = 0; // start by comparing full relationship type strings
        do {
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/extended-properties")
                return XLRelationshipType::ExtendedProperties;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/custom-properties")
                return XLRelationshipType::CustomProperties;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/officeDocument")
                return XLRelationshipType::Workbook;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/worksheet")
                return XLRelationshipType::Worksheet;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/styles")
                return XLRelationshipType::Styles;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/sharedStrings")
                return XLRelationshipType::SharedStrings;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/calcChain")
                return XLRelationshipType::CalculationChain;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/externalLink")
                return XLRelationshipType::ExternalLink;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/theme")
                return XLRelationshipType::Theme;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/chartsheet")
                return XLRelationshipType::Chartsheet;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/drawing")
                return XLRelationshipType::Drawing;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/image")
                return XLRelationshipType::Image;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/chart")
                return XLRelationshipType::Chart;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/externalLinkPath")
                return XLRelationshipType::ExternalLinkPath;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/printerSettings")
                return XLRelationshipType::PrinterSettings;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/vmlDrawing")
                return XLRelationshipType::VMLDrawing;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/ctrlProp")
                return XLRelationshipType::ControlProperties;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006) + "/relationships/comments")
                return XLRelationshipType::Comments;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainOpenXml2006CoreProps) + "/relationships/metadata/core-properties")
                return XLRelationshipType::CoreProperties;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainMicrosoft2006) + "/relationships/vbaProject")
                return XLRelationshipType::VBAProject;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainMicrosoft2011) + "/relationships/chartStyle")
                return XLRelationshipType::ChartStyle;
            if (typeString.substr(comparePos) == (comparePos ? "" : relationshipDomainMicrosoft2011) + "/relationships/chartColorStyle")
                return XLRelationshipType::ChartColorStyle;

            // ===== relationship could not be identified
            if (comparePos == 0 )    // If fallback solution has not yet been tried
                comparePos = typeString.find("/relationships/");    // attempt to find the relationships section of the type string, regardless of domain
            else                     // If fallback solution was tried & unsuccessful
                comparePos = 0;                                     // trigger loop exit

        } while(comparePos > 0 && comparePos != std::string::npos);
        // ===== loop exits if comparePos is not within typeString (= fallback solution failed or not possible)

        return XLRelationshipType::Unknown; // default: relationship could not be identified
    }

} //    namespace

namespace OpenXLSX_XLRelationships {    // make GetStringFromType accessible throughout the project (for use by XLAppProperties)
    using namespace OpenXLSX;
    std::string GetStringFromType(XLRelationshipType type)
    {
        switch (type) {
            case XLRelationshipType::ExtendedProperties: return relationshipDomainOpenXml2006 + "/relationships/extended-properties";
            case XLRelationshipType::CustomProperties:   return relationshipDomainOpenXml2006 + "/relationships/custom-properties";
            case XLRelationshipType::Workbook:           return relationshipDomainOpenXml2006 + "/relationships/officeDocument";
            case XLRelationshipType::Worksheet:          return relationshipDomainOpenXml2006 + "/relationships/worksheet";
            case XLRelationshipType::Styles:             return relationshipDomainOpenXml2006 + "/relationships/styles";
            case XLRelationshipType::SharedStrings:      return relationshipDomainOpenXml2006 + "/relationships/sharedStrings";
            case XLRelationshipType::CalculationChain:   return relationshipDomainOpenXml2006 + "/relationships/calcChain";
            case XLRelationshipType::ExternalLink:       return relationshipDomainOpenXml2006 + "/relationships/externalLink";
            case XLRelationshipType::Theme:              return relationshipDomainOpenXml2006 + "/relationships/theme";
            case XLRelationshipType::Chartsheet:         return relationshipDomainOpenXml2006 + "/relationships/chartsheet";
            case XLRelationshipType::Drawing:            return relationshipDomainOpenXml2006 + "/relationships/drawing";
            case XLRelationshipType::Image:              return relationshipDomainOpenXml2006 + "/relationships/image";
            case XLRelationshipType::Chart:              return relationshipDomainOpenXml2006 + "/relationships/chart";
            case XLRelationshipType::ExternalLinkPath:   return relationshipDomainOpenXml2006 + "/relationships/externalLinkPath";
            case XLRelationshipType::PrinterSettings:    return relationshipDomainOpenXml2006 + "/relationships/printerSettings";
            case XLRelationshipType::VMLDrawing:         return relationshipDomainOpenXml2006 + "/relationships/vmlDrawing";
            case XLRelationshipType::ControlProperties:  return relationshipDomainOpenXml2006 + "/relationships/ctrlProp";
            case XLRelationshipType::Comments:           return relationshipDomainOpenXml2006 + "/relationships/comments";
            case XLRelationshipType::CoreProperties:     return relationshipDomainOpenXml2006CoreProps + "/relationships/metadata/core-properties";
            case XLRelationshipType::VBAProject:         return relationshipDomainMicrosoft2006 + "/relationships/vbaProject";
            case XLRelationshipType::ChartStyle:         return relationshipDomainMicrosoft2011 + "/relationships/chartStyle";
            case XLRelationshipType::ChartColorStyle:    return relationshipDomainMicrosoft2011 + "/relationships/chartColorStyle";
            default:
                throw XLInternalError("RelationshipType not recognized!");
        }
    }
} //    namespace OpenXLSX_XLRelationships

namespace { //    re-open anonymous namespace
    uint32_t GetNewRelsID(XMLNode relationshipsNode)
    {
        using namespace std::literals::string_literals;
        // ===== workaround for pugi::xml_node currently not having an iterator for node_element only
        XMLNode  relationship = relationshipsNode.first_child_of_type(pugi::node_element);
        uint32_t newId        = 1;    // default
        while (not relationship.empty()) {
            uint32_t id;
            try {
                id = std::stoi(std::string(relationship.attribute("Id").value()).substr(3));
            }
            catch (std::invalid_argument const & e) {    // expected stoi exception
                throw XLInputError("GetNewRelsID could not convert attribute Id to uint32_t ("s + e.what() + ")"s);
            }
            catch (...) {    // catch all other errors during conversion of attribute to uint32_t
                throw XLInputError("GetNewRelsID could not convert attribute Id to uint32_t"s);
            }
            if (id >= newId) newId = id + 1;
            relationship = relationship.next_sibling_of_type(pugi::node_element);
        }
        return newId;
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
XLRelationshipType XLRelationshipItem::type() const { return GetTypeFromString(m_relationshipNode->attribute("Type").value()); }

/**
 * @details Returns the m_relationshipTarget member variable by getValue.
 */
std::string XLRelationshipItem::target() const { return m_relationshipNode->attribute("Target").value(); }

/**
 * @details Returns the m_relationshipId member variable by getValue.
 */
std::string XLRelationshipItem::id() const { return m_relationshipNode->attribute("Id").value(); }

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
 * @details Returns a const reference to the internal datastructure (std::vector)
 */
std::vector<XLRelationshipItem> XLRelationships::relationships() const
{
    // ===== workaround for pugi::xml_node currently not having an iterator for node_element only
    auto    result = std::vector<XLRelationshipItem>();
    XMLNode item   = xmlDocument().document_element().first_child_of_type(pugi::node_element);
    while (not item.empty()) {
        result.emplace_back(XLRelationshipItem(item));
        item = item.next_sibling_of_type(pugi::node_element);
    }
    // ===== if a node_element iterator can be implemented for pugi::xml_node, the below code can be used again
    // for (const auto& item : xmlDocument().document_element().children()) result.emplace_back(XLRelationshipItem(item));

    return result;
}

void XLRelationships::deleteRelationship(const std::string& relID)
{
    xmlDocument().document_element().remove_child(xmlDocument().document_element().find_child_by_attribute("Id", relID.c_str()));
}

void XLRelationships::deleteRelationship(const XLRelationshipItem& item) { deleteRelationship(item.id()); }

/**
 * @details Adds a new relationship by creating new XML node in the .rels file and creating a new XLRelationshipItem
 * based on the newly created node.
 */
XLRelationshipItem XLRelationships::addRelationship(XLRelationshipType type, const std::string& target)
{
    const std::string typeString = OpenXLSX_XLRelationships::GetStringFromType(type);
    const std::string id         = "rId" + std::to_string(GetNewRelsID(xmlDocument().document_element()));

    // Create new node in the .rels file
    XMLNode node = xmlDocument().document_element().last_child_of_type(pugi::node_element); // pretty formatting: insert new node before
                                                                                            // whitespaces following last element node
    if (not node.empty())
        node = xmlDocument().document_element().insert_child_after("Relationship", node);
    else
        node = xmlDocument().document_element().append_child("Relationship");
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
