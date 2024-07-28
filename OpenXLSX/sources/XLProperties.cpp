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
#include <pugixml.hpp>
#include <string>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLProperties.hpp"
#include "XLRelationships.hpp"    // GetStringFromType

#include <XLException.hpp>

using namespace OpenXLSX;

namespace
{
    XMLNode headingPairsNode(XMLNode docNode) { return docNode.child("HeadingPairs").first_child_of_type(pugi::node_element); }

    XMLAttribute headingPairsSize(XMLNode docNode)
    {
        return docNode.child("HeadingPairs").first_child_of_type(pugi::node_element).attribute("size");
    }

    std::vector<std::string> headingPairsCategoriesStrings(XMLNode docNode)
    {
        // 2024-05-28 DONE: tested this code with two pairs in headingPairsNode
        std::vector<std::string> result;
        XMLNode                  item = headingPairsNode(docNode).first_child_of_type(pugi::node_element);
        while (not item.empty()) {
            result.push_back(item.first_child_of_type(pugi::node_element).child_value());
            item = item.next_sibling_of_type(pugi::node_element)
                       .next_sibling_of_type(pugi::node_element);    // advance two elements to skip count node
        }
        return result; // 2024-05-28: std::move should not be used when the operand of a return statement is the name of a local variable
                       // as this can prevent named return value optimization (NRVO, copy elision)
    }

    XMLNode sheetNames(XMLNode docNode) { return docNode.child("TitlesOfParts").first_child_of_type(pugi::node_element); }

    XMLAttribute sheetCount(XMLNode docNode) { return sheetNames(docNode).attribute("size"); }
}    // namespace

/**
 * @details
 */
void XLProperties::createFromTemplate()
{
    // std::cout << "XLProperties created with empty docProps/core.xml, creating from scratch!" << std::endl;
    if( m_xmlData == nullptr )
        throw XLInternalError("XLProperties m_xmlData is nullptr");

    // NOTE: there is no functionality in pugixml to include the standalone="yes" attribute in the <xml> element node

    // ===== OpenXLSX_XLRelationships::GetStringFromType yields almost the string needed here, with added /relationships
    //       TBD: use hardcoded string?
    std::string xmlns = OpenXLSX_XLRelationships::GetStringFromType(XLRelationshipType::CoreProperties);
    const std::string rels = "/relationships/";
    size_t pos = xmlns.find(rels);
    if (pos != std::string::npos)
        xmlns.replace(pos, rels.size(), "/");
    else
        xmlns = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"; // fallback to hardcoded string

    XMLNode props = xmlDocument().prepend_child("cp:coreProperties");
    props.append_attribute("xmlns:cp") = xmlns.c_str();
    props.append_attribute("xmlns:dc") = "http://purl.org/dc/elements/1.1/";
    props.append_attribute("xmlns:dcterms") = "http://purl.org/dc/terms/";
    props.append_attribute("xmlns:dcmitype") = "http://purl.org/dc/dcmitype/";
    props.append_attribute("xmlns:xsi") = "http://www.w3.org/2001/XMLSchema-instance";

    props.append_child("dc:creator").text().set("Kenneth Balslev");
    props.append_child("cp:lastModifiedBy").text().set("Kenneth Balslev");

    XMLNode prop {};
    prop = props.append_child("dcterms:created");
    prop.append_attribute("xsi:type") = "dcterms:W3CDTF";
    prop.text().set("2019-08-16T00:34:14Z");
    prop = props.append_child("dcterms:modified");
    prop.append_attribute("xsi:type") = "dcterms:W3CDTF";
    prop.text().set("2019-08-16T00:34:26Z");
}

/**
 * @details
 */
XLProperties::XLProperties(XLXmlData* xmlData) : XLXmlFile(xmlData)
{
    XMLNode doc = xmlData->getXmlDocument()->document_element();
    XMLNode child = doc.first_child_of_type(pugi::node_element);
    size_t childCount = 0;
    while (not child.empty()) {
        ++childCount;
        child = child.next_sibling_of_type(pugi::node_element);
        break; // one child is enough to determine document is not empty.
    }
    if( !childCount ) createFromTemplate();
}

/**
 * @details
 */
XLProperties::~XLProperties() = default;

/**
 * @details
 */
void XLProperties::setProperty(const std::string& name, const std::string& value)
{
    if (m_xmlData == nullptr) return;
    XMLNode node = xmlDocument().document_element().child(name.c_str());
    if (node.empty())
        node = xmlDocument().document_element().append_child(name.c_str());    // .append_child, to be in line with ::property behavior

    node.text().set(value.c_str());
}

/**
 * @details
 */
void XLProperties::setProperty(const std::string& name, int value) { setProperty(name, std::to_string(value)); }

/**
 * @details
 */
void XLProperties::setProperty(const std::string& name, double value) { setProperty(name, std::to_string(value)); }

/**
 * @details
 */
std::string XLProperties::property(const std::string& name) const
{
    if (m_xmlData == nullptr) return "";
    XMLNode property = xmlDocument().document_element().child(name.c_str());
    if (property.empty()) property = xmlDocument().document_element().append_child(name.c_str());

    return property.text().get();
}

/**
 * @details
 */
void XLProperties::deleteProperty(const std::string& name)
{
    if (m_xmlData == nullptr) return;
    if (const XMLNode property = xmlDocument().document_element().child(name.c_str()); not property.empty())
        xmlDocument().document_element().remove_child(property);
}


/**
 * @details
 */
void XLAppProperties::createFromTemplate(XMLDocument const & workbookXml)
{
    // std::cout << "XLAppProperties created with empty docProps/app.xml, creating from scratch!" << std::endl;
    if( m_xmlData == nullptr )
        throw XLInternalError("XLAppProperties m_xmlData is nullptr");

    std::map< uint32_t, std::string > sheetsOrderedById;

    auto sheet = workbookXml.document_element().child("sheets").first_child_of_type(pugi::node_element);
    while (not sheet.empty()) {
        std::string sheetName = sheet.attribute("name").as_string();
        uint32_t sheetId = sheet.attribute("sheetId").as_uint();
        sheetsOrderedById.insert(std::pair<uint32_t, std::string>(sheetId, sheetName));
        sheet = sheet.next_sibling_of_type();
    }

    uint32_t worksheetCount = 0;
    for (const auto & [key, value] : sheetsOrderedById) {
        if (key != ++worksheetCount)
           throw XLInputError( "xl/workbook.xml is missing sheet with sheetId=\"" + std::to_string(worksheetCount) + "\"" );
    }

    // NOTE: there is no functionality in pugixml to include the standalone="yes" attribute in the <xml> element node

    // ===== OpenXLSX_XLRelationships::GetStringFromType yields almost the string needed here, with added /relationships
    //       TBD: use hardcoded string?
    std::string xmlns = OpenXLSX_XLRelationships::GetStringFromType(XLRelationshipType::ExtendedProperties);
    const std::string rels = "/relationships/";
    size_t pos = xmlns.find(rels);
    if (pos != std::string::npos)
        xmlns.replace(pos, rels.size(), "/");
    else
        xmlns = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"; // fallback to hardcoded string

    XMLNode props = xmlDocument().prepend_child("Properties");
    props.append_attribute("xmlns") = xmlns.c_str();
    props.append_attribute("xmlns:vt") = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    XMLNode prop {};

    props.append_child("Application").text().set("Microsoft Macintosh Excel");
    props.append_child("DocSecurity").text().set(0);
    props.append_child("ScaleCrop").text().set(false);

    XMLNode headingPairs = props.append_child("HeadingPairs");
    XMLNode vecHP = headingPairs.append_child("vt:vector");
    vecHP.append_attribute("size") = 2;
    vecHP.append_attribute("baseType") = "variant";
    vecHP.append_child("vt:variant").append_child("vt:lpstr").text().set("Worksheets");
    vecHP.append_child("vt:variant").append_child("vt:i4").text().set(1); // TBD: should this be count of worksheets?

    XMLNode sheetsVector = headingPairs.append_child("TitlesOfParts").append_child("vt:vector");
    sheetsVector.append_attribute("size") = worksheetCount;
    sheetsVector.append_attribute("baseType") = "lpstr";
    for (const auto & [key, value] : sheetsOrderedById)
        sheetsVector.append_child("vt:lpstr").text().set(value.c_str());

    props.append_child("Company").text().set("");
    props.append_child("LinksUpToDate").text().set(false);
    props.append_child("SharedDoc").text().set(false);
    props.append_child("HyperlinksChanged").text().set(false);
    props.append_child("AppVersion").text().set("16.0300");
}

/**
 * @details
 */
XLAppProperties::XLAppProperties(XLXmlData* xmlData, XMLDocument const & workbookXml)
    : XLXmlFile(xmlData)
{
    XMLNode doc = xmlData->getXmlDocument()->document_element();
    XMLNode child = doc.first_child_of_type(pugi::node_element);
    size_t childCount = 0;
    while (not child.empty()) {
        ++childCount;
        child = child.next_sibling_of_type(pugi::node_element);
        break; // one child is enough to determine document is not empty.
    }
    if (!childCount) createFromTemplate(workbookXml); //    create fresh docProps/app.xml
}

/**
 * @details
 */
XLAppProperties::XLAppProperties(XLXmlData* xmlData) : XLXmlFile(xmlData) {}

/**
 * @details
 */
XLAppProperties::~XLAppProperties() = default;

/**
 * @details
 */
void XLAppProperties::addSheetName(const std::string& title)
{
    if (m_xmlData == nullptr) return;
    XMLNode theNode = sheetNames(xmlDocument().document_element()).append_child("vt:lpstr");
    theNode.text().set(title.c_str());
    sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);
}

/**
 * @details
 */
void XLAppProperties::deleteSheetName(const std::string& title)
{
    if (m_xmlData == nullptr) return;
    XMLNode theNode = sheetNames(xmlDocument().document_element()).first_child_of_type(pugi::node_element);
    while (not theNode.empty()) {
        if (theNode.child_value() == title) {
            sheetNames(xmlDocument().document_element()).remove_child(theNode);
            sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() - 1);
            return;
        }
        theNode = theNode.next_sibling_of_type(pugi::node_element);
    }
}

/**
 * @details
 */
void XLAppProperties::setSheetName(const std::string& oldTitle, const std::string& newTitle)
{
    if (m_xmlData == nullptr) return;
    XMLNode theNode = sheetNames(xmlDocument().document_element()).first_child_of_type(pugi::node_element);
    while (not theNode.empty()) {
        if (theNode.child_value() == oldTitle) {
            theNode.text().set(newTitle.c_str());
            return;
        }
        theNode = theNode.next_sibling_of_type(pugi::node_element);
    }
}

/**
 * @details
 */
void XLAppProperties::addHeadingPair(const std::string& name, int value)
{
    if (m_xmlData == nullptr) return;

    XMLNode HeadingPairsNode = headingPairsNode(xmlDocument().document_element());
    XMLNode item             = HeadingPairsNode.first_child_of_type(pugi::node_element);
    while (not item.empty() && item.first_child_of_type(pugi::node_element).child_value() != name)
        item = item.next_sibling_of_type(pugi::node_element)
                   .next_sibling_of_type(pugi::node_element);    // advance two elements to skip count node

    XMLNode pairCategory = item;    // could be an empty node
    XMLNode pairCountValue {};      // initialize to empty node
    if (not pairCategory.empty())
        pairCountValue = pairCategory.next_sibling_of_type(pugi::node_element).first_child_of_type(pugi::node_element);
    else {
        item = HeadingPairsNode.last_child_of_type(pugi::node_element);
        if (not item.empty())
            pairCategory = HeadingPairsNode.insert_child_after("vt:variant", item);
        else
            pairCategory = HeadingPairsNode.append_child("vt:variant");
        XMLNode categoryName = pairCategory.append_child("vt:lpstr");
        categoryName.text().set(name.c_str());
        XMLNode pairCount = HeadingPairsNode.insert_child_after("vt:variant", pairCategory);
        pairCountValue    = pairCount.append_child("vt:i4");
    }

    if (not pairCountValue.empty())
        pairCountValue.text().set(std::to_string(value).c_str());
    else {
        using namespace std::literals::string_literals;
        throw XLInternalError("XLAppProperties::addHeadingPair: found no matching pair count value to name "s + name);
    }
    headingPairsSize(xmlDocument().document_element()).set_value(HeadingPairsNode.child_count_of_type());
}

/**
 * @details
 */
void XLAppProperties::deleteHeadingPair(const std::string& name)
{
    if (m_xmlData == nullptr) return;

    XMLNode HeadingPairsNode = headingPairsNode(xmlDocument().document_element());
    XMLNode item             = HeadingPairsNode.first_child_of_type(pugi::node_element);
    while (not item.empty() && item.first_child_of_type(pugi::node_element).child_value() != name)
        item = item.next_sibling_of_type(pugi::node_element)
                   .next_sibling_of_type(pugi::node_element);    // advance two elements to skip count node

    // ===== If item with name was found, remove pair and update headingPairsSize
    if (not item.empty()) {
        const XMLNode count = item.next_sibling_of_type(pugi::node_element);
        // ===== 2024-05-28: delete all (non-element) nodes between item and count node, *then* delete non-element nodes following a count node
        if (not count.empty()) {
            while (item.next_sibling() != count)
                HeadingPairsNode.remove_child(item.next_sibling());     // remove nodes between item & count nodes to be deleted jointly

            // ===== Delete all non-element nodes following the count node
            while ((not count.next_sibling().empty()) && (count.next_sibling().type() != pugi::node_element))
                HeadingPairsNode.remove_child(count.next_sibling());    // remove all non-element nodes following a count node

            HeadingPairsNode.remove_child(count);
        }
        // REMOVED: formatting doesn't get prettier by removing whitespaces on both sides of a pair
        // while (item.previous_sibling().type() == pugi::node_pcdata) HeadingPairsNode.remove_child(item.previous_sibling());

        HeadingPairsNode.remove_child(item);
        headingPairsSize(xmlDocument().document_element()).set_value(HeadingPairsNode.child_count_of_type());
    }
}

/**
 * @details
 */
void XLAppProperties::setHeadingPair(const std::string& name, int newValue)
{
    if (m_xmlData == nullptr) return;

    const XMLNode HeadingPairsNode = headingPairsNode(xmlDocument().document_element());
    XMLNode       item             = HeadingPairsNode.first_child_of_type(pugi::node_element);
    while (not item.empty() && item.first_child_of_type(pugi::node_element).child_value() != name)
        item = item.next_sibling_of_type(pugi::node_element)
                   .next_sibling_of_type(pugi::node_element);    // advance two elements to skip count node

    if (not item.empty()) {
        XMLNode pairCountValue = item.next_sibling_of_type(pugi::node_element).first_child_of_type(pugi::node_element);
        using namespace std::literals::string_literals;
        if (not pairCountValue.empty() && (pairCountValue.name() == "vt:i4"s))
            pairCountValue.text().set(std::to_string(newValue).c_str());
        else
            throw XLInternalError("XLAppProperties::setHeadingPair: found no matching pair count value to name "s + name);
    }
}

/**
 * @details
 */
void XLAppProperties::setProperty(const std::string& name, const std::string& value)
{
    if (m_xmlData == nullptr) return;
    auto property = xmlDocument().document_element().child(name.c_str());
    if (property.empty()) xmlDocument().document_element().append_child(name.c_str());
    property.text().set(value.c_str());
}

/**
 * @details
 */
std::string XLAppProperties::property(const std::string& name) const
{
    if (m_xmlData == nullptr) return "";
    XMLNode property = xmlDocument().document_element().child(name.c_str());
    if (property.empty())
        property = xmlDocument().document_element().append_child(name.c_str());    // BUGFIX 2024-05-21: re-assign the newly created node to
                                                                                   // property, so that .text().get() is defined behavior
    return property.text().get();
    // NOTE 2024-05-21: this was previously defined behavior because XMLNode::text() called from an empty xml_node returns an xml_text node
    // constructed on an empty _root pointer, while in turn xml_text::get() returns a PUGIXML_TEXT("") for an empty xml_text node
    // However, relying on all this implicit functionality was really ugly ;)
}

/**
 * @details
 */
void XLAppProperties::deleteProperty(const std::string& name)
{
    if (m_xmlData == nullptr) return;
    const auto property = xmlDocument().document_element().child(name.c_str());
    if (property.empty()) return;

    xmlDocument().document_element().remove_child(property);
}

/**
 * @details
 */
void XLAppProperties::appendSheetName(const std::string& sheetName)
{
    if (m_xmlData == nullptr) return;
    auto theNode = sheetNames(xmlDocument().document_element()).append_child("vt:lpstr");
    theNode.text().set(sheetName.c_str());
    sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);
}

/**
 * @details
 */
void XLAppProperties::prependSheetName(const std::string& sheetName)
{
    if (m_xmlData == nullptr) return;
    auto theNode = sheetNames(xmlDocument().document_element()).prepend_child("vt:lpstr");
    theNode.text().set(sheetName.c_str());
    sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);
}

/**
 * @details
 */
void XLAppProperties::insertSheetName(const std::string& sheetName, unsigned int index)
{
    if (m_xmlData == nullptr) return;

    if (index <= 1) {
        prependSheetName(sheetName);
        return;
    }

    // ===== If sheet needs to be appended...
    if (index > sheetCount(xmlDocument().document_element()).as_uint()) {
        // ===== If at least one sheet node exists, apply some pretty formatting by appending new sheet name between last sheet and trailing
        // whitespaces.
        const XMLNode lastSheet = sheetNames(xmlDocument().document_element()).last_child_of_type(pugi::node_element);
        if (not lastSheet.empty()) {
            XMLNode theNode = sheetNames(xmlDocument().document_element()).insert_child_after("vt:lpstr", lastSheet);
            theNode.text().set(sheetName.c_str());
            // ===== Update sheet count before return statement.
            sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);
        }
        else
            appendSheetName(sheetName);
        return;
    }

    XMLNode  curNode = sheetNames(xmlDocument().document_element()).first_child_of_type(pugi::node_element);
    unsigned idx     = 1;
    while (not curNode.empty()) {
        if (idx == index) break;
        curNode = curNode.next_sibling_of_type(pugi::node_element);
        ++idx;
    }

    if (curNode.empty()) {
        appendSheetName(sheetName);
        return;
    }

    XMLNode theNode = sheetNames(xmlDocument().document_element()).insert_child_before("vt:lpstr", curNode);
    theNode.text().set(sheetName.c_str());

    sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);
}
