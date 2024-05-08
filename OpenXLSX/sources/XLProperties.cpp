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
        // TODO: test this again
        std::vector<std::string> result;
        XMLNode                  item = headingPairsNode(docNode).first_child_of_type(pugi::node_element);
        while (item) {
            result.push_back(item.first_child_of_type(pugi::node_element).child_value());
            item = item.next_sibling_of_type(pugi::node_element)
                       .next_sibling_of_type(pugi::node_element);    // advance two elements to skip count node
        }
        return std::move(result);
    }

    XMLNode sheetNames(XMLNode docNode) { return docNode.child("TitlesOfParts").first_child_of_type(pugi::node_element); }

    XMLAttribute sheetCount(XMLNode docNode) { return sheetNames(docNode).attribute("size"); }
}    // namespace

/**
 * @details
 */
XLProperties::XLProperties(XLXmlData* xmlData) : XLXmlFile(xmlData) {}

/**
 * @details
 */
XLProperties::~XLProperties() = default;

/**
 * @details
 */
void XLProperties::setProperty(const std::string& name, const std::string& value)
{
    if (!m_xmlData) return;
    XMLNode node = xmlDocument().document_element().child(name.c_str());
    if (node.empty())
        node = xmlDocument().document_element().append_child(
            name.c_str());    // changed this to append, to be in line with ::property behavior, TBD whether this should stay prepend

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
    if (!m_xmlData) return "";
    XMLNode property = xmlDocument().document_element().child(name.c_str());
    if (property.empty()) property = xmlDocument().document_element().append_child(name.c_str());

    return property.text().get();
}

/**
 * @details
 */
void XLProperties::deleteProperty(const std::string& name)
{
    if (!m_xmlData) return;
    if (const XMLNode property = xmlDocument().document_element().child(name.c_str()); not property.empty())
        xmlDocument().document_element().remove_child(property);
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
    if (!m_xmlData) return;
    const XMLNode theNode = sheetNames(xmlDocument().document_element()).append_child("vt:lpstr");
    theNode.text().set(title.c_str());
    sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);
}

/**
 * @details
 */
void XLAppProperties::deleteSheetName(const std::string& title)
{
    if (!m_xmlData) return;
    XMLNode iter = sheetNames(xmlDocument().document_element()).first_child_of_type(pugi::node_element);
    while (iter) {
        if (iter.child_value() == title) {
            sheetNames(xmlDocument().document_element()).remove_child(iter);
            sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() - 1);
            return;
        }
        iter = iter.next_sibling_of_type(pugi::node_element);
    }
}

/**
 * @details
 */
void XLAppProperties::setSheetName(const std::string& oldTitle, const std::string& newTitle)
{
    if (!m_xmlData) return;
    XMLNode iter = sheetNames(xmlDocument().document_element()).first_child_of_type(pugi::node_element);
    while (iter) {
        if (iter.child_value() == oldTitle) {
            iter.text().set(newTitle.c_str());
            return;
        }
        iter = iter.next_sibling_of_type(pugi::node_element);
    }
}

/**
 * @details
 */
void XLAppProperties::addHeadingPair(const std::string& name, int value)
{
    if (!m_xmlData) return;

    XMLNode HeadingPairsNode = headingPairsNode(xmlDocument().document_element());
    XMLNode item             = HeadingPairsNode.first_child_of_type(pugi::node_element);
    while (!item.empty() && item.first_child_of_type(pugi::node_element).child_value() != name)
        item = item.next_sibling_of_type(pugi::node_element)
                   .next_sibling_of_type(pugi::node_element);    // advance two elements to skip count node

    XMLNode pairCategory = item;    // could be an empty node
    XMLNode pairCountValue {};      // initialize to empty node
    if (!pairCategory.empty())
        pairCountValue = pairCategory.next_sibling_of_type(pugi::node_element).first_child_of_type(pugi::node_element);
    else {
        item = HeadingPairsNode.last_child_of_type(pugi::node_element);
        if (item)
            pairCategory = HeadingPairsNode.insert_child_after("vt:variant", item);
        else
            pairCategory = HeadingPairsNode.append_child("vt:variant");
        const XMLNode categoryName = pairCategory.append_child("vt:lpstr");
        categoryName.text().set(name.c_str());
        XMLNode pairCount = HeadingPairsNode.insert_child_after("vt:variant", pairCategory);
        pairCountValue    = pairCount.append_child("vt:i4");
    }

    if (pairCountValue)
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
    if (!m_xmlData) return;

    XMLNode HeadingPairsNode = headingPairsNode(xmlDocument().document_element());
    XMLNode item             = HeadingPairsNode.first_child_of_type(pugi::node_element);
    while (item && item.first_child_of_type(pugi::node_element).child_value() != name)
        item = item.next_sibling_of_type(pugi::node_element)
                   .next_sibling_of_type(pugi::node_element);    // advance two elements to skip count node

    // ===== If item with name was found, remove pair and update headingPairsSize
    if (!item.empty()) {
        if (const XMLNode count = item.next_sibling_of_type(pugi::node_element)) {
            if (!count.empty()) {    // 2024-05-02: TBD that change to loop deleting whitespaces following a count node is bug-free
                while (item.next_sibling() != count)
                    HeadingPairsNode.remove_child(item.next_sibling());    // remove all nodes between element nodes to be deleted jointly
                HeadingPairsNode.remove_child(count);
            }
        }
        // ===== Apply some pretty formatting by removing all whitespace nodes before pair name.
        while (item.previous_sibling().type() == pugi::node_pcdata) HeadingPairsNode.remove_child(item.previous_sibling());

        HeadingPairsNode.remove_child(item);
        headingPairsSize(xmlDocument().document_element()).set_value(HeadingPairsNode.child_count_of_type());
    }
}

/**
 * @details
 */
void XLAppProperties::setHeadingPair(const std::string& name, int newValue)
{
    if (!m_xmlData) return;

    const XMLNode HeadingPairsNode = headingPairsNode(xmlDocument().document_element());
    XMLNode       item             = HeadingPairsNode.first_child_of_type(pugi::node_element);
    while (item && item.first_child_of_type(pugi::node_element).child_value() != name)
        item = item.next_sibling_of_type(pugi::node_element)
                   .next_sibling_of_type(pugi::node_element);    // advance two elements to skip count node

    if (item) {
        const XMLNode pairCountValue = item.next_sibling_of_type(pugi::node_element).first_child_of_type(pugi::node_element);    //
        using namespace std::literals::string_literals;
        if (pairCountValue && (pairCountValue.name() == "vt:i4"s))
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
    if (!m_xmlData) return;
    const auto property = xmlDocument().document_element().child(name.c_str());
    if (!property) xmlDocument().document_element().append_child(name.c_str());
    property.text().set(value.c_str());
}

/**
 * @details
 */
std::string XLAppProperties::property(const std::string& name) const
{
    if (!m_xmlData) return "";
    const auto property = xmlDocument().document_element().child(name.c_str());
    if (!property) xmlDocument().document_element().append_child(name.c_str());

    return property.text().get();
}

/**
 * @details
 */
void XLAppProperties::deleteProperty(const std::string& name)
{
    if (!m_xmlData) return;
    const auto property = xmlDocument().document_element().child(name.c_str());
    if (!property) return;

    xmlDocument().document_element().remove_child(property);
}

/**
 * @details
 */
void XLAppProperties::appendSheetName(const std::string& sheetName)
{
    if (!m_xmlData) return;
    const auto theNode = sheetNames(xmlDocument().document_element()).append_child("vt:lpstr");
    theNode.text().set(sheetName.c_str());
    sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);
}

/**
 * @details
 */
void XLAppProperties::prependSheetName(const std::string& sheetName)
{
    if (!m_xmlData) return;
    const auto theNode = sheetNames(xmlDocument().document_element()).prepend_child("vt:lpstr");
    theNode.text().set(sheetName.c_str());
    sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);
}

/**
 * @details
 */
void XLAppProperties::insertSheetName(const std::string& sheetName, unsigned int index)
{
    if (!m_xmlData) return;

    if (index <= 1) {
        prependSheetName(sheetName);
        return;
    }

    // ===== If sheet needs to be appended...
    if (index > sheetCount(xmlDocument().document_element()).as_uint()) {
        // ===== If at least one sheet node exists, apply some pretty formatting by appending new sheet name between last sheet and trailing
        // whitespaces.
        const XMLNode lastSheet = sheetNames(xmlDocument().document_element()).last_child_of_type(pugi::node_element);
        if (lastSheet) {
            const XMLNode theNode = sheetNames(xmlDocument().document_element()).insert_child_after("vt:lpstr", lastSheet);
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
    while (curNode) {
        if (idx == index) break;
        curNode = curNode.next_sibling_of_type(pugi::node_element);
        ++idx;
    }

    if (!curNode) {
        appendSheetName(sheetName);
        return;
    }

    const XMLNode theNode = sheetNames(xmlDocument().document_element()).insert_child_before("vt:lpstr", curNode);
    theNode.text().set(sheetName.c_str());

    sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);
}
