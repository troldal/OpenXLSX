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

#include <pugixml.hpp>

// ===== External Includes ===== //
#include "XLDocument.hpp"
#include "XLProperties.hpp"

using namespace OpenXLSX;

namespace
{
    inline XMLAttribute headingPairsSize(XMLNode docNode)
    {
        return docNode.child("HeadingPairs").first_child().attribute("size");
    }

    inline XMLNode headingPairsCategories(XMLNode docNode)
    {
        return docNode.child("HeadingPairs").first_child().first_child();
    }

    inline XMLNode headingPairsCounts(XMLNode docNode)
    {
        return headingPairsCategories(docNode).next_sibling();
    }

    inline XMLNode sheetNames(XMLNode docNode)
    {
        return docNode.child("TitleOfParts").first_child();
    }

    inline XMLAttribute sheetCount(XMLNode docNode)
    {
        return sheetNames(docNode).attribute("size");
    }
}    // namespace

/**
 * @details
 */
XLProperties::XLProperties(XLXmlData* xmlData) : XLXmlFile(xmlData) {}

XLProperties::~XLProperties() = default;

/**
 * @details
 */
bool XLProperties::setProperty(const std::string& name, const std::string& value)
{
    XMLNode node;
    if (xmlDocument().first_child().child(name.c_str()) != nullptr)
        node = xmlDocument().first_child().child(name.c_str());
    else
        node = xmlDocument().first_child().prepend_child(name.c_str());

    node.text().set(value.c_str());
    return true;
}

/**
 * @details
 */
bool XLProperties::setProperty(const std::string& name, int value)
{
    return setProperty(name, std::to_string(value));
}

/**
 * @details
 */
bool XLProperties::setProperty(const std::string& name, double value)
{
    return setProperty(name, std::to_string(value));
}

/**
 * @details
 */
std::string XLProperties::property(const std::string& name) const
{
    auto property = xmlDocument().first_child().child(name.c_str());
    if (!property) {
        property = xmlDocument().first_child().append_child(name.c_str());
    }

    return property.text().get();
}

/**
 * @details
 */
void XLProperties::deleteProperty(const std::string& name)
{
    auto property = xmlDocument().first_child().child(name.c_str());
    if (property != nullptr) xmlDocument().first_child().remove_child(property);
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
    auto theNode = sheetNames(xmlDocument().document_element()).append_child("vt:lpstr");
    theNode.text().set(title.c_str());
    sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);
}

/**
 * @details
 */
void XLAppProperties::deleteSheetName(const std::string& title)
{
    for (auto& iter : sheetNames(xmlDocument().document_element()).children()) {
        if (iter.child_value() == title) {
            sheetNames(xmlDocument().document_element()).remove_child(iter);
            sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() - 1);
            return;
        }
    }
}

/**
 * @details
 */
void XLAppProperties::setSheetName(const std::string& oldTitle, const std::string& newTitle)
{
    for (auto& iter : sheetNames(xmlDocument().document_element()).children()) {
        if (iter.child_value() == oldTitle) {
            iter.text().set(newTitle.c_str());
            return;
        }
    }
}

/**
 * @details
 */
void XLAppProperties::addHeadingPair(const std::string& name, int value)
{
    for (auto& item : headingPairsCategories(xmlDocument().document_element()).children()) {
        if (item.child_value() == name) return;
    }

    auto pairCategory = headingPairsCategories(xmlDocument().document_element()).append_child("vt:lpstr");
    pairCategory.set_value(name.c_str());

    auto pairCount = headingPairsCounts(xmlDocument().document_element()).append_child("vt:i4");
    pairCount.set_value(std::to_string(value).c_str());

    headingPairsSize(xmlDocument().document_element())
        .set_value(std::distance(headingPairsCategories(xmlDocument().document_element()).begin(),
                                 headingPairsCategories(xmlDocument().document_element()).end()));
}

/**
 * @details
 */
void XLAppProperties::deleteHeadingPair(const std::string& name)
{
    auto category = headingPairsCategories(xmlDocument().document_element()).begin();
    auto count    = headingPairsCounts(xmlDocument().document_element()).begin();
    while (category != headingPairsCategories(xmlDocument().document_element()).end() &&
           count != headingPairsCounts(xmlDocument().document_element()).end())
    {
        if (category->child_value() == name) {
            headingPairsCategories(xmlDocument().document_element()).remove_child(*category);
            headingPairsCounts(xmlDocument().document_element()).remove_child(*count);
            break;
        }
        ++category;
        ++count;
    }
}

/**
 * @details
 */
void XLAppProperties::setHeadingPair(const std::string& name, int newValue)
{
    auto category = headingPairsCategories(xmlDocument().document_element()).begin();
    auto count    = headingPairsCounts(xmlDocument().document_element()).begin();
    while (category != headingPairsCategories(xmlDocument().document_element()).end() &&
           count != headingPairsCounts(xmlDocument().document_element()).end())
    {
        if (category->child_value() == name) {
            count->text().set(std::to_string(newValue).c_str());
            break;
        }
        ++category;
        ++count;
    }
}

/**
 * @details
 */
bool XLAppProperties::setProperty(const std::string& name, const std::string& value)
{
    auto property = xmlDocument().first_child().child(name.c_str());
    if (!property) xmlDocument().first_child().append_child(name.c_str());
    property.text().set(value.c_str());
    return true;
}

/**
 * @details
 */
std::string XLAppProperties::property(const std::string& name) const
{
    auto property = xmlDocument().first_child().child(name.c_str());
    if (!property) xmlDocument().first_child().append_child(name.c_str());

    return property.text().get();
}

/**
 * @details
 */
void XLAppProperties::deleteProperty(const std::string& name)
{
    auto property = xmlDocument().first_child().child(name.c_str());
    if (!property) return;

    xmlDocument().first_child().remove_child(property);
}

/**
 * @details
 */
void XLAppProperties::appendSheetName(const std::string& sheetName)
{
    auto theNode = sheetNames(xmlDocument().document_element()).append_child("vt:lpstr");
    theNode.text().set(sheetName.c_str());
    sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);
}

/**
 * @details
 */
void XLAppProperties::prependSheetName(const std::string& sheetName)
{
    auto theNode = sheetNames(xmlDocument().document_element()).prepend_child("vt:lpstr");
    theNode.text().set(sheetName.c_str());
    sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);
}

/**
 * @details
 */
void XLAppProperties::insertSheetName(const std::string& sheetName, unsigned int index)
{
    if (index <= 1) {
        prependSheetName(sheetName);
        return;
    }

    if (index > sheetCount(xmlDocument().document_element()).as_uint()) {
        appendSheetName(sheetName);
        return;
    }

    auto     curNode = sheetNames(xmlDocument().document_element()).first_child();
    unsigned idx     = 1;
    while (curNode != nullptr) {
        if (idx == index) break;
        curNode = curNode.next_sibling();
        ++idx;
    }

    if (!curNode) {
        appendSheetName(sheetName);
        return;
    }

    auto theNode = sheetNames(xmlDocument().document_element()).insert_child_before("vt:lpstr", curNode);
    theNode.text().set(sheetName.c_str());

    sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);
}