//
// Created by Troldal on 15/08/16.
//

#include <cstring>
#include <iterator>
#include <utility>
//#include <pugixml.hpp>

#include "XLAppProperties.hpp"
#include "XLDocument.hpp"

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
XLAppProperties::XLAppProperties(XLXmlData* xmlData) : XLXmlFile(xmlData) {}

/**
 * @details
 */
XMLNode XLAppProperties::addSheetName(const std::string& title)
{
    auto theNode = sheetNames(xmlDocument().document_element()).append_child("vt:lpstr");
    theNode.text().set(title.c_str());
    sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);

    return theNode;
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
XMLNode XLAppProperties::sheetNameNode(const std::string& title)
{
    for (auto& sheet : sheetNames(xmlDocument().document_element()).children()) {
        if (std::string_view(sheet.child_value()) == title) {
            return sheet;
        }
    }

    return XMLNode();
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
    property(name).text().set(value.c_str());
    return true;
}

/**
 * @details
 */
XMLNode XLAppProperties::property(const std::string& name) const
{
    auto property = xmlDocument().first_child().child(name.c_str());
    if (!property) xmlDocument().first_child().append_child(name.c_str());

    return property;
}

/**
 * @details
 */
void XLAppProperties::deleteProperty(const std::string& name)
{
    xmlDocument().first_child().remove_child(property(name));
}

/**
 * @details
 */
XMLNode XLAppProperties::appendSheetName(const std::string& sheetName)
{
    auto theNode = sheetNames(xmlDocument().document_element()).append_child("vt:lpstr");
    theNode.text().set(sheetName.c_str());
    sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);

    return theNode;
}

/**
 * @details
 */
XMLNode XLAppProperties::prependSheetName(const std::string& sheetName)
{
    auto theNode = sheetNames(xmlDocument().document_element()).prepend_child("vt:lpstr");
    theNode.text().set(sheetName.c_str());
    sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);

    return theNode;
}

/**
 * @details
 */
XMLNode XLAppProperties::insertSheetName(const std::string& sheetName, unsigned int index)
{
    if (index <= 1) return prependSheetName(sheetName);
    if (index > sheetCount(xmlDocument().document_element()).as_uint()) return appendSheetName(sheetName);

    auto     curNode = sheetNames(xmlDocument().document_element()).first_child();
    unsigned idx     = 1;
    while (curNode != nullptr) {
        if (idx == index) break;
        curNode = curNode.next_sibling();
        ++idx;
    }

    if (!curNode) return appendSheetName(sheetName);

    auto theNode = sheetNames(xmlDocument().document_element()).insert_child_before("vt:lpstr", curNode);
    theNode.text().set(sheetName.c_str());

    sheetCount(xmlDocument().document_element()).set_value(sheetCount(xmlDocument().document_element()).as_uint() + 1);

    return theNode;
}
