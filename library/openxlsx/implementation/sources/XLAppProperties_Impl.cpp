//
// Created by Troldal on 15/08/16.
//

#include <cstring>
#include <utility>
#include <iterator>
#include <pugixml.hpp>

#include "XLAppProperties_Impl.h"
#include "XLDocument_Impl.h"

using namespace std;
using namespace OpenXLSX;

/**
 * @details
 */
Impl::XLAppProperties::XLAppProperties(XLDocument& parent, const std::string& filePath)
        : XLAbstractXMLFile(parent, filePath),
          m_sheetCountAttribute(XMLAttribute()),
          m_sheetNamesParent(XMLNode()),
          m_headingPairsSize(XMLAttribute()),
          m_headingPairsCategories(XMLNode()),
          m_headingPairsCounts(XMLNode()) {

    ParseXMLData();
}

/**
 * @details
 */
bool Impl::XLAppProperties::ParseXMLData() {

    auto node = XmlDocument()->first_child().first_child();

    while (node != nullptr) {
        if (string(node.name()) == "HeadingPairs") {
            m_headingPairsSize = node.first_child().attribute("size");
            m_headingPairsCategories = node.first_child().first_child();
            m_headingPairsCounts = m_headingPairsCategories.next_sibling();
        }
        else if (string(node.name()) == "TitlesOfParts") {
            m_sheetCountAttribute = node.first_child().attribute("size");
            m_sheetNamesParent = node.first_child();
        }

        node = node.next_sibling();
    }

    return true;
}

/**
 * @details
 */
XMLNode Impl::XLAppProperties::AddSheetName(const string& title) {

    auto theNode = m_sheetNamesParent.append_child("vt:lpstr");
    theNode.text().set(title.c_str());
    m_sheetCountAttribute.set_value(m_sheetCountAttribute.as_uint() + 1);

    return theNode;
}

/**
 * @details
 */
void Impl::XLAppProperties::DeleteSheetName(const string& title) {

    for (auto& iter : m_sheetNamesParent.children()) {
        if (iter.child_value() == title) {
            m_sheetNamesParent.remove_child(iter);
            m_sheetCountAttribute.set_value(m_sheetCountAttribute.as_uint() - 1);
            return;
        }
    }
}

/**
 * @details
 */
void Impl::XLAppProperties::SetSheetName(const string& oldTitle, const string& newTitle) {

    for (auto& iter : m_sheetNamesParent.children()) {
        if (iter.child_value() == oldTitle) {
            iter.text().set(newTitle.c_str());
            return;
        }
    }
}

/**
 * @details
 */
XMLNode Impl::XLAppProperties::SheetNameNode(const string& title) {

    for (auto& sheet : m_sheetNamesParent.children()) {
        if (string_view(sheet.child_value()) == title) {
            return sheet;
        }
    }

    return XMLNode();
}

/**
 * @details
 */
void Impl::XLAppProperties::AddHeadingPair(const string& name, int value) {

    for (auto& item : m_headingPairsCategories.children()) {
        if (item.child_value() == name)
            return;
    }

    auto pairCategory = m_headingPairsCategories.append_child("vt:lpstr");
    pairCategory.set_value(name.c_str());

    auto pairCount = m_headingPairsCounts.append_child("vt:i4");
    pairCount.set_value(to_string(value).c_str());

    m_headingPairsSize.set_value(std::distance(m_headingPairsCategories.begin(), m_headingPairsCategories.end()));
}

/**
 * @details
 */
void Impl::XLAppProperties::DeleteHeadingPair(const string& name) {

    auto category = m_headingPairsCategories.begin();
    auto count = m_headingPairsCounts.begin();
    while (category != m_headingPairsCategories.end() && count != m_headingPairsCounts.end()) {
        if (category->child_value() == name) {
            m_headingPairsCategories.remove_child(*category);
            m_headingPairsCounts.remove_child(*count);
            break;
        }
        ++category;
        ++count;
    }
}

/**
 * @details
 */
void Impl::XLAppProperties::SetHeadingPair(const string& name, int newValue) {

    auto category = m_headingPairsCategories.begin();
    auto count = m_headingPairsCounts.begin();
    while (category != m_headingPairsCategories.end() && count != m_headingPairsCounts.end()) {
        if (category->child_value() == name) {
            count->text().set(to_string(newValue).c_str());
            break;
        }
        ++category;
        ++count;
    }
}

/**
 * @details
 */
bool Impl::XLAppProperties::SetProperty(const std::string& name, const std::string& value) {

    Property(name).text().set(value.c_str());
    return true;
}

/**
 * @details
 */
XMLNode Impl::XLAppProperties::Property(const string& name) const {

    auto property = XmlDocument()->first_child().child(name.c_str());
    if (!property)
        XmlDocument()->first_child().append_child(name.c_str());

    return property;
}

/**
 * @details
 */
void Impl::XLAppProperties::DeleteProperty(const string& name) {

    XmlDocument()->first_child().remove_child(Property(name));
}

/**
 * @details
 */
XMLNode Impl::XLAppProperties::AppendSheetName(const std::string& sheetName) {

    auto theNode = m_sheetNamesParent.append_child("vt:lpstr");
    theNode.text().set(sheetName.c_str());
    m_sheetCountAttribute.set_value(m_sheetCountAttribute.as_uint() + 1);

    return theNode;
}

/**
 * @details
 */
XMLNode Impl::XLAppProperties::PrependSheetName(const std::string& sheetName) {

    auto theNode = m_sheetNamesParent.prepend_child("vt:lpstr");
    theNode.text().set(sheetName.c_str());
    m_sheetCountAttribute.set_value(m_sheetCountAttribute.as_uint() + 1);

    return theNode;
}

/**
 * @details
 */
XMLNode Impl::XLAppProperties::InsertSheetName(const std::string& sheetName, unsigned int index) {

    if (index <= 1)
        return PrependSheetName(sheetName);
    if (index > m_sheetCountAttribute.as_uint())
        return AppendSheetName(sheetName);

    auto curNode = m_sheetNamesParent.first_child();
    unsigned idx = 1;
    while (curNode != nullptr) {
        if (idx == index)
            break;
        curNode = curNode.next_sibling();
        ++idx;
    }

    if (!curNode)
        return AppendSheetName(sheetName);

    auto theNode = m_sheetNamesParent.insert_child_before("vt:lpstr", curNode);
    theNode.text().set(sheetName.c_str());

    m_sheetCountAttribute.set_value(m_sheetCountAttribute.as_uint() + 1);

    return theNode;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::MoveSheetName(const std::string& sheetName, unsigned int index) {

    return XMLNode(); //TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::AppendWorksheetName(const std::string& sheetName) {

    return XMLNode(); //TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::PrependWorksheetName(const std::string& sheetName) {

    return XMLNode(); //TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::InsertWorksheetName(const std::string& sheetName, unsigned int index) {

    return XMLNode(); //TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::MoveWorksheetName(const std::string& sheetName, unsigned int index) {

    return XMLNode(); //TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::AppendChartsheetName(const std::string& sheetName) {

    return XMLNode(); //TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::PrependChartsheetName(const std::string& sheetName) {

    return XMLNode(); //TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::InsertChartsheetName(const std::string& sheetName, unsigned int index) {

    return XMLNode(); //TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode Impl::XLAppProperties::MoveChartsheetName(const std::string& sheetName, unsigned int index) {

    return XMLNode(); //TODO: Dummy implementation.
}

