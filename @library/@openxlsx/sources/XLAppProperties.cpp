//
// Created by Troldal on 15/08/16.
//

#include <cstring>
#include <utility>
#include <pugixml.hpp>

#include "XLAppProperties.h"
#include "XLDocument.h"

using namespace std;
using namespace OpenXLSX::Impl;

/**
 * @details
 */
XLAppProperties::XLAppProperties(XLDocument &parent,
                                 const std::string &filePath)
    : XLAbstractXMLFile(parent, filePath),
      XLSpreadsheetElement(parent),
      m_sheetCountAttribute(std::make_unique<XMLAttribute>()),
      m_sheetNamesParent(std::make_unique<XMLNode>()),
      m_sheetNameNodes(),
      m_headingPairsSize(std::make_unique<XMLAttribute>()),
      m_headingPairsCategoryParent(std::make_unique<XMLNode>()),
      m_headingPairsCountParent(std::make_unique<XMLNode>()),
      m_headingPairs(),
      m_properties(),
      m_worksheetCount(0),
      m_chartsheetCount(0)
{
    ParseXMLData();
    SetModified();
}

/**
 * @details
 */
bool XLAppProperties::ParseXMLData()
{
    m_sheetNameNodes.clear();
    m_headingPairs.clear();
    m_properties.clear();

    auto node = XmlDocument()->first_child().first_child();

    while (node) {
        if (string(node.name()) == "HeadingPairs") {
            *m_headingPairsSize = node.first_child().attribute("size");
            *m_headingPairsCategoryParent = node.first_child().first_child();
            *m_headingPairsCountParent = m_headingPairsCategoryParent->next_sibling();

            for (int i = 0; i < stoi(m_headingPairsSize->value()) / 2; ++i) {
                auto categoryNode = m_headingPairsCategoryParent->first_child();
                auto countNode = m_headingPairsCountParent->first_child();

                auto element = std::make_pair(categoryNode, countNode);

                m_headingPairs.push_back(std::move(element));

                if (i < stoi(m_headingPairsSize->value()) / 2 - 1) {
                    *m_headingPairsCategoryParent = m_headingPairsCountParent->next_sibling();
                    *m_headingPairsCountParent = m_headingPairsCategoryParent->next_sibling();
                }
            }
        }
        else if (string(node.name()) == "TitlesOfParts") {
            *m_sheetCountAttribute = node.first_child().attribute("size");
            *m_sheetNamesParent = node.first_child();

            for (auto &currentNode : m_sheetNamesParent->children())
                    m_sheetNameNodes[currentNode.text().get()] = currentNode;
        }
        else {
            m_properties.insert_or_assign(node.name(), node);
        }

        node = node.next_sibling();
    }

    return true;
}

/**
 * @details
 */
XMLNode XLAppProperties::AddSheetName(const string &title)
{
    auto theNode = m_sheetNamesParent->append_child("vt:lpstr");
    theNode.set_value(title.c_str());

    m_sheetNameNodes[theNode.value()] = theNode;
    m_sheetCountAttribute->set_value(m_sheetNameNodes.size());

    SetModified();

    return m_sheetNameNodes.at(theNode.value());
}

/**
 * @details
 */
void XLAppProperties::DeleteSheetName(const string &title)
{
    for (auto &iter : m_sheetNameNodes) {
        if (iter.second.value() == title) {
            iter.second.parent().remove_child(iter.second);
            m_sheetNameNodes.erase(iter.first);
            m_sheetCountAttribute->set_value(m_sheetNameNodes.size());
            SetModified();
            return;
        }
    }
}

/**
 * @details
 */
void XLAppProperties::SetSheetName(const string &oldTitle,
                                      const string &newTitle)
{
    for (auto &iter : m_sheetNameNodes) {
        if (iter.second.value() == oldTitle) {
            iter.second.set_value(newTitle.c_str());
            SetModified();
            return;
        }
    }
}

/**
 * @details
 */
XMLNode XLAppProperties::SheetNameNode(const string &title)
{
    return m_sheetNameNodes.at(title);
}

/**
 * @details
 */
void XLAppProperties::AddHeadingPair(const string &name,
                                        int value)
{
    auto pairCategory = m_headingPairsCategoryParent->append_child("vt:lpstr");
    pairCategory.set_value(name.c_str());

    auto pairCount = m_headingPairsCountParent->append_child("vt:i4");
    pairCount.set_value(to_string(value).c_str());

    m_headingPairs.push_back(std::make_pair(pairCategory, pairCount));

    m_headingPairsSize->set_value(m_headingPairs.size());
    SetModified();
}

/**
 * @details
 */
void XLAppProperties::DeleteHeadingPair(const string &name)
{
    for (auto iter = m_headingPairs.begin(); iter != m_headingPairs.end(); iter++) {
        if (iter->first.value() == name) {
            iter->first.parent().remove_child(iter->first);
            iter->second.parent().remove_child(iter->second);
            m_headingPairs.erase(iter);
            m_headingPairsSize->set_value(m_headingPairs.size());
            SetModified();
        }
    }
}

/**
 * @details
 */
void XLAppProperties::SetHeadingPair(const string &name,
                                        int newValue)
{
    for (auto &iter : m_headingPairs) {
        if (string(iter.first.text().get()) == name) {
            iter.second.text().set(to_string(newValue).c_str());
            SetModified();
            return;
        }
    }
}

/**
 * @details
 */
bool XLAppProperties::SetProperty(const std::string &name,
                                     const std::string &value)
{
    m_properties.at(name).set_value(value.c_str());
    SetModified();
    return true;
}

/**
 * @details
 */
bool XLAppProperties::SetProperty(const std::string &name,
                                     int value)
{
    m_properties.at(name).set_value(to_string(value).c_str());
    SetModified();
    return true;
}

/**
 * @details
 */
bool XLAppProperties::SetProperty(const std::string &name,
                                     double value)
{
    m_properties.at(name).set_value(to_string(value).c_str());
    SetModified();
    return true;
}

/**
 * @details
 */
XMLNode XLAppProperties::Property(const string &name) const
{
    return m_properties.at(name);
}

/**
 * @details
 */
void XLAppProperties::DeleteProperty(const string &name)
{
    m_properties.at(name).parent().remove_child(m_properties.at(name));
    m_properties.erase(name);
    SetModified();
}

/**
 * @details
 */
XMLNode XLAppProperties::AppendSheetName(const std::string &sheetName)
{
    auto theNode = m_sheetNamesParent->append_child("vt:lpstr");
    theNode.text().set(sheetName.c_str());

    m_sheetNameNodes[sheetName] = theNode;
    m_sheetCountAttribute->set_value(m_sheetNameNodes.size());

    SetModified();

    return m_sheetNameNodes.at(sheetName);
}

/**
 * @details
 */
XMLNode XLAppProperties::PrependSheetName(const std::string &sheetName)
{
    auto theNode = m_sheetNamesParent->prepend_child("vt:lpstr");
    theNode.text().set(sheetName.c_str());

    m_sheetNameNodes[sheetName] = theNode;
    m_sheetCountAttribute->set_value(m_sheetNameNodes.size());

    SetModified();

    return m_sheetNameNodes.at(sheetName);
}

/**
 * @details
 */
XMLNode XLAppProperties::InsertSheetName(const std::string &sheetName,
                                         unsigned int index)
{
    if (index <= 1) return PrependSheetName(sheetName);
    if (index > m_sheetNameNodes.size()) return AppendSheetName(sheetName);

    auto curNode = m_sheetNamesParent->first_child();
    unsigned idx = 1;
    while (curNode) {
        if (idx == index) break;
        curNode = curNode.next_sibling();
        ++idx;
    }

    if (!curNode) return AppendSheetName(sheetName);

    auto theNode = m_sheetNamesParent->insert_child_before("vt:lpstr", curNode);
    theNode.text().set(sheetName.c_str());

    m_sheetNameNodes[sheetName] = theNode;
    m_sheetCountAttribute->set_value(m_sheetNameNodes.size());

    SetModified();

    return m_sheetNameNodes.at(sheetName);
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode XLAppProperties::MoveSheetName(const std::string &sheetName,
                                       unsigned int index)
{
    return XMLNode(); //TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode XLAppProperties::AppendWorksheetName(const std::string &sheetName)
{
    return XMLNode(); //TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode XLAppProperties::PrependWorksheetName(const std::string &sheetName)
{
    return XMLNode(); //TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode XLAppProperties::InsertWorksheetName(const std::string &sheetName,
                                             unsigned int index)
{
    return XMLNode(); //TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode XLAppProperties::MoveWorksheetName(const std::string &sheetName,
                                           unsigned int index)
{
    return XMLNode(); //TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode XLAppProperties::AppendChartsheetName(const std::string &sheetName)
{
    return XMLNode(); //TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode XLAppProperties::PrependChartsheetName(const std::string &sheetName)
{
    return XMLNode(); //TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode XLAppProperties::InsertChartsheetName(const std::string &sheetName,
                                              unsigned int index)
{
    return XMLNode(); //TODO: Dummy implementation.
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode XLAppProperties::MoveChartsheetName(const std::string &sheetName,
                                            unsigned int index)
{
    return XMLNode(); //TODO: Dummy implementation.
}
