//
// Created by Troldal on 15/08/16.
//

#include <cstring>

#include "XLAppProperties.h"
#include "XLDocument.h"

using namespace std;
using namespace OpenXLSX;

/**
 * @details
 */
XLAppProperties::XLAppProperties(XLDocument &parent,
                                 const std::string &filePath)
    : XLAbstractXMLFile(parent, filePath),
      XLSpreadsheetElement(parent),
      m_sheetCountAttribute(nullptr),
      m_sheetNamesParent(nullptr),
      m_sheetNameNodes(),
      m_headingPairsSize(nullptr),
      m_headingPairsCategoryParent(nullptr),
      m_headingPairsCountParent(nullptr),
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

    auto node = XmlDocument()->first_child();

    while (node) {
        if (strcmp(node.name(), "HeadingPairs") == 0) {
            m_headingPairsSize = make_unique<XMLAttribute>(node.first_child().attribute("size"));
            m_headingPairsCategoryParent = make_unique<XMLNode>(node.first_child().first_child());
            m_headingPairsCountParent = make_unique<XMLNode>(m_headingPairsCategoryParent->next_sibling());

            for (int i = 0; i < stoi(m_headingPairsSize->value()) / 2; ++i) {
                auto categoryNode = m_headingPairsCategoryParent->first_child();
                auto countNode = m_headingPairsCountParent->first_child();

                auto element = std::make_pair(make_unique<XMLNode>(categoryNode), make_unique<XMLNode>(countNode));

                m_headingPairs.push_back(std::move(element));

                if (i < stoi(m_headingPairsSize->value()) / 2 - 1) {
                    m_headingPairsCategoryParent = make_unique<XMLNode>(m_headingPairsCountParent->next_sibling());
                    m_headingPairsCountParent = make_unique<XMLNode>(m_headingPairsCategoryParent->next_sibling());
                }
            }
        }
        else if (strcmp(node.name(), "TitlesOfParts") == 0) {
            m_sheetCountAttribute = make_unique<XMLAttribute>(node.first_child().attribute("size"));
            m_sheetNamesParent = make_unique<XMLNode>(node.first_child());

            auto currentNode = m_sheetNamesParent->first_child();
            for (int i = 0; i < stoi(m_sheetCountAttribute->value()); ++i) {
                m_sheetNameNodes[currentNode.value()] = make_unique<XMLNode>(currentNode);
                currentNode = currentNode.next_sibling();
            }
        }
        else {
            m_properties.insert_or_assign(node.name(), make_unique<XMLNode>(node));
        }

        node = node.next_sibling();
    }


    return true;
}

/**
 * @details
 */
XMLNode *XLAppProperties::AddSheetName(const string &title)
{
    auto theNode = m_sheetNamesParent->append_child("vt:lpstr");
    theNode.set_value(title.c_str());

    m_sheetNameNodes[theNode.value()] = make_unique<XMLNode>(theNode);
    m_sheetCountAttribute->set_value(m_sheetNameNodes.size());

    SetModified();

    return m_sheetNameNodes.at(theNode.value()).get();
}

/**
 * @details
 */
void XLAppProperties::DeleteSheetName(const string &title)
{
    for (auto &iter : m_sheetNameNodes) {
        if (iter.second->value() == title) {
            iter.second->parent().remove_child(*iter.second);
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
        if (iter.second->value() == oldTitle) {
            iter.second->set_value(newTitle.c_str());
            SetModified();
            return;
        }
    }
}

/**
 * @details
 */
XMLNode *XLAppProperties::SheetNameNode(const string &title)
{
    return m_sheetNameNodes.at(title).get();
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

    m_headingPairs.push_back(std::make_pair(make_unique<XMLNode>(pairCategory), make_unique<XMLNode>(pairCount)));

    m_headingPairsSize->set_value(m_headingPairs.size());
    SetModified();
}

/**
 * @details
 */
void XLAppProperties::DeleteHeadingPair(const string &name)
{
    for (auto iter = m_headingPairs.begin(); iter != m_headingPairs.end(); iter++) {
        if (iter->first->value() == name) {
            iter->first->parent().remove_child(*iter->first);
            iter->second->parent().remove_child(*iter->second);
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
        if (iter.first->value() == name) {
            iter.second->set_value(to_string(newValue).c_str());
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
    m_properties.at(name)->set_value(value.c_str());
    SetModified();
    return true;
}

/**
 * @details
 */
bool XLAppProperties::SetProperty(const std::string &name,
                                     int value)
{
    m_properties.at(name)->set_value(to_string(value).c_str());
    SetModified();
    return true;
}

/**
 * @details
 */
bool XLAppProperties::SetProperty(const std::string &name,
                                     double value)
{
    m_properties.at(name)->set_value(to_string(value).c_str());
    SetModified();
    return true;
}

/**
 * @details
 */
XMLNode *XLAppProperties::Property(const string &name) const
{
    return m_properties.at(name).get();
}

/**
 * @details
 */
void XLAppProperties::DeleteProperty(const string &name)
{
    m_properties.at(name)->parent().remove_child(*m_properties.at(name));
    m_properties.erase(name);
    SetModified();
}

/**
 * @details
 */
XMLNode *XLAppProperties::AppendSheetName(const std::string &sheetName)
{
    auto theNode = m_sheetNamesParent->append_child("vt:lpstr");
    theNode.set_value(sheetName.c_str());

    m_sheetNameNodes[sheetName] = make_unique<XMLNode>(theNode);
    m_sheetCountAttribute->set_value(m_sheetNameNodes.size());

    SetModified();

    return m_sheetNameNodes.at(sheetName).get();
}

/**
 * @details
 */
XMLNode *XLAppProperties::PrependSheetName(const std::string &sheetName)
{
    auto theNode = m_sheetNamesParent->prepend_child("vt:lpstr");
    theNode.set_value(sheetName.c_str());

    m_sheetNameNodes[sheetName] = make_unique<XMLNode>(theNode);
    m_sheetCountAttribute->set_value(m_sheetNameNodes.size());

    SetModified();

    return m_sheetNameNodes.at(sheetName).get();
}

/**
 * @details
 */
XMLNode *XLAppProperties::InsertSheetName(const std::string &sheetName,
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
    theNode.set_value(sheetName.c_str());

    m_sheetNameNodes[sheetName] = make_unique<XMLNode>(theNode);
    m_sheetCountAttribute->set_value(m_sheetNameNodes.size());

    SetModified();

    return m_sheetNameNodes.at(sheetName).get();
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLAppProperties::MoveSheetName(const std::string &sheetName,
                                           unsigned int index)
{
    return nullptr;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLAppProperties::AppendWorksheetName(const std::string &sheetName)
{
    return nullptr;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLAppProperties::PrependWorksheetName(const std::string &sheetName)
{
    return nullptr;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLAppProperties::InsertWorksheetName(const std::string &sheetName,
                                                 unsigned int index)
{
    return nullptr;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLAppProperties::MoveWorksheetName(const std::string &sheetName,
                                               unsigned int index)
{
    return nullptr;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLAppProperties::AppendChartsheetName(const std::string &sheetName)
{
    return nullptr;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLAppProperties::PrependChartsheetName(const std::string &sheetName)
{
    return nullptr;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLAppProperties::InsertChartsheetName(const std::string &sheetName,
                                                  unsigned int index)
{
    return nullptr;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLAppProperties::MoveChartsheetName(const std::string &sheetName,
                                                unsigned int index)
{
    return nullptr;
}
