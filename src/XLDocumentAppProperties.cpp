//
// Created by Troldal on 15/08/16.
//

#include "XLDocumentAppProperties.h"
#include "XLDocument.h"

using namespace std;
using namespace RapidXLSX;

/**
 * @details
 */
XLDocAppProperties::XLDocAppProperties(XLDocument &parent,
                                       const std::string &filePath)
    : XLAbstractXMLFile(parent.RootDirectory().string(), filePath),
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
    LoadXMLData();
    SetModified();
}

/**
 * @details
 */
XLDocAppProperties::~XLDocAppProperties()
{

}

/**
 * @details
 */
bool XLDocAppProperties::ParseXMLData()
{
    m_sheetNameNodes.clear();
    m_headingPairs.clear();
    m_properties.clear();

    auto node = XmlDocument().firstNode();

    while (node) {
        if (node->name() == "HeadingPairs") {
            m_headingPairsSize = node->childNode()->attribute("size");
            m_headingPairsCategoryParent = node->childNode()->childNode();
            m_headingPairsCountParent = m_headingPairsCategoryParent->nextSibling();

            for (int i = 0; i < stoi(m_headingPairsSize->value()) / 2; ++i) {
                auto categoryNode = m_headingPairsCategoryParent->childNode();
                auto countNode = m_headingPairsCountParent->childNode();

                auto element = std::make_pair(categoryNode, countNode);

                m_headingPairs.push_back(element);

                if (i < stoi(m_headingPairsSize->value()) / 2 - 1) {
                    m_headingPairsCategoryParent = m_headingPairsCountParent->nextSibling();
                    m_headingPairsCountParent = m_headingPairsCategoryParent->nextSibling();
                }
            }
        }
        else if (node->name() == "TitlesOfParts") {
            m_sheetCountAttribute = node->childNode()->attribute("size");
            m_sheetNamesParent = node->childNode();

            auto currentNode = m_sheetNamesParent->childNode();
            for (int i = 0; i < stoi(m_sheetCountAttribute->value()); ++i) {
                string temp = currentNode->value();
                m_sheetNameNodes[currentNode->value()] = currentNode;
                currentNode = currentNode->nextSibling();
            }
        }
        else {
            m_properties.insert_or_assign(node->name(), node);
        }

        node = node->nextSibling();
    }


    return true;
}

/**
 * @details
 */
XMLNode &XLDocAppProperties::AddSheetName(const string &title)
{
    XMLNode *theNode = XmlDocument().createNode("vt:lpstr");
    m_sheetNamesParent->appendNode(theNode);
    theNode->setValue(title);

    m_sheetNameNodes[theNode->value()] = theNode;
    m_sheetCountAttribute->setValue((int) m_sheetNameNodes.size());

    SetModified();

    return *theNode;
}

/**
 * @details
 */
void XLDocAppProperties::DeleteSheetName(const string &title)
{
    for (auto iter = m_sheetNameNodes.begin(); iter != m_sheetNameNodes.end(); ++iter) {
        if ((*iter).second->value() == title) {
            (*iter).second->deleteNode();
            m_sheetNameNodes.erase(iter);
            m_sheetCountAttribute->setValue((int) m_sheetNameNodes.size());
            SetModified();
            return;
        }
    }
}

/**
 * @details
 */
void XLDocAppProperties::SetSheetName(const string &oldTitle,
                                      const string &newTitle)
{
    for (auto iter = m_sheetNameNodes.begin(); iter != m_sheetNameNodes.end(); ++iter) {
        if ((*iter).second->value() == oldTitle) {
            (*iter).second->setValue(newTitle);
            SetModified();
            return;
        }
    }
}

/**
 * @details
 */
XMLNode &XLDocAppProperties::SheetNameNode(const string &title)
{
    return *m_sheetNameNodes.at(title);
}

/**
 * @details
 */
void XLDocAppProperties::AddHeadingPair(const string &name,
                                        int value)
{
    XMLNode *pairCategory = XmlDocument().createNode("vt:lpstr");
    m_headingPairsCategoryParent->appendNode(pairCategory);
    pairCategory->setValue(name);

    XMLNode *pairCount = XmlDocument().createNode("vt:i4");
    m_headingPairsCountParent->appendNode(pairCount);
    pairCount->setValue(value);

    m_headingPairs.push_back(std::make_pair(pairCategory, pairCount));

    m_headingPairsSize->setValue((int) m_headingPairs.size());
    SetModified();
}

/**
 * @details
 */
void XLDocAppProperties::DeleteHeadingPair(const string &name)
{
    for (auto iter = m_headingPairs.begin(); iter != m_headingPairs.end(); ++iter) {
        if (iter->first->value() == name) {
            iter->first->deleteNode();
            iter->second->deleteNode();
            m_headingPairs.erase(iter);
            m_headingPairsSize->setValue((int) m_headingPairs.size());
            SetModified();
        }
    }
}

/**
 * @details
 */
void XLDocAppProperties::SetHeadingPair(const string &name,
                                        int newValue)
{
    for (auto iter = m_headingPairs.begin(); iter != m_headingPairs.end(); ++iter) {
        if (iter->first->value() == name) {
            iter->second->setValue(newValue);
            SetModified();
            return;
        }
    }
}

/**
 * @details
 */
bool XLDocAppProperties::SetProperty(const std::string &name,
                                     const std::string &value)
{
    m_properties.at(name)->setValue(value);
    SetModified();
    return true;
}

/**
 * @details
 */
bool XLDocAppProperties::SetProperty(const std::string &name,
                                     int value)
{
    m_properties.at(name)->setValue(value);
    SetModified();
    return true;
}

/**
 * @details
 */
bool XLDocAppProperties::SetProperty(const std::string &name,
                                     double value)
{
    m_properties.at(name)->setValue(value);
    SetModified();
    return true;
}

/**
 * @details
 */
XMLNode &XLDocAppProperties::Property(const string &name) const
{
    return *m_properties.at(name);
}

/**
 * @details
 */
void XLDocAppProperties::DeleteProperty(const string &name)
{
    auto element = m_properties.at(name);
    element->deleteNode();
    m_properties.erase(name);
    SetModified();
}

/**
 * @details
 */
XMLNode &XLDocAppProperties::AppendSheetName(const std::string &sheetName)
{
    XMLNode *theNode = XmlDocument().createNode("vt:lpstr");
    theNode->setValue(sheetName);

    m_sheetNamesParent->appendNode(theNode);
    m_sheetNameNodes[sheetName] = theNode;
    m_sheetCountAttribute->setValue((int) m_sheetNameNodes.size());

    SetModified();

    return *theNode;
}

/**
 * @details
 */
XMLNode &XLDocAppProperties::PrependSheetName(const std::string &sheetName)
{
    XMLNode *theNode = XmlDocument().createNode("vt:lpstr");
    theNode->setValue(sheetName);

    m_sheetNamesParent->prependNode(theNode);
    m_sheetNameNodes[sheetName] = theNode;
    m_sheetCountAttribute->setValue((int) m_sheetNameNodes.size());

    SetModified();

    return *theNode;
}

/**
 * @details
 */
XMLNode &XLDocAppProperties::InsertSheetName(const std::string &sheetName,
                                             unsigned int index)
{
    if (index <= 1) return PrependSheetName(sheetName);

    XMLNode *theNode = XmlDocument().createNode("vt:lpstr");
    theNode->setValue(sheetName);

    XMLNode *curNode = m_sheetNamesParent->childNode();
    unsigned idx = 1;

    while (curNode) {
        if (idx == index) break;

        curNode = curNode->nextSibling();
        ++idx;
    }

    if (!curNode) return AppendSheetName(sheetName);

    m_sheetNamesParent->insertNode(curNode, theNode);
    m_sheetNameNodes[sheetName] = theNode;
    m_sheetCountAttribute->setValue((int) m_sheetNameNodes.size());

    SetModified();

    return *theNode;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLDocAppProperties::MoveSheetName(const std::string &sheetName,
                                           unsigned int index)
{
    return nullptr;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLDocAppProperties::AppendWorksheetName(const std::string &sheetName)
{
    return nullptr;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLDocAppProperties::PrependWorksheetName(const std::string &sheetName)
{
    return nullptr;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLDocAppProperties::InsertWorksheetName(const std::string &sheetName,
                                                 unsigned int index)
{
    return nullptr;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLDocAppProperties::MoveWorksheetName(const std::string &sheetName,
                                               unsigned int index)
{
    return nullptr;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLDocAppProperties::AppendChartsheetName(const std::string &sheetName)
{
    return nullptr;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLDocAppProperties::PrependChartsheetName(const std::string &sheetName)
{
    return nullptr;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLDocAppProperties::InsertChartsheetName(const std::string &sheetName,
                                                  unsigned int index)
{
    return nullptr;
}

/**
 * @details
 * @todo Not yet implemented
 */
XMLNode *XLDocAppProperties::MoveChartsheetName(const std::string &sheetName,
                                                unsigned int index)
{
    return nullptr;
}
