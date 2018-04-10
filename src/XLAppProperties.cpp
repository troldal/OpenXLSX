//
// Created by Troldal on 15/08/16.
//

#include "XLAppProperties.h"
#include "XLDocument.h"

using namespace std;
using namespace OpenXLSX;

/**
 * @details
 */
XLAppProperties::XLAppProperties(XLDocument &parent,
                                 const std::string &filePath)
    : XLAbstractXMLFile(parent.RootDirectory()->string(), filePath),
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
XLAppProperties::~XLAppProperties()
{

}

/**
 * @details
 */
bool XLAppProperties::ParseXMLData()
{
    m_sheetNameNodes.clear();
    m_headingPairs.clear();
    m_properties.clear();

    auto node = XmlDocument()->FirstNode();

    while (node) {
        if (node->Name() == "HeadingPairs") {
            m_headingPairsSize = node->ChildNode()->Attribute("size");
            m_headingPairsCategoryParent = node->ChildNode()->ChildNode();
            m_headingPairsCountParent = m_headingPairsCategoryParent->NextSibling();

            for (int i = 0; i < stoi(m_headingPairsSize->Value()) / 2; ++i) {
                auto categoryNode = m_headingPairsCategoryParent->ChildNode();
                auto countNode = m_headingPairsCountParent->ChildNode();

                auto element = std::make_pair(categoryNode, countNode);

                m_headingPairs.push_back(element);

                if (i < stoi(m_headingPairsSize->Value()) / 2 - 1) {
                    m_headingPairsCategoryParent = m_headingPairsCountParent->NextSibling();
                    m_headingPairsCountParent = m_headingPairsCategoryParent->NextSibling();
                }
            }
        }
        else if (node->Name() == "TitlesOfParts") {
            m_sheetCountAttribute = node->ChildNode()->Attribute("size");
            m_sheetNamesParent = node->ChildNode();

            auto currentNode = m_sheetNamesParent->ChildNode();
            for (int i = 0; i < stoi(m_sheetCountAttribute->Value()); ++i) {
                string temp = currentNode->Value();
                m_sheetNameNodes[currentNode->Value()] = currentNode;
                currentNode = currentNode->NextSibling();
            }
        }
        else {
            m_properties.insert_or_assign(node->Name(), node);
        }

        node = node->NextSibling();
    }


    return true;
}

/**
 * @details
 */
XMLNode *XLAppProperties::AddSheetName(const string &title)
{
    XMLNode *theNode = XmlDocument()->CreateNode("vt:lpstr");
    m_sheetNamesParent->AppendNode(theNode);
    theNode->SetValue(title);

    m_sheetNameNodes[theNode->Value()] = theNode;
    m_sheetCountAttribute->SetValue((int) m_sheetNameNodes.size());

    SetModified();

    return theNode;
}

/**
 * @details
 */
void XLAppProperties::DeleteSheetName(const string &title)
{
    for (auto iter = m_sheetNameNodes.begin(); iter != m_sheetNameNodes.end(); ++iter) {
        if ((*iter).second->Value() == title) {
            (*iter).second->DeleteNode();
            m_sheetNameNodes.erase(iter);
            m_sheetCountAttribute->SetValue((int) m_sheetNameNodes.size());
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
    for (auto iter = m_sheetNameNodes.begin(); iter != m_sheetNameNodes.end(); ++iter) {
        if ((*iter).second->Value() == oldTitle) {
            (*iter).second->SetValue(newTitle);
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
    return m_sheetNameNodes.at(title);
}

/**
 * @details
 */
void XLAppProperties::AddHeadingPair(const string &name,
                                        int value)
{
    XMLNode *pairCategory = XmlDocument()->CreateNode("vt:lpstr");
    m_headingPairsCategoryParent->AppendNode(pairCategory);
    pairCategory->SetValue(name);

    XMLNode *pairCount = XmlDocument()->CreateNode("vt:i4");
    m_headingPairsCountParent->AppendNode(pairCount);
    pairCount->SetValue(value);

    m_headingPairs.push_back(std::make_pair(pairCategory, pairCount));

    m_headingPairsSize->SetValue((int) m_headingPairs.size());
    SetModified();
}

/**
 * @details
 */
void XLAppProperties::DeleteHeadingPair(const string &name)
{
    for (auto iter = m_headingPairs.begin(); iter != m_headingPairs.end(); ++iter) {
        if (iter->first->Value() == name) {
            iter->first->DeleteNode();
            iter->second->DeleteNode();
            m_headingPairs.erase(iter);
            m_headingPairsSize->SetValue((int) m_headingPairs.size());
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
    for (auto iter = m_headingPairs.begin(); iter != m_headingPairs.end(); ++iter) {
        if (iter->first->Value() == name) {
            iter->second->SetValue(newValue);
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
    m_properties.at(name)->SetValue(value);
    SetModified();
    return true;
}

/**
 * @details
 */
bool XLAppProperties::SetProperty(const std::string &name,
                                     int value)
{
    m_properties.at(name)->SetValue(value);
    SetModified();
    return true;
}

/**
 * @details
 */
bool XLAppProperties::SetProperty(const std::string &name,
                                     double value)
{
    m_properties.at(name)->SetValue(value);
    SetModified();
    return true;
}

/**
 * @details
 */
XMLNode *XLAppProperties::Property(const string &name) const
{
    return m_properties.at(name);
}

/**
 * @details
 */
void XLAppProperties::DeleteProperty(const string &name)
{
    auto element = m_properties.at(name);
    element->DeleteNode();
    m_properties.erase(name);
    SetModified();
}

/**
 * @details
 */
XMLNode *XLAppProperties::AppendSheetName(const std::string &sheetName)
{
    XMLNode *theNode = XmlDocument()->CreateNode("vt:lpstr");
    theNode->SetValue(sheetName);

    m_sheetNamesParent->AppendNode(theNode);
    m_sheetNameNodes[sheetName] = theNode;
    m_sheetCountAttribute->SetValue((int) m_sheetNameNodes.size());

    SetModified();

    return theNode;
}

/**
 * @details
 */
XMLNode *XLAppProperties::PrependSheetName(const std::string &sheetName)
{
    XMLNode *theNode = XmlDocument()->CreateNode("vt:lpstr");
    theNode->SetValue(sheetName);

    m_sheetNamesParent->PrependNode(theNode);
    m_sheetNameNodes[sheetName] = theNode;
    m_sheetCountAttribute->SetValue((int) m_sheetNameNodes.size());

    SetModified();

    return theNode;
}

/**
 * @details
 */
XMLNode *XLAppProperties::InsertSheetName(const std::string &sheetName,
                                           unsigned int index)
{
    if (index <= 1) return PrependSheetName(sheetName);

    XMLNode *theNode = XmlDocument()->CreateNode("vt:lpstr");
    theNode->SetValue(sheetName);

    XMLNode *curNode = m_sheetNamesParent->ChildNode();
    unsigned idx = 1;

    while (curNode) {
        if (idx == index) break;

        curNode = curNode->NextSibling();
        ++idx;
    }

    if (!curNode) return AppendSheetName(sheetName);

    m_sheetNamesParent->InsertNode(curNode, theNode);
    m_sheetNameNodes[sheetName] = theNode;
    m_sheetCountAttribute->SetValue((int) m_sheetNameNodes.size());

    SetModified();

    return theNode;
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
