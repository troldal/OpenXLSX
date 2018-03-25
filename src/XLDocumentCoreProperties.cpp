//
// Created by Troldal on 18/08/16.
//

#include "XLDocumentCoreProperties.h"
#include "XLDocument.h"

using namespace std;
using namespace RapidXLSX;

/**
 * @details
 */
XLDocCoreProperties::XLDocCoreProperties(XLDocument &parent,
                                         const std::string &filePath)
    : XLAbstractXMLFile(parent.RootDirectory().string(), filePath),
      XLSpreadsheetElement(parent),
      m_properties()
{
    LoadXMLData();
}

/**
 * @details
 */
XLDocCoreProperties::~XLDocCoreProperties()
{

}

/**
 * @details
 */
bool XLDocCoreProperties::ParseXMLData()
{
    m_properties.clear();

    auto node = XmlDocument().firstNode();

    while (node) {
        m_properties.insert_or_assign(node->name(), node);
        node = node->nextSibling();
    }

    return true;
}

/**
 * @details
 */
bool XLDocCoreProperties::SetProperty(const std::string &name,
                                      const std::string &value)
{
    m_properties.at(name)->setValue(value);
    SetModified();
    return true;
}

/**
 * @details
 */
bool XLDocCoreProperties::SetProperty(const std::string &name,
                                      int value)
{
    return SetProperty(name, to_string(value));
}

/**
 * @details
 */
bool XLDocCoreProperties::SetProperty(const std::string &name,
                                      double value)
{
    return SetProperty(name, to_string(value));
}

/**
 * @details
 */
XMLNode &XLDocCoreProperties::Property(const std::string &name) const
{
    return *m_properties.at(name);
}

/**
 * @details
 */
void XLDocCoreProperties::DeleteProperty(const std::string &name)
{
    auto element = m_properties.at(name);
    element->deleteNode();
    m_properties.erase(name);
    SetModified();
}

