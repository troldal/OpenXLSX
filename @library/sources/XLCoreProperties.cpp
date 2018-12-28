//
// Created by Troldal on 18/08/16.
//

#include "XLCoreProperties.h"
#include "XLDocument.h"

using namespace std;
using namespace OpenXLSX;

/**
 * @details
 */
XLCoreProperties::XLCoreProperties(XLDocument &parent,
                                   const std::string &filePath)
    : XLAbstractXMLFile(parent, filePath),
      XLSpreadsheetElement(parent),
      m_properties()
{
    ParseXMLData();
}

/**
 * @details
 */
XLCoreProperties::~XLCoreProperties()
{

}

/**
 * @details
 */
bool XLCoreProperties::ParseXMLData()
{
    m_properties.clear();
    for (auto &node : XmlDocument()->first_child().children()) m_properties.insert_or_assign(node.name(), node);
    return true;
}

/**
 * @details
 */
bool XLCoreProperties::SetProperty(const std::string &name,
                                   const std::string &value)
{
    m_properties.at(name).set_value(value.c_str());
    SetModified();
    return true;
}

/**
 * @details
 */
bool XLCoreProperties::SetProperty(const std::string &name,
                                   int value)
{
    return SetProperty(name, to_string(value));
}

/**
 * @details
 */
bool XLCoreProperties::SetProperty(const std::string &name,
                                      double value)
{
    return SetProperty(name, to_string(value));
}

/**
 * @details
 */
const XMLNode XLCoreProperties::Property(const std::string &name) const
{
    return m_properties.at(name);
}

/**
 * @details
 */
void XLCoreProperties::DeleteProperty(const std::string &name)
{
    auto element = m_properties.at(name);
    element.parent().remove_child(element);
    m_properties.erase(name);
    SetModified();
}

