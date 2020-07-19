//
// Created by Troldal on 18/08/16.
//

//#include <pugixml.hpp>

#include "XLCoreProperties.hpp"

#include "XLDocument.hpp"

using namespace std;
using namespace OpenXLSX;

/**
 * @details
 */
XLCoreProperties::XLCoreProperties(XLXmlData* xmlData) : XLXmlFile(xmlData) {}

/**
 * @details
 */
XLCoreProperties::~XLCoreProperties() = default;

/**
 * @details
 */
bool XLCoreProperties::SetProperty(const std::string& name, const std::string& value)
{
    XMLNode node;
    if (XmlDocument().first_child().child(name.c_str()) != nullptr)
        node = XmlDocument().first_child().child(name.c_str());
    else
        node = XmlDocument().first_child().prepend_child(name.c_str());

    node.text().set(value.c_str());
    return true;
}

/**
 * @details
 */
bool XLCoreProperties::SetProperty(const std::string& name, int value)
{
    return SetProperty(name, to_string(value));
}

/**
 * @details
 */
bool XLCoreProperties::SetProperty(const std::string& name, double value)
{
    return SetProperty(name, to_string(value));
}

/**
 * @details
 */
std::string XLCoreProperties::Property(const std::string& name) const
{
    auto property = XmlDocument().first_child().child(name.c_str());
    if (!property) {
        property = XmlDocument().first_child().append_child(name.c_str());
    }

    return property.text().get();
}

/**
 * @details
 */
void XLCoreProperties::DeleteProperty(const std::string& name)
{
    auto property = XmlDocument().first_child().child(name.c_str());
    if (property != nullptr) XmlDocument().first_child().remove_child(property);
}
