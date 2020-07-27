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
bool XLCoreProperties::setProperty(const std::string& name, const std::string& value)
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
bool XLCoreProperties::setProperty(const std::string& name, int value)
{
    return setProperty(name, to_string(value));
}

/**
 * @details
 */
bool XLCoreProperties::setProperty(const std::string& name, double value)
{
    return setProperty(name, to_string(value));
}

/**
 * @details
 */
std::string XLCoreProperties::property(const std::string& name) const
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
void XLCoreProperties::deleteProperty(const std::string& name)
{
    auto property = xmlDocument().first_child().child(name.c_str());
    if (property != nullptr) xmlDocument().first_child().remove_child(property);
}
