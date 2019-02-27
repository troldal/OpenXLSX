//
// Created by Troldal on 18/08/16.
//

#include <pugixml.hpp>

#include "XLCoreProperties_Impl.h"
#include "XLDocument_Impl.h"

using namespace std;
using namespace OpenXLSX;

/**
 * @details
 */
Impl::XLCoreProperties::XLCoreProperties(XLDocument& parent, const std::string& filePath)
        : XLAbstractXMLFile(parent, filePath) {

    ParseXMLData();
}

/**
 * @details
 */
Impl::XLCoreProperties::~XLCoreProperties() {

}

/**
 * @details
 */
bool Impl::XLCoreProperties::ParseXMLData() {

    return true;
}

/**
 * @details
 */
bool Impl::XLCoreProperties::SetProperty(const std::string& name, const std::string& value) {

    XMLNode node;
    if (XmlDocument()->first_child().child(name.c_str()) != nullptr)
        node = XmlDocument()->first_child().child(name.c_str());
    else
        node = XmlDocument()->first_child().prepend_child(name.c_str());

    node.text().set(value.c_str());
    return true;
}

/**
 * @details
 */
bool Impl::XLCoreProperties::SetProperty(const std::string& name, int value) {

    return SetProperty(name, to_string(value));
}

/**
 * @details
 */
bool Impl::XLCoreProperties::SetProperty(const std::string& name, double value) {

    return SetProperty(name, to_string(value));
}

/**
 * @details
 */
const XMLNode Impl::XLCoreProperties::Property(const std::string& name) const {

    auto property = XmlDocument()->first_child().child(name.c_str());
    if (!property) {
        property = XmlDocument()->first_child().append_child(name.c_str());
    }

    return property;
}

/**
 * @details
 */
void Impl::XLCoreProperties::DeleteProperty(const std::string& name) {

    auto property = XmlDocument()->first_child().child(name.c_str());
    if (property != nullptr)
        XmlDocument()->first_child().remove_child(property);

}

