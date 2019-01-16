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
Impl::XLCoreProperties::XLCoreProperties(XLDocument& parent,
                                         const std::string& filePath)
        : XLAbstractXMLFile(parent,
                            filePath),
          XLSpreadsheetElement(parent),
          m_properties() {

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

    m_properties.clear();
    for (auto& node : XmlDocument()->first_child().children())
        m_properties.insert_or_assign(node.name(),
                                      node);
    return true;
}

/**
 * @details
 */
bool Impl::XLCoreProperties::SetProperty(const std::string& name,
                                         const std::string& value) {

    XMLNode node;
    if (m_xmlDocument->first_child().child(name.c_str()))
        node = m_xmlDocument->first_child().child(name.c_str());
    else
        node = m_xmlDocument->first_child().prepend_child(name.c_str());

    node.text().set(value.c_str());
    m_properties[name] = node;
    SetModified();
    return true;
}

/**
 * @details
 */
bool Impl::XLCoreProperties::SetProperty(const std::string& name,
                                         int value) {

    return SetProperty(name,
                       to_string(value));
}

/**
 * @details
 */
bool Impl::XLCoreProperties::SetProperty(const std::string& name,
                                         double value) {

    return SetProperty(name,
                       to_string(value));
}

/**
 * @details
 */
const XMLNode Impl::XLCoreProperties::Property(const std::string& name) const {

    auto result = m_properties.find(name);
    if (result == m_properties.end())
        return XMLNode();
    else
        return m_properties.at(name);
}

/**
 * @details
 */
void Impl::XLCoreProperties::DeleteProperty(const std::string& name) {

    auto element = m_properties.at(name);
    element.parent().remove_child(element);
    m_properties.erase(name);
    SetModified();
}

