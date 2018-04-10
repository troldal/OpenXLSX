//
// Created by Troldal on 28/08/16.
//

#include "XMLNode.h"

using namespace OpenXLSX;
using namespace rapidxml;
using namespace std;

/**
 * @details
 */
XMLNode::XMLNode(XMLDocument *parentDocument,
                 rapidxml::xml_node<> *node)
    : m_node(node),
      m_parentXMLDocument(parentDocument),
      m_valueCache(""),
      m_valueLoaded(false),
      m_nameCache(""),
      m_nameLoaded(false)
{
}

/**
 * @details
 */
bool XMLNode::isValid() const
{
    if (m_node == nullptr) {
        return false;
    }
    else {
        return true;
    }
}

/**
 * @details
 */
void XMLNode::setName(const std::string &name)
{
    auto theName = m_parentXMLDocument->m_document->allocate_string(name.c_str());
    m_node->name(theName);
    m_nameCache = name;
    m_nameLoaded = true;
}

/**
 * @details
 */
const string &XMLNode::name() const
{
    if (!m_nameLoaded) {
        m_nameCache = string(m_node->name());
        m_nameLoaded = true;
    }
    return m_nameCache;
}

/**
 * @details
 */
void XMLNode::setValue(const std::string &value)
{
    auto theValue = m_parentXMLDocument->m_document->allocate_string(value.c_str());
    m_node->value(theValue);
    m_valueCache = value;
    m_valueLoaded = true;
}

/**
 * @details
 */
void XMLNode::setValue(int value)
{
    setValue(to_string(value));
}

/**
 * @details
 */
void XMLNode::setValue(double value)
{
    setValue(to_string(value));
}

/**
 * @details
 */
const string &XMLNode::value() const
{
    if (!m_valueLoaded) {
        m_valueCache = string(m_node->value());
        m_valueLoaded = true;
    }
    return m_valueCache;
}

/**
 * @details
 */
XMLNode *XMLNode::childNode(const std::string &name)
{
    xml_node<> *node = nullptr;

    if (name == "") {
        node = m_node->first_node();
    }
    else {
        node = m_node->first_node(name.c_str());
    }

    if (node)
        return m_parentXMLDocument->getNode(node);
    else
        return nullptr;
}

/**
 * @details
 */
const XMLNode *XMLNode::childNode(const std::string &name) const
{
    xml_node<> *node = nullptr;

    if (name == "") {
        node = m_node->first_node();
    }
    else {
        node = m_node->first_node(name.c_str());
    }

    if (node)
        return m_parentXMLDocument->getNode(node);
    else
        return nullptr;
}

/**
 * @details
 */
XMLNode *XMLNode::parent() const
{
    xml_node<> *node = m_node->parent();
    if (node)
        return m_parentXMLDocument->getNode(node);
    else
        return nullptr;
}


/**
 * @details
 */
XMLNode *XMLNode::previousSibling() const
{
    xml_node<> *node = m_node->previous_sibling();

    if (node)
        return m_parentXMLDocument->getNode(node);
    else
        return nullptr;
}

/**
 * @details
 */
XMLNode *XMLNode::nextSibling() const
{
    xml_node<> *node = m_node->next_sibling();

    if (node)
        return m_parentXMLDocument->getNode(node);
    else
        return nullptr;
}

/**
 * @details
 */
void XMLNode::prependNode(XMLNode *node)
{
    m_node->prepend_node(node->m_node);
}

/**
 * @details
 */
void XMLNode::appendNode(XMLNode *node)
{
    m_node->append_node(node->m_node);
}

/**
 * @details
 */
void XMLNode::insertNode(XMLNode *location, XMLNode *node)
{
    m_node->insert_node(location->m_node, node->m_node);
}

/**
 * @details
 */
void XMLNode::moveNodeUp(XMLNode *node)
{

}

/**
 * @details
 */
void XMLNode::moveNodeDown(XMLNode *node)
{

}

/**
 * @details
 */
void XMLNode::moveNodeTo(XMLNode *node, unsigned int index)
{

}

/**
 * @details
 */
void XMLNode::deleteNode()
{
    m_node->parent()->remove_node(m_node);
    m_parentXMLDocument->m_xmlNodes.erase(m_node);
}

/**
 * @details
 */
void XMLNode::deleteChildNodes()
{
    m_node->remove_all_nodes();
    // TODO: What about the pointers in the XmlDocument?
}

/**
 * @details
 */
XMLAttribute *XMLNode::attribute(const std::string &name)
{
    xml_attribute<> *attribute = nullptr;

    if (name.empty()) {
        attribute = m_node->first_attribute();
    }
    else {
        attribute = m_node->first_attribute(name.c_str());
    }

    if (attribute)
        return m_parentXMLDocument->getAttribute(attribute);
    else
        return nullptr;
}

/**
 * @details
 */
const XMLAttribute *XMLNode::attribute(const std::string &name) const
{
    xml_attribute<> *attribute = nullptr;

    if (name == "") {
        attribute = m_node->first_attribute();
    }
    else {
        attribute = m_node->first_attribute(name.c_str());
    }

    if (attribute)
        return m_parentXMLDocument->getAttribute(attribute);
    else
        return nullptr;
}

/**
 * @details
 */
void XMLNode::prependAttribute(XMLAttribute *attribute)
{
    m_node->prepend_attribute(attribute->m_attribute);
}

/**
 * @details
 */
void XMLNode::appendAttribute(XMLAttribute *attribute)
{
    m_node->append_attribute(attribute->m_attribute);
}
