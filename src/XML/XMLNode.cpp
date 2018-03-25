//
// Created by Troldal on 28/08/16.
//

#include "XMLNode.h"

using namespace RapidXLSX;
using namespace rapidxml;
using namespace std;

/*
 * =============================================================================================================
 * XMLNode::XMLNode
 * =============================================================================================================
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

/*
 * =============================================================================================================
 * XMLNode::XMLNode
 * =============================================================================================================
 */
XMLNode::XMLNode(const XMLNode &other)
    : m_node(other.m_node),
      m_parentXMLDocument(other.m_parentXMLDocument),
      m_valueCache(other.m_valueCache),
      m_valueLoaded(other.m_valueLoaded),
      m_nameCache(other.m_nameCache),
      m_nameLoaded(other.m_nameLoaded)
{
}

/*
 * =============================================================================================================
 * XMLNode::operator=
 * =============================================================================================================
 */
XMLNode &XMLNode::operator=(const XMLNode &other)
{
    m_node = other.m_node;
    m_parentXMLDocument = other.m_parentXMLDocument;
    m_valueCache = other.m_valueCache;
    m_valueLoaded = other.m_valueLoaded;
    m_nameCache = other.m_nameCache;
    m_nameLoaded = other.m_nameLoaded;
    return *this;
}

/*
 * =============================================================================================================
 * XMLNode::isValid
 * =============================================================================================================
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

/*
 * =============================================================================================================
 * XMLNode::SetName
 * =============================================================================================================
 */
void XMLNode::setName(const std::string &name)
{
    auto theName = m_parentXMLDocument->m_document->allocate_string(name.c_str());
    m_node->name(theName);
    m_nameCache = name;
    m_nameLoaded = true;
}

/*
 * =============================================================================================================
 * XMLNode::Name
 * =============================================================================================================
 */
const string &XMLNode::name() const
{
    if (!m_nameLoaded) {
        m_nameCache = string(m_node->name());
        m_nameLoaded = true;
    }
    return m_nameCache;
}

/*
 * =============================================================================================================
 * XMLNode::SetValue
 * =============================================================================================================
 */
void XMLNode::setValue(const std::string &value)
{
    auto theValue = m_parentXMLDocument->m_document->allocate_string(value.c_str());
    m_node->value(theValue);
    m_valueCache = value;
    m_valueLoaded = true;
}

/*
 * =============================================================================================================
 * XMLNode::SetValue
 * =============================================================================================================
 */
void XMLNode::setValue(int value)
{
    setValue(to_string(value));
}

/*
 * =============================================================================================================
 * XMLNode::SetValue
 * =============================================================================================================
 */
void XMLNode::setValue(double value)
{
    setValue(to_string(value));
}

/*
 * =============================================================================================================
 * XMLNode::ValueAsString
 * =============================================================================================================
 */
const string &XMLNode::value() const
{
    if (!m_valueLoaded) {
        m_valueCache = string(m_node->value());
        m_valueLoaded = true;
    }
    return m_valueCache;
}

/*
 * =============================================================================================================
 * XMLNode::childNode
 * =============================================================================================================
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

/*
 * =============================================================================================================
 * XMLNode::childNode
 * =============================================================================================================
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

/*
 * =============================================================================================================
 * XMLNode::parent
 * =============================================================================================================
 */
XMLNode *XMLNode::parent() const
{
    xml_node<> *node = m_node->parent();
    if (node)
        return m_parentXMLDocument->getNode(node);
    else
        return nullptr;
}

/*
 * =============================================================================================================
 * XMLNode::previousSibling
 * =============================================================================================================
 */
XMLNode *XMLNode::previousSibling() const
{
    xml_node<> *node = m_node->previous_sibling();

    if (node)
        return m_parentXMLDocument->getNode(node);
    else
        return nullptr;
}

/*
 * =============================================================================================================
 * XMLNode::nextSibling
 * =============================================================================================================
 */
XMLNode *XMLNode::nextSibling() const
{
    xml_node<> *node = m_node->next_sibling();

    if (node)
        return m_parentXMLDocument->getNode(node);
    else
        return nullptr;
}

/*
 * =============================================================================================================
 * XMLNode::prependNode
 * =============================================================================================================
 */
void XMLNode::prependNode(XMLNode *node)
{
    m_node->prepend_node(node->m_node);
}

/*
 * =============================================================================================================
 * XMLNode::appendNode
 * =============================================================================================================
 */
void XMLNode::appendNode(XMLNode *node)
{
    m_node->append_node(node->m_node);
}

/*
 * =============================================================================================================
 * XMLNode::insertNode
 * =============================================================================================================
 */
void XMLNode::insertNode(XMLNode *location,
                         XMLNode *node)
{
    m_node->insert_node(location->m_node, node->m_node);
}

void XMLNode::moveNodeUp(XMLNode *node)
{

}

void XMLNode::moveNodeDown(XMLNode *node)
{

}

void XMLNode::moveNodeTo(XMLNode *node,
                         unsigned int index)
{

}

/*
 * =============================================================================================================
 * XMLNode::deleteNode
 * =============================================================================================================
 */
void XMLNode::deleteNode()
{
    m_node->parent()->remove_node(m_node);
    m_parentXMLDocument->m_xmlNodes.erase(m_node);
}

void XMLNode::deleteChildNodes()
{
    m_node->remove_all_nodes();
    // TODO: What about the pointers in the XmlDocument?
}

/*
 * =============================================================================================================
 * XMLNode::firstAttribute
 * =============================================================================================================
 */
XMLAttribute *XMLNode::attribute(const std::string &name)
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

/*
 * =============================================================================================================
 * XMLNode::firstAttribute
 * =============================================================================================================
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

/*
 * =============================================================================================================
 * XMLNode::prependAttribute
 * =============================================================================================================
 */
void XMLNode::prependAttribute(XMLAttribute *attribute)
{
    m_node->prepend_attribute(attribute->m_attribute);
}

/*
 * =============================================================================================================
 * XMLNode::appendAttribute
 * =============================================================================================================
 */
void XMLNode::appendAttribute(XMLAttribute *attribute)
{
    m_node->append_attribute(attribute->m_attribute);
}
