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
bool XMLNode::IsValid() const
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
void XMLNode::SetName(const std::string &name)
{
    auto theName = m_parentXMLDocument->m_document->allocate_string(name.c_str());
    m_node->name(theName);
    m_nameCache = name;
    m_nameLoaded = true;
}

/**
 * @details
 */
const string &XMLNode::Name() const
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
void XMLNode::SetValue(const std::string &value)
{
    auto theValue = m_parentXMLDocument->m_document->allocate_string(value.c_str());
    m_node->value(theValue);
    m_valueCache = value;
    m_valueLoaded = true;
}

/**
 * @details
 */
void XMLNode::SetValue(int value)
{
    SetValue(to_string(value));
}

/**
 * @details
 */
void XMLNode::SetValue(double value)
{
    SetValue(to_string(value));
}

/**
 * @details
 */
const string &XMLNode::Value() const
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
XMLNode *XMLNode::ChildNode(const std::string &name)
{
    xml_node<> *node = nullptr;

    if (name == "") {
        node = m_node->first_node();
    }
    else {
        node = m_node->first_node(name.c_str());
    }

    if (node)
        return m_parentXMLDocument->GetNode(node);
    else
        return nullptr;
}

/**
 * @details
 */
const XMLNode *XMLNode::ChildNode(const std::string &name) const
{
    xml_node<> *node = nullptr;

    if (name == "") {
        node = m_node->first_node();
    }
    else {
        node = m_node->first_node(name.c_str());
    }

    if (node)
        return m_parentXMLDocument->GetNode(node);
    else
        return nullptr;
}

/**
 * @details
 */
XMLNode *XMLNode::Parent() const
{
    xml_node<> *node = m_node->parent();
    if (node)
        return m_parentXMLDocument->GetNode(node);
    else
        return nullptr;
}


/**
 * @details
 */
XMLNode *XMLNode::PreviousSibling() const
{
    xml_node<> *node = m_node->previous_sibling();

    if (node)
        return m_parentXMLDocument->GetNode(node);
    else
        return nullptr;
}

/**
 * @details
 */
XMLNode *XMLNode::NextSibling() const
{
    xml_node<> *node = m_node->next_sibling();

    if (node)
        return m_parentXMLDocument->GetNode(node);
    else
        return nullptr;
}

/**
 * @details
 */
void XMLNode::PrependNode(XMLNode *node)
{
    m_node->prepend_node(node->m_node);
}

/**
 * @details
 */
void XMLNode::AppendNode(XMLNode *node)
{
    m_node->append_node(node->m_node);
}

/**
 * @details
 */
void XMLNode::InsertNode(XMLNode *location, XMLNode *node)
{
    m_node->insert_node(location->m_node, node->m_node);
}

/**
 * @details
 */
void XMLNode::MoveNodeUp(XMLNode *node)
{

}

/**
 * @details
 */
void XMLNode::MoveNodeDown(XMLNode *node)
{

}

/**
 * @details
 */
void XMLNode::MoveNodeTo(XMLNode *node, unsigned int index)
{

}

/**
 * @details
 */
void XMLNode::DeleteNode()
{
    m_node->parent()->remove_node(m_node);
    m_parentXMLDocument->m_xmlNodes.erase(m_node);
}

/**
 * @details
 */
void XMLNode::DeleteChildNodes()
{
    m_node->remove_all_nodes();
    // TODO: What about the pointers in the XmlDocument?
}

/**
 * @details
 */
XMLAttribute *XMLNode::Attribute(const std::string &name)
{
    xml_attribute<> *attribute = nullptr;

    if (name.empty()) {
        attribute = m_node->first_attribute();
    }
    else {
        attribute = m_node->first_attribute(name.c_str());
    }

    if (attribute)
        return m_parentXMLDocument->GetAttribute(attribute);
    else
        return nullptr;
}

/**
 * @details
 */
const XMLAttribute *XMLNode::Attribute(const std::string &name) const
{
    xml_attribute<> *attribute = nullptr;

    if (name == "") {
        attribute = m_node->first_attribute();
    }
    else {
        attribute = m_node->first_attribute(name.c_str());
    }

    if (attribute)
        return m_parentXMLDocument->GetAttribute(attribute);
    else
        return nullptr;
}

/**
 * @details
 */
void XMLNode::PrependAttribute(XMLAttribute *attribute)
{
    m_node->prepend_attribute(attribute->m_attribute);
}

/**
 * @details
 */
void XMLNode::AppendAttribute(XMLAttribute *attribute)
{
    m_node->append_attribute(attribute->m_attribute);
}
