/*

   ____                               ____      ___ ____       ____  ____      ___
  6MMMMb                              `MM(      )M' `MM'      6MMMMb\`MM(      )M'
 8P    Y8                              `MM.     d'   MM      6M'    ` `MM.     d'
6M      Mb __ ____     ____  ___  __    `MM.   d'    MM      MM        `MM.   d'
MM      MM `M6MMMMb   6MMMMb `MM 6MMb    `MM. d'     MM      YM.        `MM. d'
MM      MM  MM'  `Mb 6M'  `Mb MMM9 `Mb    `MMd       MM       YMMMMb     `MMd
MM      MM  MM    MM MM    MM MM'   MM     dMM.      MM           `Mb     dMM.
MM      MM  MM    MM MMMMMMMM MM    MM    d'`MM.     MM            MM    d'`MM.
YM      M9  MM    MM MM       MM    MM   d'  `MM.    MM            MM   d'  `MM.
 8b    d8   MM.  ,M9 YM    d9 MM    MM  d'    `MM.   MM    / L    ,M9  d'    `MM.
  YMMMM9    MMYMMM9   YMMMM9 _MM_  _MM_M(_    _)MM_ _MMMMMMM MYMMMM9 _M(_    _)MM_
            MM
            MM
           _MM_

  Copyright (c) 2018, Kenneth Troldal Balslev

  All rights reserved.

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions are met:
  - Redistributions of source code must retain the above copyright
    notice, this list of conditions and the following disclaimer.
  - Redistributions in binary form must reproduce the above copyright
    notice, this list of conditions and the following disclaimer in the
    documentation and/or other materials provided with the distribution.
  - Neither the name of the author nor the
    names of any contributors may be used to endorse or promote products
    derived from this software without specific prior written permission.

  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
  ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
  DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
  LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
  ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
  (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
  SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

 */

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
