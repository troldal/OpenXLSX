//
// Created by Troldal on 28/08/16.
//

#include "XMLAttribute.h"

using namespace RapidXLSX;
using namespace rapidxml;
using namespace std;

/*
 * =============================================================================================================
 * XMLAttribute::XMLAttribute
 * =============================================================================================================
 */
XMLAttribute::XMLAttribute(XMLDocument *parentDocument,
                           rapidxml::xml_attribute<> *attribute)
    : m_attribute(attribute),
      m_parentXMLDocument(
          parentDocument),
      m_valueCache(""),
      m_valueLoaded(false),
      m_nameCache(""),
      m_nameLoaded(false)
{
}

/*
 * =============================================================================================================
 * XMLAttribute::XMLAttribute
 * =============================================================================================================
 */
XMLAttribute::XMLAttribute(const XMLAttribute &other)
    : m_attribute(other.m_attribute),
      m_parentXMLDocument(other.m_parentXMLDocument),
      m_valueCache(other.m_valueCache),
      m_valueLoaded(other.m_valueLoaded),
      m_nameCache(other.m_nameCache),
      m_nameLoaded(other.m_nameLoaded)
{

}

/*
 * =============================================================================================================
 * XMLAttribute::~XMLAttribute
 * =============================================================================================================
 */
XMLAttribute::~XMLAttribute()
{

}

/*
 * =============================================================================================================
 * XMLAttribute::operator=
 * =============================================================================================================
 */
XMLAttribute &XMLAttribute::operator=(const XMLAttribute &other)
{
    m_attribute = other.m_attribute;
    m_parentXMLDocument = other.m_parentXMLDocument;
    m_valueCache = other.m_valueCache;
    m_valueLoaded = other.m_valueLoaded;
    m_nameCache = other.m_nameCache;
    m_nameLoaded = other.m_nameLoaded;
    return *this;
}

/*
 * =============================================================================================================
 * XMLAttribute::operator bool
 * =============================================================================================================
 */
XMLAttribute::operator bool() const
{
    if (m_attribute == nullptr) {
        return false;
    }
    else {
        return true;
    }
}

/*
 * =============================================================================================================
 * XMLAttribute::isValid
 * =============================================================================================================
 */
bool XMLAttribute::isValid() const
{
    if (m_attribute == nullptr) {
        return false;
    }
    else {
        return true;
    }
}

/*
 * =============================================================================================================
 * XMLAttribute::SetName
 * =============================================================================================================
 */
void XMLAttribute::setName(const std::string &name)
{
    auto theName = m_parentXMLDocument->m_document->allocate_string(name.c_str());
    m_attribute->name(theName);
    m_nameCache = name;
    m_nameLoaded = true;
}

/*
 * =============================================================================================================
 * XMLAttribute::Name
 * =============================================================================================================
 */
const string &XMLAttribute::name() const
{
    if (!m_nameLoaded) {
        m_nameCache = string(m_attribute->name());
        m_nameLoaded = true;
    }
    return m_nameCache;
}

/*
 * =============================================================================================================
 * XMLAttribute::SetValue
 * =============================================================================================================
 */
void XMLAttribute::setValue(const std::string &value)
{
    auto theValue = m_parentXMLDocument->m_document->allocate_string(value.c_str());
    m_attribute->value(theValue);
    m_valueCache = value;
    m_valueLoaded = true;
}

/*
 * =============================================================================================================
 * XMLAttribute::SetValue
 * =============================================================================================================
 */
void XMLAttribute::setValue(int value)
{
    setValue(to_string(value));
}

/*
 * =============================================================================================================
 * XMLAttribute::SetValue
 * =============================================================================================================
 */
void XMLAttribute::setValue(double value)
{
    setValue(to_string(value));
}

/*
 * =============================================================================================================
 * XMLAttribute::ValueAsString
 * =============================================================================================================
 */
const string &XMLAttribute::value() const
{
    if (!m_valueLoaded) {
        m_valueCache = string(m_attribute->value());
        m_valueLoaded = true;
    }
    return m_valueCache;
}

/*
 * =============================================================================================================
 * XMLAttribute::parent
 * =============================================================================================================
 */
XMLNode *XMLAttribute::parent() const
{
    xml_node<> *node = m_attribute->parent();

    if (node)
        return m_parentXMLDocument->getNode(node);
    else
        return nullptr;
}

/*
 * =============================================================================================================
 * XMLAttribute::previousAttribute
 * =============================================================================================================
 */
XMLAttribute *XMLAttribute::previousAttribute() const
{
    xml_attribute<> *attribute = m_attribute->previous_attribute();

    if (attribute)
        return m_parentXMLDocument->getAttribute(attribute);
    else
        return nullptr;
}

/*
 * =============================================================================================================
 * XMLAttribute::nextAttribute
 * =============================================================================================================
 */
XMLAttribute *XMLAttribute::nextAttribute() const
{
    xml_attribute<> *attribute = m_attribute->next_attribute();

    if (attribute)
        return m_parentXMLDocument->getAttribute(attribute);
    else
        return nullptr;
}

/*
 * =============================================================================================================
 * XMLAttribute::deleteAttribute
 * =============================================================================================================
 */
void XMLAttribute::deleteAttribute()
{
    m_attribute->parent()->remove_attribute(m_attribute);
    m_parentXMLDocument->m_xmlAttributes.erase(m_attribute);
}
