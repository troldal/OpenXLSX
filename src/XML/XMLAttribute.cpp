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

#include "XMLAttribute.h"

using namespace OpenXLSX;
using namespace rapidxml;
using namespace std;

/**
 * @details
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

/**
 * @details
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

/**
 * @details
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

/**
 * @details
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

/**
 * @details
 */
bool XMLAttribute::IsValid() const
{
    if (m_attribute == nullptr) {
        return false;
    }
    else {
        return true;
    }
}

/**
 * @details
 */
void XMLAttribute::SetName(const std::string &name)
{
    auto theName = m_parentXMLDocument->m_document->allocate_string(name.c_str());
    m_attribute->name(theName);
    m_nameCache = name;
    m_nameLoaded = true;
}

/**
 * @details
 */
const string &XMLAttribute::Name() const
{
    if (!m_nameLoaded) {
        m_nameCache = string(m_attribute->name());
        m_nameLoaded = true;
    }
    return m_nameCache;
}

/**
 * @details
 */
void XMLAttribute::SetValue(const std::string &value)
{
    auto theValue = m_parentXMLDocument->m_document->allocate_string(value.c_str());
    m_attribute->value(theValue);
    m_valueCache = value;
    m_valueLoaded = true;
}

/**
 * @details
 */
void XMLAttribute::SetValue(int value)
{
    SetValue(to_string(value));
}

/**
 * @details
 */
void XMLAttribute::SetValue(double value)
{
    SetValue(to_string(value));
}

/**
 * @details
 */
const string &XMLAttribute::Value() const
{
    if (!m_valueLoaded) {
        m_valueCache = string(m_attribute->value());
        m_valueLoaded = true;
    }
    return m_valueCache;
}

/**
 * @details
 */
XMLNode *XMLAttribute::Parent() const
{
    xml_node<> *node = m_attribute->parent();

    if (node)
        return m_parentXMLDocument->GetNode(node);
    else
        return nullptr;
}

/**
 * @details
 */
XMLAttribute *XMLAttribute::PreviousAttribute() const
{
    xml_attribute<> *attribute = m_attribute->previous_attribute();

    if (attribute)
        return m_parentXMLDocument->GetAttribute(attribute);
    else
        return nullptr;
}

/**
 * @details
 */
XMLAttribute *XMLAttribute::NextAttribute() const
{
    xml_attribute<> *attribute = m_attribute->next_attribute();

    if (attribute)
        return m_parentXMLDocument->GetAttribute(attribute);
    else
        return nullptr;
}

/**
 * @details
 */
void XMLAttribute::DeleteAttribute()
{
    m_attribute->parent()->remove_attribute(m_attribute);
    m_parentXMLDocument->m_xmlAttributes.erase(m_attribute);
}
