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

// ===== External Includes ===== //
#include <pugixml.hpp>
#include <sstream>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLXmlData.hpp"

using namespace OpenXLSX;

/**
 * @details
 */
XLXmlData::XLXmlData(XLDocument* parentDoc, const std::string& xmlPath, const std::string& xmlId, XLContentType xmlType)
    : m_parentDoc(parentDoc),
      m_xmlPath(xmlPath),
      m_xmlID(xmlId),
      m_xmlType(xmlType),
      m_xmlDoc(std::make_unique<XMLDocument>())
{
    m_xmlDoc->reset();
}

/**
 * @details
 */
XLXmlData::~XLXmlData() = default;

/**
 * @details
 */
void XLXmlData::setRawData(const std::string& data) // NOLINT
{
    m_xmlDoc->load_string(data.c_str(), pugi_parse_settings);
}

/**
 * @details
 * @note Default encoding for pugixml xml_document::save is pugi::encoding_auto, becomes pugi::encoding_utf8
 */
std::string XLXmlData::getRawData(XLXmlSavingDeclaration savingDeclaration) const
{
    XMLDocument *doc = const_cast<XMLDocument *>(getXmlDocument());

    // ===== 2024-07-08: ensure that the default encoding UTF-8 is explicitly written to the XML document with a custom saving declaration
    XMLNode saveDeclaration = doc->first_child();
    if (saveDeclaration.empty() || saveDeclaration.type() != pugi::node_declaration) {    // if saving declaration node does not exist
        doc->prepend_child(pugi::node_pcdata).set_value("\n");                                // prepend a line break
        saveDeclaration = doc->prepend_child(pugi::node_declaration);                         // prepend a saving declaration
    }

    // ===== If a node_declaration could be fetched or created
    if (not saveDeclaration.empty()) {
        // ===== Fetch or create saving declaration attributes
        XMLAttribute attrVersion = saveDeclaration.attribute("version");
        if (attrVersion.empty())
            attrVersion = saveDeclaration.append_attribute("version");
        XMLAttribute attrEncoding = saveDeclaration.attribute("encoding");
        if (attrEncoding.empty())
            attrEncoding = saveDeclaration.append_attribute("encoding");
        XMLAttribute attrStandalone = saveDeclaration.attribute("standalone");
        if (attrStandalone.empty() && savingDeclaration.standalone_as_bool())    // only if standalone is set in passed savingDeclaration
            attrStandalone = saveDeclaration.append_attribute("standalone");         // then make sure it exists

        // ===== Set saving declaration attribute values (potentially overwriting existing values)
        attrVersion = savingDeclaration.version().c_str();              // version="1.0" is XML default
        attrEncoding = savingDeclaration.encoding().c_str();            // encoding="UTF-8" is XML default

        if (not attrStandalone.empty()) // only save standalone attribute if previously existing or newly set to standalone="yes"
            attrStandalone = savingDeclaration.standalone().c_str();    // standalone="no" is XML default
    }

    std::ostringstream ostr;
    doc->save(ostr, "", pugi::format_raw);
    return ostr.str();
}

/**
 * @details
 */
XLDocument* XLXmlData::getParentDoc()
{
    return m_parentDoc;
}

/**
 * @details
 */
const XLDocument* XLXmlData::getParentDoc() const
{
    return m_parentDoc;
}

/**
 * @details
 */
std::string XLXmlData::getXmlPath() const
{
    return m_xmlPath;
}

/**
 * @details
 */
std::string XLXmlData::getXmlID() const
{
    return m_xmlID;
}

/**
 * @details
 */
XLContentType XLXmlData::getXmlType() const
{
    return m_xmlType;
}

/**
 * @details
 */
XMLDocument* XLXmlData::getXmlDocument()
{
    if (!m_xmlDoc->document_element())
        m_xmlDoc->load_string(m_parentDoc->extractXmlFromArchive(m_xmlPath).c_str(), pugi_parse_settings);

    return m_xmlDoc.get();
}

/**
 * @details
 */
const XMLDocument* XLXmlData::getXmlDocument() const
{
    if (!m_xmlDoc->document_element())
        m_xmlDoc->load_string(m_parentDoc->extractXmlFromArchive(m_xmlPath).c_str(), pugi_parse_settings);

    return m_xmlDoc.get();
}
