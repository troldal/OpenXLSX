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

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLXmlFile.hpp"

using namespace OpenXLSX;

/**
 * @details The constructor creates a new object with the parent XLDocument and the file path as input, with
 * an optional input being a std::string with the XML data. If the XML data is provided by a string, any file with
 * the same path in the .zip file will be overwritten upon saving of the document. If no xmlData is provided,
 * the data will be read from the .zip file, using the given path.
 */
XLXmlFile::XLXmlFile(XLXmlData* xmlData) : m_xmlData(xmlData) {}

XLXmlFile::~XLXmlFile() = default;

/**
 * @details This method sets the XML data with a std::string as input. The underlying XMLDocument reads the data.
 * When envoking the load_string method in PugiXML, the flag 'parse_ws_pcdata' is passed along with the default flags.
 * This will enable parsing of whitespace characters. If not set, Excel cells with only spaces will be returned as
 * empty strings, which is not what we want. The downside is that whitespace characters such as \\n and \\t in the
 * input xml file may mess up the parsing.
 */
void XLXmlFile::setXmlData(const std::string& xmlData)
{
    m_xmlData->setRawData(xmlData);
}

/**
 * @details This method retrieves the underlying XML data as a std::string.
 */
std::string XLXmlFile::xmlData() const
{
    return m_xmlData->getRawData();
}

/**
 * @details
 */
const XLDocument& XLXmlFile::parentDoc() const
{
    return *m_xmlData->getParentDoc();
}

/**
 * @details
 */
XLDocument& XLXmlFile::parentDoc()
{
    return *m_xmlData->getParentDoc();
}

/**
 * @details
 */
std::string XLXmlFile::relationshipID() const
{
    return m_xmlData->getXmlID();
}

/**
 * @details This method returns a pointer to the underlying XMLDocument resource.
 */
XMLDocument& XLXmlFile::xmlDocument()
{
    return const_cast<XMLDocument&>(static_cast<const XLXmlFile*>(this)->xmlDocument());    // NOLINT
}

/**
 * @details This method returns a pointer to the underlying XMLDocument resource as const.
 */
const XMLDocument& XLXmlFile::xmlDocument() const
{
    return *m_xmlData->getXmlDocument();
}
