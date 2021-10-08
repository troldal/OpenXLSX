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
#include <algorithm>
#include <pugixml.hpp>

// ===== OpenXLSX Includes ===== //
#include "XLDocument.hpp"
#include "XLSharedStrings.hpp"

using namespace OpenXLSX;

/**
 * @details Constructs a new XLSharedStrings object. Only one (common) object is allowed per XLDocument instance.
 * A filepath to the underlying XML file must be provided.
 */
XLSharedStrings::XLSharedStrings(XLXmlData* xmlData, std::deque<std::string> *stringCache) : XLXmlFile(xmlData), m_stringCache(stringCache)
{
}

/**
 * @details
 */
XLSharedStrings::~XLSharedStrings() = default;

/**
 * @details Look up a string index by the string content. If the string does not exist, the returned index is -1.
 */
int32_t XLSharedStrings::getStringIndex(const std::string& str) const
{
    auto iter = std::find_if(m_stringCache->begin(), m_stringCache->end(), [&](const std::string& s) { return str == s; });

    return iter == m_stringCache->end() ? -1 : static_cast<int32_t>(std::distance(m_stringCache->begin(), iter));
}

/**
 * @details
 */
bool XLSharedStrings::stringExists(const std::string& str) const
{
    return getStringIndex(str) >= 0;
}

/**
 * @details
 */
const char* XLSharedStrings::getString(uint32_t index) const
{
    return (*m_stringCache)[index].c_str();
}

/**
 * @details Append a string by creating a new node in the XML file and adding the string to it. The index to the
 * shared string is returned
 */
int32_t XLSharedStrings::appendString(const std::string& str)
{
    auto textNode = xmlDocument().document_element().append_child("si").append_child("t");
    if (str.front() == ' ' || str.back() == ' ') textNode.append_attribute("xml:space").set_value("preserve");

    textNode.text().set(str.c_str());
    m_stringCache->emplace_back(textNode.text().get());

    return static_cast<int32_t>(std::distance(m_stringCache->begin(), m_stringCache->end()) - 1);
}

/**
 * @details Clear the string at the given index. This will affect the entire spreadsheet; everywhere the shared string
 * is used, it will be erased.
 */
void XLSharedStrings::clearString(uint64_t index)
{
    (*m_stringCache)[index] = "";
    auto iter            = xmlDocument().document_element().children().begin();
    std::advance(iter, index);
    iter->text().set("");
}
