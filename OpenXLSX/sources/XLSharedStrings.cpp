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

#include <XLException.hpp>

using namespace OpenXLSX;

/**
 * @details Constructs a new XLSharedStrings object. Only one (common) object is allowed per XLDocument instance.
 * A filepath to the underlying XML file must be provided.
 */
XLSharedStrings::XLSharedStrings(XLXmlData* xmlData, std::deque<std::string>* stringCache)
    : XLXmlFile(xmlData),
    m_stringCache(stringCache)
{
    XMLDocument & doc = xmlDocument();
    if (doc.document_element().empty())    // handle a bad (no document element) xl/sharedStrings.xml
        doc.load_string(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
                "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"    // pull request #192 -> remove count &
                                                                                                 // uniqueCount as they are optional
                // 2024-09-03: removed empty string entry on creation, as it appears to just waste a string index that will never be used
                // "  <si>\n"
                // "    <t/>\n"
                // "  </si>\n"
                "</sst>",
                pugi_parse_settings
        );
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
    const auto iter = std::find_if(m_stringCache->begin(), m_stringCache->end(), [&](const std::string& s) { return str == s; });

    return iter == m_stringCache->end() ? -1 : static_cast<int32_t>(std::distance(m_stringCache->begin(), iter));
}

/**
 * @details
 */
bool XLSharedStrings::stringExists(const std::string& str) const { return getStringIndex(str) >= 0; }

/**
 * @details
 */
const char* XLSharedStrings::getString(int32_t index) const
{
    if (index < 0 || static_cast<size_t>(index) >= m_stringCache->size()) { // 2024-04-30: added range check
        using namespace std::literals::string_literals;
        throw XLInternalError("XLSharedStrings::"s + __func__ + ": index "s + std::to_string(index) + " is out of range"s);
    }
    return (*m_stringCache)[index].c_str();
}

/**
 * @details Append a string by creating a new node in the XML file and adding the string to it. The index to the
 * shared string is returned
 */
int32_t XLSharedStrings::appendString(const std::string& str)
{
    // size_t stringCacheSize = std::distance(m_stringCache->begin(), m_stringCache->end()); // any reason why .size() would not work?
    size_t stringCacheSize = m_stringCache->size();    // 2024-05-31: analogous with already added range check in getString
    if (stringCacheSize >= XLMaxSharedStrings)    {    // 2024-05-31: added range check
        using namespace std::literals::string_literals;
        throw XLInternalError("XLSharedStrings::"s + __func__ + ": exceeded max strings count "s + std::to_string(XLMaxSharedStrings));
    }
    auto textNode = xmlDocument().document_element().append_child("si").append_child("t");
    if ((!str.empty()) && (str.front() == ' ' || str.back() == ' '))
        textNode.append_attribute("xml:space").set_value("preserve");    // pull request #161
    textNode.text().set(str.c_str());
    m_stringCache->emplace_back(textNode.text().get());    // index of this element = previous stringCacheSize

    return static_cast<int32_t>(stringCacheSize);
}

/**
 * @details Print the underlying XML using pugixml::xml_node::print
 */
void XLSharedStrings::print(std::basic_ostream<char>& ostr) const { xmlDocument().document_element().print(ostr); }

/**
 * @details Clear the string at the given index. This will affect the entire spreadsheet; everywhere the shared string
 * is used, it will be erased.
 * @note: 2024-05-31 DONE: index now int32_t everywhere, 2 billion shared strings should be plenty
 */
void XLSharedStrings::clearString(int32_t index)    // 2024-04-30: whitespace support
{
    if (index < 0 || static_cast<size_t>(index) >= m_stringCache->size()) { // 2024-04-30: added range check
        using namespace std::literals::string_literals;
        throw XLInternalError("XLSharedStrings::"s + __func__ + ": index "s + std::to_string(index) + " is out of range"s);
    }

    (*m_stringCache)[index] = "";
    // auto iter            = xmlDocument().document_element().children().begin();
    // std::advance(iter, index);
    // iter->text().set(""); // 2024-04-30: BUGFIX: this was never going to work, <si> entries can be plenty that need to be cleared,
    // including formatting

    /* 2024-04-30 CAUTION: performance critical - with whitespace support, the function can no longer know the exact iterator position of
     *   the shared string to be cleared - TBD what to do instead?
     * Potential solution: store the XML child position with each entry in m_stringCache in a std::deque<struct entry>
     *   with struct entry { std::string s; uint64_t xmlChildIndex; };
     */
    XMLNode  sharedStringNode = xmlDocument().document_element().first_child_of_type(pugi::node_element);
    int32_t sharedStringPos  = 0;
    while (sharedStringPos < index && not sharedStringNode.empty()) {
        sharedStringNode = sharedStringNode.next_sibling_of_type(pugi::node_element);
        ++sharedStringPos;
    }
    if (not sharedStringNode.empty()) {    // index was found
        sharedStringNode.remove_children();    // clear all data and formatting
        sharedStringNode.append_child("t");    // append an empty text node
    }
}
