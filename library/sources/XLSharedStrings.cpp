//
// Created by Troldal on 01/09/16.
//

#include "XLSharedStrings.hpp"

#include "XLDocument.hpp"

//#include <pugixml.hpp>
#include <algorithm>

using namespace std;
using namespace OpenXLSX;

/**
 * @details Constructs a new XLSharedStrings object. Only one (common) object is allowed per XLDocument instance.
 * A filepath to the underlying XML file must be provided.
 */
XLSharedStrings::XLSharedStrings(XLXmlData* xmlData) : XLXmlFile(xmlData) {}

/**
 * @details Look up a string index by the string content. If the string does not exist, the returned index is -1.
 */
int32_t XLSharedStrings::getStringIndex(const string& str) const
{
    auto iter = std::find_if(xmlDocument().document_element().children().begin(),
                             xmlDocument().document_element().children().end(),
                             [&](const XMLNode& node) { return strcmp(node.first_child().text().get(), str.c_str()) == 0; });

    return iter == xmlDocument().document_element().children().end()
               ? -1
               : std::distance(xmlDocument().document_element().children().begin(), iter);
}

/**
 * @details
 */
bool XLSharedStrings::stringExists(const string& str) const
{
    return getStringIndex(str) >= 0;
}

/**
 * @details
 */
bool XLSharedStrings::stringExists(uint32_t index) const
{
    return index <=
           std::distance(xmlDocument().document_element().children().begin(), xmlDocument().document_element().children().end()) - 1;
}

/**
 * @details
 */
const char* XLSharedStrings::getString(uint32_t index) const
{
    auto iter = xmlDocument().document_element().children().begin();
    std::advance(iter, index);
    return iter->first_child().text().get();
}

/**
 * @details Append a string by creating a new node in the XML file and adding the string to it. The index to the
 * shared string is returned
 */
int32_t XLSharedStrings::appendString(const string& str)
{
    xmlDocument().document_element().append_child("si").append_child("t").text().set(str.c_str());

    return std::distance(xmlDocument().document_element().children().begin(), xmlDocument().document_element().children().end()) - 1;
}

/**
 * @details Clear the string at the given index. This will affect the entire spreadsheet; everywhere the shared string
 * is used, it will be erased.
 */
void XLSharedStrings::clearString(int index)
{
    auto iter = xmlDocument().document_element().children().begin();
    std::advance(iter, index);
    iter->text().set("");
}
