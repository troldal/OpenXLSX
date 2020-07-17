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
Impl::XLSharedStrings::XLSharedStrings(XLXmlData* xmlData) : XLAbstractXMLFile(xmlData)
//          m_sharedStringNodes(),
//          m_emptyString("")
{
    ParseXMLData();
}

/**
 * @details Overrides the method in XLAbstractXMLFile. Clears the internal datastructure and fills it with the strings
 * in the XML file.
 * @todo Consider making the return value void.
 */
bool Impl::XLSharedStrings::ParseXMLData()
{
    return true;
}

/**
 * @details Look up a string index by the string content. If the string does not exist, the returned index is -1.
 */
int32_t Impl::XLSharedStrings::GetStringIndex(const string& str) const
{
    auto iter = std::find_if(XmlDocument().document_element().children().begin(),
                             XmlDocument().document_element().children().end(),
                             [&](const XMLNode& node) { return strcmp(node.first_child().text().get(), str.c_str()) == 0; });

    return iter == XmlDocument().document_element().children().end()
               ? -1
               : std::distance(XmlDocument().document_element().children().begin(), iter);
}

/**
 * @details
 */
bool Impl::XLSharedStrings::StringExists(const string& str) const
{
    return GetStringIndex(str) >= 0;
}

/**
 * @details
 */
bool Impl::XLSharedStrings::StringExists(uint32_t index) const
{
    return index <=
           std::distance(XmlDocument().document_element().children().begin(), XmlDocument().document_element().children().end()) - 1;
}

/**
 * @details
 */
const char* Impl::XLSharedStrings::GetString(uint32_t index) const
{
    auto iter = XmlDocument().document_element().children().begin();
    std::advance(iter, index);
    return iter->first_child().text().get();
}

/**
 * @details Append a string by creating a new node in the XML file and adding the string to it. The index to the
 * shared string is returned
 */
int32_t Impl::XLSharedStrings::AppendString(const string& str)
{
    XmlDocument().document_element().append_child("si").append_child("t").text().set(str.c_str());

    return std::distance(XmlDocument().document_element().children().begin(), XmlDocument().document_element().children().end()) - 1;
}

/**
 * @details Clear the string at the given index. This will affect the entire spreadsheet; everywhere the shared string
 * is used, it will be erased.
 */
void Impl::XLSharedStrings::ClearString(int index)
{
    auto iter = XmlDocument().document_element().children().begin();
    std::advance(iter, index);
    iter->text().set("");
}
