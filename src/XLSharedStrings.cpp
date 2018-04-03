//
// Created by Troldal on 01/09/16.
//

#include "XLSharedStrings.h"
#include "XLDocument.h"

using namespace std;
using namespace OpenXLSX;

/**
 * @details Constructs a new XLSharedStrings object. Only one (common) object is allowed per XLDocument instance.
 * A filepath to the underlying XML file must be provided.
 */
XLSharedStrings::XLSharedStrings(XLDocument &parent,
                                 const std::string &filePath)
    : XLAbstractXMLFile(parent.RootDirectory()->string(), filePath),
      XLSpreadsheetElement(parent),
      m_sharedStrings(),
      m_emptyString("")
{
    LoadXMLData();
}

/**
 * @details Saves the XML file (if modified) and destroys the object.
 */
XLSharedStrings::~XLSharedStrings()
{
    SaveXMLData();
}

/**
 * @details Overrides the method in XLAbstractXMLFile. Clears the internal datastructure and fills it with the strings
 * in the XML file.
 * @todo Consider making the return value void.
 */
bool XLSharedStrings::ParseXMLData()
{
    // Clear the datastructure
    m_sharedStrings.clear();

    // Find the first node and iterate through the XML file, storing all string nodes in the internal datastructure
    auto node = XmlDocument().firstNode();
    while (node) {
        m_sharedStrings.push_back(node->childNode());
        node = node->nextSibling();
    }

    return true;
}

/**
 * @details Look up and return a shared string by index. If the index is larger than the number of shared strings,
 * an empty string will be returned.
 * The resulting string is returned as reference-to-const, as the client is not supposed to modify the shared strings
 * directly.
 * @todo consider creating the shared string, if it does not exist. This will require a non-const function.
 */
const string &XLSharedStrings::GetString(unsigned long index) const
{
    if (index > m_sharedStrings.size() - 1)
        return m_emptyString;
    else
        return m_sharedStrings.at(index)->value();
}

/**
 * @details Look up and return a shared string by reference. If the requested string does not exist,
 * an empty string will be returned.
 * The resulting string is returned as reference-to-const, as the client is not supposed to modify the shared strings
 * directly.
 * @todo consider creating the shared string, if it does not exist. This will require a non-const function.
 */
const std::string &XLSharedStrings::GetString(const std::string &str) const
{
    long index = GetStringIndex(str);

    if (index < 0)
        return m_emptyString;
    else
        return m_sharedStrings.at(index)->value();

}

const std::string &XLSharedStrings::GetString(const std::string &str)
{
    long index = GetStringIndex(str);

    if (index < 0) index = AppendString(str);

    return m_sharedStrings.at(index)->value();
}

/**
 * @details Look up a shared string by index and return a pointer to the corresponding node in the underlying XML file.
 * If the index is larger than the number of shared strings a nullptr will be returned.
 * The resulting string is returned as pointer-to-const, as the client is not supposed to modify the shared strings
 * directly.
 */
const XMLNode &XLSharedStrings::GetStringNode(unsigned long index) const
{
    if (index > m_sharedStrings.size() - 1)
        throw std::range_error("Node does not exist");
    else
        return *m_sharedStrings.at(index);
}

/**
 * @details Look up a shared string node by string and return a pointer to the corresponding node in the underlying XML file.
 * If the index is larger than the number of shared strings a nullptr will be returned.
 * The resulting string is returned as pointer-to-const, as the client is not supposed to modify the shared strings
 * directly.
 */
const XMLNode &XLSharedStrings::GetStringNode(const std::string &str) const
{
    XMLNode *node = nullptr;
    for (const auto &s : m_sharedStrings) {
        if (s->value() == str) node = s;
    }

    if (node == nullptr) throw std::range_error("Node does not exist");

    return *node;
}

/**
 * @details Look up a string index by the string content. If the string does not exist, the returned index is -1.
 */
long XLSharedStrings::GetStringIndex(const std::string &str) const
{

    long result = -1;
    long counter = 0;
    for (const auto &s : m_sharedStrings) {
        if (s->value() == str) {
            result = counter;
            break;
        }
        counter++;
    }

    return result;
}

/**
 * @details
 */
bool XLSharedStrings::StringExists(const std::string &str) const
{

    if (GetStringIndex(str) < 0)
        return false;
    else
        return true;
}

/**
 * @details
 */
bool XLSharedStrings::StringExists(unsigned long index) const
{
    if (index > m_sharedStrings.size() - 1)
        throw false;
    else
        return true;
}

/**
 * @details Append a string by creating a new node in the XML file and adding the string to it. The index to the
 * shared string is returned
 */
long XLSharedStrings::AppendString(const std::string &str)
{

    // Create the required nodes
    auto node = XmlDocument().createNode("si");
    auto value = XmlDocument().createNode("t");

    // Insert add the string and add the nodes to the XML tree.
    value->setValue(str);
    node->appendNode(value);
    XmlDocument().rootNode()->appendNode(node);

    // Add the node pointer to the internal datastructure.
    m_sharedStrings.push_back(value);
    SetModified();

    // Return the Index of the new string.
    return m_sharedStrings.size() - 1;
}

/**
 * @details Clear the string at the given index. This will affect the entire spreadsheet; everywhere the shared string
 * is used, it will be erased.
 */
void XLSharedStrings::ClearString(int index)
{
    m_sharedStrings.at(index)->setValue("");
    SetModified();
}
