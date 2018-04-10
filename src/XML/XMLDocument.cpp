//
// Created by Troldal on 29/08/16.
//

#include <iostream>
#include <memory>
#include <stdexcept>
#include "XMLDocument.h"

using namespace rapidxml;
using namespace OpenXLSX;
using namespace std;

/**
 * @details The default constructor creates a new object, with the member variables set to sensible initial values.
 */
XMLDocument::XMLDocument()
    : m_document(new rapidxml::xml_document<>),
      m_file(nullptr),
      m_filePath(""),
      m_xmlNodes(),
      m_xmlAttributes()
{
}

/**
 * @details The ReadData function parses the content of the input .xml string, using the RapidXML parse function.
 * The input string is copied to the string pool of the xml document object. If the m_file and m_filePath
 * member variables are set, those will be erased or set to null.
 */
void XMLDocument::ReadData(const std::string &xmlData)
{
    m_filePath.erase();
    m_file.reset(nullptr);
    m_document->parse<parse_no_data_nodes>(m_document->allocate_string(xmlData.c_str()));
}

/**
 * @details The GetData function creates a new std::string and populates it with the contents of the xml document
 * object. The resulting string is returned to the caller.
 * @todo Use the print_no_indenting for production code.
 */
std::string XMLDocument::GetData() const
{
    string s;
    rapidxml::print(back_inserter(s), *m_document, 0);
    return s;
}

/**
 * @details LoadFile sets the m_filePath to the input string and loads the file with the given path by creating a new
 * RapidXML file object. The contents of the file is parsed by the document object.
 * If no argument is provided, the value of the m_filePath member variable will be used. However, in case it is empty
 * and no path is provided as input, an exception will be thrown.
 */
void XMLDocument::LoadFile(const std::string &filePath)
{
    if (!filePath.empty()) m_filePath = filePath;
    if (m_filePath.empty()) throw invalid_argument("The file path for xml file must not be blank");
    m_file.reset(new file<>(m_filePath.c_str()));
    m_document->parse<parse_no_data_nodes>(m_file->data());
}

/**
 * @details
 */
void XMLDocument::SaveFile(const std::string &filePath) const
{
    ofstream outputFile(filePath);
    outputFile << *m_document;
    outputFile.close();
}

/**
 * @details
 */
void XMLDocument::Print() const
{
    cout << *m_document;
}

/**
 * @details
 */
XMLNode *XMLDocument::RootNode()
{
    xml_node<> *node = m_document->first_node();

    if (node) return GetNode(node);
    else return nullptr;
}

/**
 * @details
 */
const XMLNode *XMLDocument::RootNode() const
{
    xml_node<> *node = m_document->first_node();

    if (node) return GetNode(node);
    else return nullptr;
}

/**
 * @details
 */
XMLNode *XMLDocument::FirstNode()
{
    xml_node<> *root = m_document->first_node();

    if (root) {
        xml_node<> *node = root->first_node();

        if (node) return GetNode(node);
        else return nullptr;
    }
    else return nullptr;
}

/**
 * @details
 */
const XMLNode *XMLDocument::FirstNode() const
{
    xml_node<> *root = m_document->first_node();

    if (root) {
        xml_node<> *node = root->first_node();

        if (node) return GetNode(node);
        else return nullptr;
    }
    else return nullptr;
}

/**
 * @details
 */
XMLNode *XMLDocument::CreateNode(const std::string &nodeName,
                                 const std::string &nodeValue)
{
    auto str = m_document->allocate_string(nodeName.c_str());
    auto result = m_document->allocate_node(node_element, str);
    m_xmlNodes.insert_or_assign(result, unique_ptr<XMLNode>(new XMLNode(this, result)));

    if (nodeValue != "") {
        m_xmlNodes.at(result)->SetValue(nodeValue);
    }

    return m_xmlNodes.at(result).get();
}

/**
 * @details
 */
XMLAttribute *XMLDocument::CreateAttribute(const std::string &attributeName,
                                           const std::string &attributeValue)
{
    auto str = m_document->allocate_string(attributeName.c_str());
    auto result = m_document->allocate_attribute(str);
    m_xmlAttributes.insert_or_assign(result, unique_ptr<XMLAttribute>(new XMLAttribute(this, result)));

    if (attributeValue != "") {
        m_xmlAttributes.at(result)->SetValue(attributeValue);
    }

    return m_xmlAttributes.at(result).get();
}

/**
 * @details
 */
XMLNode *XMLDocument::GetNode(rapidxml::xml_node<> *theNode)
{
    if (theNode == nullptr) return nullptr;

    // Check if the XMLNode exists. If yes, return a the pointer.
    auto result = m_xmlNodes.find(theNode);
    if (result != m_xmlNodes.end()) return result->second.get();

    // Otherwise, create a new XMLNode, add it to the dictionary and return the pointer.
    m_xmlNodes.insert_or_assign(theNode, unique_ptr<XMLNode>(new XMLNode(this, theNode)));
    return m_xmlNodes.at(theNode).get();
}

/**
 * @details
 */
const XMLNode *XMLDocument::GetNode(const rapidxml::xml_node<> *theNode) const
{
    if (theNode == nullptr) return nullptr;

    // For some reason the std::map functions can't take a const pointer as key, although they are specified as const. Hence the use of const_cast.

    // Check if the XMLNode exists. If yes, return a the pointer.
    auto result = m_xmlNodes.find(const_cast<xml_node<> *>(theNode));
    if (result != m_xmlNodes.end()) return result->second.get();

    // Otherwise, return a nullptr.
    return nullptr;
}

/**
 * @details
 */
XMLAttribute *XMLDocument::GetAttribute(rapidxml::xml_attribute<> *theAttribute)
{
    if (theAttribute == nullptr) return nullptr;

    if (!theAttribute) return nullptr;
    // Check if the XMLNode exists. If yes, return a the pointer.
    auto result = m_xmlAttributes.find(theAttribute);
    if (result != m_xmlAttributes.end()) return result->second.get();

    // Otherwise, create a new XMLNode, add it to the dictionary and return the pointer.
    m_xmlAttributes.insert_or_assign(theAttribute, unique_ptr<XMLAttribute>(new XMLAttribute(this, theAttribute)));
    return m_xmlAttributes.at(theAttribute).get();
}

/**
 * @details
 */
const XMLAttribute *XMLDocument::GetAttribute(const rapidxml::xml_attribute<> *theAttribute) const
{
    if (theAttribute == nullptr) return nullptr;

    // For some reason the std::map functions can't take a const pointer as key, although they are specified as const. Hence the use of const_cast.

    // Check if the XMLNode exists. If yes, return a the pointer.
    auto result = m_xmlAttributes.find(const_cast<xml_attribute<> *>(theAttribute));
    if (result != m_xmlAttributes.end()) return result->second.get();

    // Otherwise, return a nullptr.
    return nullptr;
}
