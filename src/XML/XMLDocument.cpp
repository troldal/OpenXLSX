//
// Created by Troldal on 29/08/16.
//

#include <iostream>
#include "XMLDocument.h"

using namespace rapidxml;
using namespace OpenXLSX;
using namespace std;

/*
 * =============================================================================================================
 * XMLDocument::XMLDocument
 * =============================================================================================================
 */
XMLDocument::XMLDocument()
    : m_document(new rapidxml::xml_document<>),
      m_file(nullptr),
      m_filePath(""),
      m_xmlNodes(),
      m_xmlAttributes()
{
}

/*
 * =============================================================================================================
 * XMLDocument::readData
 * =============================================================================================================
 */
void XMLDocument::readData(const std::string &xmlData)
{
    m_document->parse<parse_no_data_nodes>(m_document->allocate_string(xmlData.c_str()));
}

/*
 * =============================================================================================================
 * XMLDocument::getData
 * =============================================================================================================
 */
std::string XMLDocument::getData() const
{
    string s;
    rapidxml::print(back_inserter(s), *m_document, 0);
    // TODO: Use the print_no_indenting for production code.
    //rapidxml_ns::Print(back_inserter(s),*m_document, print_no_indenting);
    return s;
}

/*
 * =============================================================================================================
 * XMLDocument::loadFile
 * =============================================================================================================
 */
void XMLDocument::loadFile(const std::string &filePath)
{
    m_filePath = filePath;
    m_file.reset(new file<>(m_filePath.c_str()));
    m_document->parse<parse_no_data_nodes>(m_file->data());
}

/*
 * =============================================================================================================
 * XMLDocument::saveFile
 * =============================================================================================================
 */
void XMLDocument::saveFile(const std::string &filePath) const
{
    ofstream outputFile(filePath);
    outputFile << *m_document;
    outputFile.close();
}

/*
 * =============================================================================================================
 * XMLDocument::Print
 * =============================================================================================================
 */
void XMLDocument::print() const
{
    cout << *m_document;
}

/*
 * =============================================================================================================
 * XMLDocument::rootNode
 * =============================================================================================================
 */
XMLNode *XMLDocument::rootNode()
{
    xml_node<> *node = m_document->first_node();

    if (node)
        return getNode(node);
    else
        return nullptr;
}

/*
 * =============================================================================================================
 * XMLDocument::rootNode
 * =============================================================================================================
 */
const XMLNode *XMLDocument::rootNode() const
{
    xml_node<> *node = m_document->first_node();

    if (node)
        return getNode(node);
    else
        return nullptr;
}

/*
 * =============================================================================================================
 * XMLDocument::firstNode
 * =============================================================================================================
 */
XMLNode *XMLDocument::firstNode()
{
    xml_node<> *root = m_document->first_node();

    if (root) {
        xml_node<> *node = root->first_node();

        if (node)
            return getNode(node);
        else
            return nullptr;
    }
    else
        return nullptr;
}

/*
 * =============================================================================================================
 * XMLDocument::firstNode
 * =============================================================================================================
 */
const XMLNode *XMLDocument::firstNode() const
{
    xml_node<> *root = m_document->first_node();

    if (root) {
        xml_node<> *node = root->first_node();

        if (node)
            return getNode(node);
        else
            return nullptr;
    }
    else
        return nullptr;
}

/*
 * =============================================================================================================
 * XMLDocument::createNode
 * =============================================================================================================
 */
XMLNode *XMLDocument::createNode(const std::string &nodeName,
                                 const std::string &nodeValue)
{
    auto str = m_document->allocate_string(nodeName.c_str());
    auto result = m_document->allocate_node(node_element, str);
    m_xmlNodes.insert_or_assign(result, make_unique<XMLNode>(this, result));

    if (nodeValue != "") {
        m_xmlNodes.at(result)->setValue(nodeValue);
    }

    return m_xmlNodes.at(result).get();
}

/*
 * =============================================================================================================
 * XMLDocument::createAttribute
 * =============================================================================================================
 */
XMLAttribute *XMLDocument::createAttribute(const std::string &attributeName,
                                           const std::string &attributeValue)
{
    auto str = m_document->allocate_string(attributeName.c_str());
    auto result = m_document->allocate_attribute(str);
    m_xmlAttributes.insert_or_assign(result, make_unique<XMLAttribute>(this, result));

    if (attributeValue != "") {
        m_xmlAttributes.at(result)->setValue(attributeValue);
    }

    return m_xmlAttributes.at(result).get();
}

/*
 * =============================================================================================================
 * XMLDocument::getNode
 * =============================================================================================================
 */
XMLNode *XMLDocument::getNode(rapidxml::xml_node<> *theNode)
{
    if (theNode == nullptr) return nullptr;

    // Check if the XMLNode exists. If yes, return a the pointer.
    auto result = m_xmlNodes.find(theNode);
    if (result != m_xmlNodes.end()) return result->second.get();

    // Otherwise, create a new XMLNode, add it to the dictionary and return the pointer.
    m_xmlNodes.insert_or_assign(theNode, make_unique<XMLNode>(this, theNode));
    return m_xmlNodes.at(theNode).get();
}

/*
 * =============================================================================================================
 * XMLDocument::getNode
 * =============================================================================================================
 */
const XMLNode *XMLDocument::getNode(const rapidxml::xml_node<> *theNode) const
{
    if (theNode == nullptr) return nullptr;

    // For some reason the std::map functions can't take a const pointer as key, although they are specified as const. Hence the use of const_cast.

    // Check if the XMLNode exists. If yes, return a the pointer.
    auto result = m_xmlNodes.find(const_cast<xml_node<> *>(theNode));
    if (result != m_xmlNodes.end()) return result->second.get();

    // Otherwise, return a nullptr.
    return nullptr;
}

/*
 * =============================================================================================================
 * XMLDocument::getAttribute
 * =============================================================================================================
 */
XMLAttribute *XMLDocument::getAttribute(rapidxml::xml_attribute<> *theAttribute)
{
    if (theAttribute == nullptr) return nullptr;

    if (!theAttribute) return nullptr;
    // Check if the XMLNode exists. If yes, return a the pointer.
    auto result = m_xmlAttributes.find(theAttribute);
    if (result != m_xmlAttributes.end()) return result->second.get();

    // Otherwise, create a new XMLNode, add it to the dictionary and return the pointer.
    m_xmlAttributes.insert_or_assign(theAttribute, make_unique<XMLAttribute>(this, theAttribute));
    return m_xmlAttributes.at(theAttribute).get();
}

/*
 * =============================================================================================================
 * XMLDocument::getAttribute
 * =============================================================================================================
 */
const XMLAttribute *XMLDocument::getAttribute(const rapidxml::xml_attribute<> *theAttribute) const
{
    if (theAttribute == nullptr) return nullptr;

    // For some reason the std::map functions can't take a const pointer as key, although they are specified as const. Hence the use of const_cast.

    // Check if the XMLNode exists. If yes, return a the pointer.
    auto result = m_xmlAttributes.find(const_cast<xml_attribute<> *>(theAttribute));
    if (result != m_xmlAttributes.end()) return result->second.get();

    // Otherwise, return a nullptr.
    return nullptr;
}
