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
 */
void XMLDocument::LoadFile(const std::string &filePath)
{
    m_filePath = filePath;
    m_file.reset(new file<>(m_filePath.c_str()));
    m_document->parse<parse_no_data_nodes>(m_file->data());
}

/**
 * @details SaveFile saves the contents of the xml document to a file at the given path.
 * If no argument is provided, the value of the m_filePath member variable will be used. However, in case it is empty
 * and no path is provided as input, an exception will be thrown.
 */
void XMLDocument::SaveFile(const std::string &filePath) const
{
    if (m_filePath.empty() && filePath.empty()) throw invalid_argument("The file path for xml file must not be blank");

    std::string fp;
    if (filePath.empty()) fp = m_filePath;
    else fp = filePath;

    ofstream outputFile(fp);
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
    auto result = m_document->allocate_node(node_element, m_document->allocate_string(nodeName.c_str()));
    m_xmlNodes.insert_or_assign(result, unique_ptr<XMLNode>(new XMLNode(this, result)));

    if (!nodeValue.empty()) m_xmlNodes.at(result)->SetValue(nodeValue);

    return m_xmlNodes.at(result).get();
}

/**
 * @details
 */
XMLAttribute *XMLDocument::CreateAttribute(const std::string &attributeName,
                                           const std::string &attributeValue)
{
    auto result = m_document->allocate_attribute(m_document->allocate_string(attributeName.c_str()));
    m_xmlAttributes.insert_or_assign(result, unique_ptr<XMLAttribute>(new XMLAttribute(this, result)));

    if (!attributeValue.empty()) m_xmlAttributes.at(result)->SetValue(attributeValue);

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

    // Check if the XMLNode exists. If yes, return a the pointer.
    auto result = m_xmlAttributes.find(theAttribute);
    if (result != m_xmlAttributes.end()) return result->second.get();

    // Otherwise, create a new XMLAttribute, add it to the dictionary and return the pointer.
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
