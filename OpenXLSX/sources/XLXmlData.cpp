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
#include <cstring>
#include <stack>
#include <algorithm>

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

int XLXmlData::verifyData(std::string* dbgMsg) const
{
    int errorCount = 0;
    
    // Call verifyUniqueXMLRecords to check for duplicates
    errorCount += verifyUniqueXMLRecords(dbgMsg);
    
    return errorCount;
}

int XLXmlData::verifyUniqueXMLRecords(std::string* dbgMsg) const
{
    int errorCount = 0;
    
    // Get the XML document and call recursive verification on root node
    const XMLDocument* xmlDoc = getXmlDocument();
    if (xmlDoc != nullptr) {
        errorCount += verifyUniqueXMLRecordsRecursive(xmlDoc->document_element(), dbgMsg);
    }
    
    return errorCount;
}

/*
pugi::xml_node::operator<() only compares pointers.  This function compares data.
*/
int XLXmlData::compareXMLNode(const pugi::xml_node& x, const pugi::xml_node& y)
{
    // Handle empty nodes
    if (x.empty() && !y.empty()) return -1;
    if (y.empty() && !x.empty()) return 1;
    if (x.empty() && y.empty()) return 0;
    
    // Compare node types
    if (x.type() != y.type()) {
        return (x.type() < y.type()) ? -1 : 1;
    }
    
    // Compare names (handle null pointers)
    const char* xName = x.name();
    const char* yName = y.name();
    
    if (xName == nullptr && yName != nullptr) return -1;
    if (yName == nullptr && xName != nullptr) return 1;
    if (xName == nullptr && yName == nullptr) {
        // Both null, continue to value comparison
    } else {
        int nameCompare = std::strcmp(xName, yName);
        if (nameCompare != 0) return nameCompare;
    }
    
    // Compare values
    const char* xValue = x.value();
    const char* yValue = y.value();
    
    if (xValue == nullptr && yValue != nullptr) return -1;
    if (yValue == nullptr && xValue != nullptr) return 1;
    if (xValue == nullptr && yValue == nullptr) {
        // Both null, continue to attribute comparison
    } else {
        int valueCompare = std::strcmp(xValue, yValue);
        if (valueCompare != 0) return valueCompare;
    }
    
    // Compare attributes (this is crucial for distinguishing nodes)
    auto xAttr = x.attributes_begin();
    auto yAttr = y.attributes_begin();
    auto xAttrEnd = x.attributes_end();
    auto yAttrEnd = y.attributes_end();
    
    while (xAttr != xAttrEnd && yAttr != yAttrEnd) {
        // Compare attribute names
        int attrNameCompare = std::strcmp(xAttr->name(), yAttr->name());
        if (attrNameCompare != 0) return attrNameCompare;
        
        // Compare attribute values
        int attrValueCompare = std::strcmp(xAttr->value(), yAttr->value());
        if (attrValueCompare != 0) return attrValueCompare;
        
        ++xAttr;
        ++yAttr;
    }
    
    // If one has more attributes than the other
    if (xAttr != xAttrEnd) return 1;  // x has more attributes
    if (yAttr != yAttrEnd) return -1; // y has more attributes
    
    return 0; // All comparisons equal
}

bool XLXmlData::lessXMLNode(const pugi::xml_node& x, const pugi::xml_node& y)
{
    const int compareResult = compareXMLNode(x, y);
    return (compareResult < 0);
}


int XLXmlData::verifyUniqueXMLRecordsRecursive(const pugi::xml_node& rootNode, std::string* dbgMsg)
{
    int errorCount = 0;
    
    std::stack<pugi::xml_node> nodeStack;
    if( !rootNode.empty() ){
        nodeStack.push(rootNode);
    }
    std::vector<pugi::xml_node> children; // Declare outside the loop

    while (!nodeStack.empty()) {
        const pugi::xml_node currentNode = nodeStack.top();
        nodeStack.pop();
        
        // Clear and collect children using first_child() and next_sibling()
        // Only check element nodes for duplicate detection - ignore formatting nodes
        children.resize(0); 
        for (pugi::xml_node child = currentNode.first_child(); !child.empty(); child = child.next_sibling()) {
            // Only process element nodes (node_element) - these contain actual data
            // Only process leaf nodes
            if ((child.type() == pugi::node_element) && (child.first_child().empty()) ) {
                children.push_back(child);
            }
        }

        // Sort children using lessXMLNode comparison function
        std::sort(children.begin(), children.end(), &XLXmlData::lessXMLNode);

        // Iterate sorted children and detect duplicates
        for (size_t i = 1; i < children.size(); ++i) {
            if (compareXMLNode(children[i-1], children[i]) == 0) {
                // Found duplicate nodes
                errorCount++;
                
                // Count how many adjacent children are the same
                size_t duplicateCount = 1;
                while (i + duplicateCount < children.size() && 
                       compareXMLNode(children[i-1], children[i + duplicateCount]) == 0) {
                    ++duplicateCount;
                }
                
                if (dbgMsg != nullptr) {
                    // Create deluxe debug message with comprehensive node information
                    *dbgMsg += "=== DUPLICATE XML NODES DETECTED ===\n";
                    
                    // Show root node information
                    *dbgMsg += "Root Node: ";
                    if (rootNode.empty()) {
                        *dbgMsg += "[EMPTY ROOT]\n";
                    } else {
                        *dbgMsg += "Type=" + std::to_string(static_cast<int>(rootNode.type())) + 
                                    ", Name='" + (rootNode.name() ? rootNode.name() : "[NULL]") + "'" +
                                    ", Value='" + (rootNode.value() ? rootNode.value() : "[NULL]") + "'";
                            
                        // Show root node attributes if any
                        if (!rootNode.attributes().empty()) {
                            *dbgMsg += ", Attributes: ";
                            bool firstAttr = true;
                            for (const pugi::xml_attribute& attr : rootNode.attributes()) {
                                std::string attrName = (attr.name()) ? attr.name() : "";
                                std::string attrValue = (attr.value()) ? attr.value() : "";
                                if (!firstAttr) *dbgMsg += ", ";
                                *dbgMsg += attrName + "='" + attrValue + "'";
                                firstAttr = false;
                            }
                        }
                        *dbgMsg += "\n";
                    }
                    
                    *dbgMsg += "Parent Node: ";
                    if (currentNode.empty()) {
                        *dbgMsg += "[ROOT/DOCUMENT]\n";
                    } else {
                        *dbgMsg += "Type=" + std::to_string(static_cast<int>(currentNode.type())) + 
                                  ", Name='" + (currentNode.name() ? currentNode.name() : "[NULL]") + "'" +
                                  ", Value='" + (currentNode.value() ? currentNode.value() : "[NULL]") + "'";
                        
                        // Show parent node attributes if any
                        if (!currentNode.attributes().empty()) {
                            *dbgMsg += ", Attributes: ";
                            bool firstAttr = true;
                            for (const pugi::xml_attribute& attr : currentNode.attributes()) {
                                std::string attrName = (attr.name()) ? attr.name() : "";
                                std::string attrValue = (attr.value()) ? attr.value() : "";
                                if (!firstAttr) *dbgMsg += ", ";
                                *dbgMsg += attrName + "='" + attrValue + "'";
                                firstAttr = false;
                            }
                        }
                        *dbgMsg += "\n";
                    }
                    
                    *dbgMsg += "Duplicate Count: " + std::to_string(duplicateCount + 1) + " identical nodes\n";
                    *dbgMsg += "Duplicate Node Details:\n";
                    
                    // Show details of the first duplicate node
                    const pugi::xml_node& duplicateNode = children[i-1];
                    *dbgMsg += "  - Type=" + std::to_string(static_cast<int>(duplicateNode.type())) + 
                              ", Name='" + (duplicateNode.name() ? duplicateNode.name() : "[NULL]") + "'" +
                              ", Value='" + (duplicateNode.value() ? duplicateNode.value() : "[NULL]") + "'\n";
                    
                    // Show attributes if any
                    if (!duplicateNode.attributes().empty()) {
                        *dbgMsg += "  - Attributes: ";
                        bool firstAttr = true;
                        for (const pugi::xml_attribute& attr : duplicateNode.attributes()) {
                            std::string attrName = (attr.name()) ? attr.name() : "";
                            std::string attrValue = (attr.value()) ? attr.value() : "";
                            if (!firstAttr) *dbgMsg += ", ";
                            *dbgMsg += attrName + "='" + attrValue + "'";
                            firstAttr = false;
                        }
                        *dbgMsg += "\n";
                    }
                    
                    *dbgMsg += "=====================================\n";
                }
                
                // Skip past all the duplicates
                i += duplicateCount - 1;
            }
        }

        // Push children onto the stack for further processing
        for (const pugi::xml_node& child : children) {
            nodeStack.push(child);
        }
    }
    
    return errorCount;
}
