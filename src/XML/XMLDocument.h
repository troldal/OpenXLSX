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

#ifndef OPENXL_XMLDOCUMENT_H
#define OPENXL_XMLDOCUMENT_H

#include <string>
#include <memory>
#include <map>
#include "rapidxml.hpp"
#include "rapidxml_utils.hpp"
#include "rapidxml_print.hpp"
#include "XMLAttribute.h"
#include "XMLNode.h"

namespace OpenXLSX
{
    class XMLNode;

    class XMLAttribute;

/**
 * @brief Encapsulates the concept of an XML document.
 */
    class XMLDocument
    {
        friend class XMLNode;

        friend class XMLAttribute;

    public:
        /*
         * =============================================================================================================
         * XMLDocument::XMLDocument
         * =============================================================================================================
         */

        /**
         * @brief Default constructor.
         * @todo Consider constructors for filePath and xmlData. This will require defining classes for XMLData and/or XMLFile.
         */
        XMLDocument();//const std::string& FilePath = "");

        /*
         * =============================================================================================================
         * XMLDocument::~XMLDocument
         * =============================================================================================================
         */

        /**
         * @brief Destructor
         */
        virtual ~XMLDocument() = default;

        /*
         * =============================================================================================================
         * XMLDocument::readData
         * =============================================================================================================
         */

        /**
         * @brief Loads the XML data as a std::string, and parses the contents.
         * @param xmlData A std::string with the XML data.
         */
        virtual void readData(const std::string &xmlData);

        /*
         * =============================================================================================================
         * XMLDocument::getData
         * =============================================================================================================
         */

        /**
         * @brief Get the XML data as a std::string.
         * @return A std::string with the XML data.
         */
        std::string getData() const;

        /*
         * =============================================================================================================
         * XMLDocument::loadFile
         * =============================================================================================================
         */

        /**
         * @brief Opens an XML file and parses the document. If a document is already open, it will be closed without saving.
         * @param filePath The path of the XML file with the document.
         */
        virtual void loadFile(const std::string &filePath);

        /*
         * =============================================================================================================
         * XMLDocument::saveFile
         * =============================================================================================================
         */

        /**
         * @brief Saves the current XML file. If the file exists, it will be overwritten.
         * @param filePath The path of the XML file to be saved
         */
        virtual void saveFile(const std::string &filePath) const;

        /*
         * =============================================================================================================
         * XMLDocument::Print
         * =============================================================================================================
         */

        /**
         * @brief Prints the entire document to the standard output.
         */
        virtual void print() const;

        /*
         * =============================================================================================================
         * XMLDocument::rootNode
         * =============================================================================================================
         */

        /**
         * @brief Get a pointer to the root of the XML document.
         * @return A pointer to the XMLNode object. The object must not be explicitly deleted by the caller.
         */
        virtual XMLNode *rootNode();

        /*
         * =============================================================================================================
         * XMLDocument::rootNode
         * =============================================================================================================
         */

        /**
         * @brief Get a pointer to the root of the XML document not including the declaration node).
         * @return A const pointer to the XMLNode object. The object must not be explicitly deleted by the caller.
         */
        virtual const XMLNode *rootNode() const;

        /*
         * =============================================================================================================
         * XMLDocument::firstNode
         * =============================================================================================================
         */

        /**
         * @brief Get a pointer to the first node of the XML document (first child node of the root).
         * @return A pointer to the XMLNode object. he object must not be explicitly deleted by the caller.
         */
        virtual XMLNode *firstNode();

        /*
         * =============================================================================================================
         * XMLDocument::firstNode
         * =============================================================================================================
         */

        /**
         * @brief Get a pointer to the first node of the XML document (first child node of the root).
         * @return A pointer to the XMLNode object. he object must not be explicitly deleted by the caller.
         */
        virtual const XMLNode *firstNode() const;

        /*
         * =============================================================================================================
         * XMLDocument::createNode
         * =============================================================================================================
         */

        /**
         * @brief Create a new node to the document. The node has to be explicitly added as child to an existing node.
         * @param nodeName The name of the node to create.
         * @param nodeValue The value of the node to create.
         * @return A pointer to the newly created node.
         */
        virtual XMLNode *createNode(const std::string &nodeName,
                                    const std::string &nodeValue = "");

        /*
         * =============================================================================================================
         * XMLDocument::createAttribute
         * =============================================================================================================
         */

        /**
         * @brief Create a new node attribute to the document. The node has to be explicitly added to an existing node.
         * @param attributeName The name of the attribute to create.
         * @param attributeValue the value of the attribute to create.
         * @return A pointer to the newly created attribute.
         */
        virtual XMLAttribute *createAttribute(const std::string &attributeName,
                                              const std::string &attributeValue);

    protected:
        /*
         * =============================================================================================================
         * XMLDocument::getNode
         * =============================================================================================================
         */

        /**
         * @brief Get the XMLNode associated with the underlying xml_node resource.
         * @param theNode A const pointer to the xml_node to look up.
         * @return The XMLNode associated with the provided xml_node resource.
         */
        virtual XMLNode *getNode(rapidxml::xml_node<> *theNode);

        /*
         * =============================================================================================================
         * XMLDocument::getNode
         * =============================================================================================================
         */

        /**
         * @brief Get the XMLNode associated with the underlying xml_node resource.
         * @param theNode A const pointer to the xml_node to look up.
         * @return The const XMLNode associated with the provided xml_node resource.
         */
        virtual const XMLNode *getNode(const rapidxml::xml_node<> *theNode) const;

        /*
         * =============================================================================================================
         * XMLDocument::getAttribute
         * =============================================================================================================
         */

        /**
         * @brief Get the XMLAttribute associated with the underlying xml_attribute resource.
         * @param theAttribute A const pointer to the xml_node to look up.
         * @return The XMLNode associated with the provided xml_node resource.
         */
        virtual XMLAttribute *getAttribute(rapidxml::xml_attribute<> *theAttribute);

        /*
         * =============================================================================================================
         * XMLDocument::getAttribute
         * =============================================================================================================
         */

        /**
         * @brief Get the XMLAttribute associated with the underlying xml_attribute resource.
         * @param theAttribute A const pointer to the xml_node to look up.
         * @return The const XMLNode associated with the provided xml_node resource.
         */
        virtual const XMLAttribute *getAttribute(const rapidxml::xml_attribute<> *theAttribute) const;

        /*
         * =============================================================================================================
         * Member variables
         * =============================================================================================================
         */

        std::unique_ptr<rapidxml::xml_document<>> m_document; /**< A pointer to the underlying xml_document resource. */
        std::unique_ptr<rapidxml::file<>> m_file; /**< A pointer to the underlying file resource. */
        std::string m_filePath; /**< The path of the XML file. */

        std::map<rapidxml::xml_node<> *, std::unique_ptr<XMLNode>>
            m_xmlNodes; /**< A std::map with the XMLNodes for the document, using the xml_nodes as key. */
        std::map<rapidxml::xml_attribute<> *, std::unique_ptr<XMLAttribute>>
            m_xmlAttributes; /**< A std::map with the XMLAttributes for the document, using the xml_attributes as key. */
    };

}

#endif //OPENXL_XMLDOCUMENT_H
