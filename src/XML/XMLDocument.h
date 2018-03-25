//
// Created by Troldal on 29/08/16.
//

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

namespace RapidXLSX
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
