//
// Created by Troldal on 28/08/16.
//

#ifndef OPENXL_XMLNODE_H
#define OPENXL_XMLNODE_H

#include <string>
#include "rapidxml.hpp"
#include "XMLAttribute.h"
#include "XMLDocument.h"

namespace RapidXLSX
{
    class XMLDocument;

    class XMLAttribute;

/**
 * @brief Encapsulates the concept of an XML node
 */
    class XMLNode
    {
    public:
        /*
         * =============================================================================================================
         * XMLNode::XMLNode
         * =============================================================================================================
         */

        /**
         * @brief Constructor.
         * @param parentDocument A pointer to the parent XMLDocument object.
         * @param node A pointer to the underlying xml_node object.
         * @return Nothing
         */
        explicit XMLNode(XMLDocument *parentDocument,
                         rapidxml::xml_node<> *node);

        /*
         * =============================================================================================================
         * XMLNode::XMLNode
         * =============================================================================================================
         */

        /**
         * @brief Copy constructor
         * @param other A const reference to the original XMLNode object.
         * @return Nothing
         */
        XMLNode(const XMLNode &other);

        /*
         * =============================================================================================================
         * XMLNode::~XMLNode
         * =============================================================================================================
         */

        /**
         * @brief Destructor
         */
        virtual ~XMLNode() = default;

        /*
         * =============================================================================================================
         * XMLNode::operator=
         * =============================================================================================================
         */

        /**
         * @brief Assignment operator
         * @param other A const reference to the original XMLNode object.
         * @return A reference to the XMLNode object.
         */
        virtual XMLNode &operator=(const XMLNode &other);

        /*
         * =============================================================================================================
         * XMLNode::isValid
         * =============================================================================================================
         */

        /**
         * @brief Indicates if the object is valid, i.e. if the underlying xml_node is valid.
         * @return true if the object is valid, otherwise false.
         */
        virtual bool isValid() const;

        /*
         * =============================================================================================================
         * XMLNode::SetName
         * =============================================================================================================
         */

        /**
         * @brief Sets the name of the XMLNode.
         * @param name The new name of the XMLNode.
         */
        virtual void setName(const std::string &name);

        /*
         * =============================================================================================================
         * XMLNode::Name
         * =============================================================================================================
         */

        /**
         * @brief Get the name of the XMLNode
         * @return A const reference to the XMLNode name.
         */
        virtual const std::string &name() const;

        /*
         * =============================================================================================================
         * XMLNode::SetValue
         * =============================================================================================================
         */

        /**
         * @brief Sets the value of the XMLNode.
         * @param value The new value of the XMLNode.
         */
        virtual void setValue(const std::string &value);

        /*
         * =============================================================================================================
         * XMLNode::SetValue
         * =============================================================================================================
         */

        /**
         * @brief Sets the value of the XMLNode.
         * @param value The new value of the XMLNode.
         */
        virtual void setValue(int value);

        /*
         * =============================================================================================================
         * XMLNode::SetValue
         * =============================================================================================================
         */

        /**
         * @brief Sets the value of the XMLNode.
         * @param value The new value of the XMLNode.
         */
        virtual void setValue(double value);

        /*
         * =============================================================================================================
         * XMLNode::ValueAsString
         * =============================================================================================================
         */

        /**
         * @brief Get the value of the XMLNode
         * @return A const reference to the XMLNode value.
         */
        virtual const std::string &value() const;

        /*
         * =============================================================================================================
         * XMLNode::childNode
         * =============================================================================================================
         */

        /**
         * @brief Get the child node by name.
         * @param name the name of the XMLNode. If no name is provided, the first child node will be returned.
         * @return A pointer to the XMLNode with the given name.
         */
        virtual XMLNode *childNode(const std::string &name = "");

        /*
         * =============================================================================================================
         * XMLNode::childNode
         * =============================================================================================================
         */

        /**
         * @brief Get the child node by name.
         * @param name the name of the XMLNode. If no name is provided, the first child node will be returned.
         * @return A pointer to the XMLNode with the given name.
         */
        virtual const XMLNode *childNode(const std::string &name = "") const;

        /*
         * =============================================================================================================
         * XMLNode::parent
         * =============================================================================================================
         */

        /**
         * @brief Get a pointer to the parent XMLNode for the current object.
         * @return A pointer to an XMLNode. The returned pointer is not const, as the parent is not owned by the object.
         */
        virtual XMLNode *parent() const;

        /*
         * =============================================================================================================
         * XMLNode::previousSibling
         * =============================================================================================================
         */

        /**
         * @brief Get a pointer to the previous sibling of the current object.
         * @return A pointer to an XMLNode. The returned pointer is not const, as the sibling is not owned by the object.
         */
        virtual XMLNode *previousSibling() const;

        /*
         * =============================================================================================================
         * XMLNode::nextSibling
         * =============================================================================================================
         */

        /**
         * @brief Get a pointer to the next sibling of the current object.
         * @return A pointer to an XMLNode. The returned pointer is not const, as the sibling is not owned by the object.
         */
        virtual XMLNode *nextSibling() const;

        /*
         * =============================================================================================================
         * XMLNode::prependNode
         * =============================================================================================================
         */

        /**
         * @brief Prepend a XMLNode to the list of child nodes (if any).
         * @param node A pointer to the XMLNode to add.
         */
        virtual void prependNode(XMLNode *node);

        /*
         * =============================================================================================================
         * XMLNode::appendNode
         * =============================================================================================================
         */

        /**
         * @brief Append an XMLNode to the list of child nodes (if any).
         * @param node A pointer to the XMLNode to add.
         */
        virtual void appendNode(XMLNode *node);

        /*
         * =============================================================================================================
         * XMLNode::insertnode
         * =============================================================================================================
         */

        /**
         * @brief Insert an XMLNode at (before) the given location.
         * @param location The location to insert node.
         * @param node The node to insert.
         */
        virtual void insertNode(XMLNode *location,
                                XMLNode *node);

        virtual void moveNodeUp(XMLNode *node);

        virtual void moveNodeDown(XMLNode *node);

        virtual void moveNodeTo(XMLNode *node,
                                unsigned int index);

        /*
         * =============================================================================================================
         * XMLNode::deleteNode
         * =============================================================================================================
         */

        /**
         * @brief Delete the current node. This will delete the node from the XMLDocument and free the underlying resource.
         * Remember to set the variable to = nullptr after the call to deleteNode().
         * @todo Consider a better way to do this.
         */
        virtual void deleteNode();

        virtual void deleteChildNodes();
        /*
         * =============================================================================================================
         * XMLNode::firstAttribute
         * =============================================================================================================
         */

        /**
         * @brief Get the XMLAttribute with the given name.
         * @param name The name of the XMLAttribute (optional).
         * @return The XMLAttribute with the given name, or the first attribute if no name is provided.
         */
        virtual XMLAttribute *attribute(const std::string &name = "");

        /*
         * =============================================================================================================
         * XMLNode::firstAttribute
         * =============================================================================================================
         */

        /**
         * @brief Get the XMLAttribute with the given name.
         * @param name The name of the XMLAttribute (optional).
         * @return The XMLAttribute with the given name, or the first attribute if no name is provided.
         */
        virtual const XMLAttribute *attribute(const std::string &name = "") const;

        /*
         * =============================================================================================================
         * XMLNode::prependAttribute
         * =============================================================================================================
         */

        /**
         * @brief Prepend the XMLAttribute to the list of attributes of the XMLNode (if any).
         * @param attribute A pointer to the XMLAttribute to add.
         */
        virtual void prependAttribute(XMLAttribute *attribute);

        /*
         * =============================================================================================================
         * XMLNode::appendAttribute
         * =============================================================================================================
         */

        /**
         * @brief Append the XMLAttribute to the list of attributes of the XMLNode (if any).
     * @param attribute A pointer to the XMLAttribute to add.
         */
        virtual void appendAttribute(XMLAttribute *attribute);

    protected:
        rapidxml::xml_node<> *m_node; /**< A pointer to the underlying xml_node resource. */
        XMLDocument *m_parentXMLDocument; /**< A pointer to the parent XMLDocument. */

        mutable std::string m_valueCache;
        mutable bool m_valueLoaded;

        mutable std::string m_nameCache;
        mutable bool m_nameLoaded;
    };

}

#endif //OPENXL_XMLNODE_H
