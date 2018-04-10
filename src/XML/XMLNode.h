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

#ifndef OPENXL_XMLNODE_H
#define OPENXL_XMLNODE_H

#include <string>
#include "rapidxml.hpp"
#include "XMLAttribute.h"
#include "XMLDocument.h"

namespace OpenXLSX
{
    class XMLDocument;
    class XMLAttribute;

//======================================================================================================================
//========== XMLNode Class =============================================================================================
//======================================================================================================================

    /**
     * @brief Encapsulates the concept of an XML node
     */
    class XMLNode final
    {
        friend class XMLDocument;
        friend class XMLAttribute;

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Copy constructor
         * @param other A const reference to the original XMLNode object.
         * @return Nothing
         */
        XMLNode(const XMLNode &other) = default;

        /**
         * @brief Destructor
         */
        ~XMLNode() = default;

        /**
         * @brief Assignment operator
         * @param other A const reference to the original XMLNode object.
         * @return A reference to the XMLNode object.
         */
        XMLNode &operator=(const XMLNode &other) = default;

        /**
         * @brief Indicates if the object is valid, i.e. if the underlying xml_node is valid.
         * @return true if the object is valid, otherwise false.
         */
        bool IsValid() const;

        /**
         * @brief Sets the name of the XMLNode.
         * @param name The new name of the XMLNode.
         */
        void SetName(const std::string &name);

        /**
         * @brief Get the name of the XMLNode
         * @return A const reference to the XMLNode name.
         */
        const std::string &Name() const;

        /**
         * @brief Sets the value of the XMLNode.
         * @param value The new value of the XMLNode.
         */
        void SetValue(const std::string &value);

        /**
         * @brief Sets the value of the XMLNode.
         * @param value The new value of the XMLNode.
         */
        void SetValue(int value);

        /**
         * @brief Sets the value of the XMLNode.
         * @param value The new value of the XMLNode.
         */
        void SetValue(double value);

        /**
         * @brief Get the value of the XMLNode
         * @return A const reference to the XMLNode value.
         */
        const std::string &Value() const;

        /**
         * @brief Get the child node by name.
         * @param name the name of the XMLNode. If no name is provided, the first child node will be returned.
         * @return A pointer to the XMLNode with the given name.
         */
        XMLNode *ChildNode(const std::string &name = "");

        /**
         * @brief Get the child node by name.
         * @param name the name of the XMLNode. If no name is provided, the first child node will be returned.
         * @return A pointer to the XMLNode with the given name.
         */
        const XMLNode *ChildNode(const std::string &name = "") const;

        /**
         * @brief Get a pointer to the parent XMLNode for the current object.
         * @return A pointer to an XMLNode. The returned pointer is not const, as the parent is not owned by the object.
         */
        XMLNode *Parent() const;

        /**
         * @brief Get a pointer to the previous sibling of the current object.
         * @return A pointer to an XMLNode. The returned pointer is not const, as the sibling is not owned by the object.
         */
        XMLNode *PreviousSibling() const;

        /**
         * @brief Get a pointer to the next sibling of the current object.
         * @return A pointer to an XMLNode. The returned pointer is not const, as the sibling is not owned by the object.
         */
        XMLNode *NextSibling() const;

        /**
         * @brief Prepend a XMLNode to the list of child nodes (if any).
         * @param node A pointer to the XMLNode to add.
         */
        void PrependNode(XMLNode *node);

        /**
         * @brief Append an XMLNode to the list of child nodes (if any).
         * @param node A pointer to the XMLNode to add.
         */
        void AppendNode(XMLNode *node);

        /**
         * @brief Insert an XMLNode at (before) the given location.
         * @param location The location to insert node.
         * @param node The node to insert.
         */
        void InsertNode(XMLNode *location, XMLNode *node);

        /**
         * @brief
         * @param node
         */
        void MoveNodeUp(XMLNode *node);

        /**
         * @brief
         * @param node
         */
        void MoveNodeDown(XMLNode *node);

        /**
         * @brief
         * @param node
         * @param index
         */
        void MoveNodeTo(XMLNode *node, unsigned int index);

        /**
         * @brief Delete the current node. This will delete the node from the XMLDocument and free the underlying resource.
         * Remember to set the variable to = nullptr after the call to deleteNode().
         * @todo Consider a better way to do this.
         */
        void DeleteNode();

        /**
         * @brief
         */
        void DeleteChildNodes();

        /**
         * @brief Get the XMLAttribute with the given name.
         * @param name The name of the XMLAttribute (optional).
         * @return The XMLAttribute with the given name, or the first attribute if no name is provided.
         */
        XMLAttribute *Attribute(const std::string &name = "");

        /**
         * @brief Get the XMLAttribute with the given name.
         * @param name The name of the XMLAttribute (optional).
         * @return The XMLAttribute with the given name, or the first attribute if no name is provided.
         */
        const XMLAttribute *Attribute(const std::string &name = "") const;

        /**
         * @brief Prepend the XMLAttribute to the list of attributes of the XMLNode (if any).
         * @param attribute A pointer to the XMLAttribute to add.
         */
        void PrependAttribute(XMLAttribute *attribute);

        /**
         * @brief Append the XMLAttribute to the list of attributes of the XMLNode (if any).
         * @param attribute A pointer to the XMLAttribute to add.
         */
        void AppendAttribute(XMLAttribute *attribute);

//----------------------------------------------------------------------------------------------------------------------
//           Protected Member Functions
//----------------------------------------------------------------------------------------------------------------------

    protected:

        /**
         * @brief Constructor.
         * @param parentDocument A pointer to the parent XMLDocument object.
         * @param node A pointer to the underlying xml_node object.
         * @return Nothing
         */
        explicit XMLNode(XMLDocument *parentDocument,
                         rapidxml::xml_node<> *node);

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:
        rapidxml::xml_node<> *m_node; /**< A pointer to the underlying xml_node resource. */
        XMLDocument *m_parentXMLDocument; /**< A pointer to the parent XMLDocument. */

        mutable std::string m_valueCache; /**< */
        mutable bool m_valueLoaded; /**< */

        mutable std::string m_nameCache; /**< */
        mutable bool m_nameLoaded; /**< */
    };

}

#endif //OPENXL_XMLNODE_H
