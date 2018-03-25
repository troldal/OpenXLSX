//
// Created by Troldal on 28/08/16.
//

#ifndef OPENXL_XMLATTRIBUTE_H
#define OPENXL_XMLATTRIBUTE_H

#include <string>
#include "rapidxml.hpp"
#include "XMLNode.h"
#include "XMLDocument.h"

namespace RapidXLSX
{
    class XMLNode;

    class XMLDocument;

/**
 * @brief Encapsulates the concept of an XML attribute
 */
    class XMLAttribute
    {
        friend class XMLNode;

    public:
        /*
         * =============================================================================================================
         * XMLAttribute::XMLAttribute
         * =============================================================================================================
         */

        /**
         * @brief Constructor
         * @param parentDocument A pointer to the parent XMLDocument.
         * @param attribute A pointer to the underlying xml_attribute resource.
         * @return Nothing
         */
        explicit XMLAttribute(XMLDocument *parentDocument,
                              rapidxml::xml_attribute<> *attribute);

        /*
         * =============================================================================================================
         * XMLAttribute::XMLAttribute
         * =============================================================================================================
         */

        /**
         * @brief Copy constructor
         * @param other Reference to the object to copy
         * @return Nothing
         */
        XMLAttribute(const XMLAttribute &other);

        /*
         * =============================================================================================================
         * XMLAttribute::~XMLAttribute
         * =============================================================================================================
         */

        /**
         * @brief Dectructor
         */
        virtual ~XMLAttribute();

        /*
         * =============================================================================================================
         * XMLAttribute::operator=
         * =============================================================================================================
         */

        /**
         * @brief Assignment operator.
         * @param other Reference to the original object
         * @return A reference to the new object
         */
        virtual XMLAttribute &operator=(const XMLAttribute &other);

        /*
         * =============================================================================================================
         * XMLAttribute::bool
         * =============================================================================================================
         */

        /**
         * @brief
         * @return
         * @todo For some reason, this doesn't work properly.
         */
        virtual operator bool() const;

        /*
         * =============================================================================================================
         * XMLAttribute::isValid
         * =============================================================================================================
         */

        /**
         * @brief Determines hether or not the object is valid, i.e. that the underlying xml_attribute is valid.
         * @return true if the object is valid, otherwise false.
         */
        virtual bool isValid() const;


        /*
         * =============================================================================================================
         * XMLAttribute::SetName
         * =============================================================================================================
         */

        /**
         * @brief Sets the name of the XMLAttribute
         * @param name The name
         */
        virtual void setName(const std::string &name);

        /*
         * =============================================================================================================
         * XMLAttribute::Name
         * =============================================================================================================
         */

        /**
         * @brief Gets the name of the XMLAttribute
         * @return The name
         */
        virtual const std::string &name() const;

        /*
         * =============================================================================================================
         * XMLAttribute::SetValue
         * =============================================================================================================
         */

        /**
         * @brief Sets the value of the XMLAttribute.
         * @param value The value
         */
        virtual void setValue(const std::string &value);

        /*
         * =============================================================================================================
         * XMLAttribute::SetValue
         * =============================================================================================================
         */

        /**
         * @brief Sets the value of the XMLAttribute.
         * @param value The value
         */
        virtual void setValue(int value);

        /*
         * =============================================================================================================
         * XMLAttribute::SetValue
         * =============================================================================================================
         */

        /**
         * @brief Sets the value of the XMLAttribute.
         * @param value The value
         */
        virtual void setValue(double value);

        /*
         * =============================================================================================================
         * XMLAttribute::ValueAsString
         * =============================================================================================================
         */

        /**
         * @brief Gets the value of the XMLAttribute.
         * @return The value
         */
        virtual const std::string &value() const;

        /*
         * =============================================================================================================
         * XMLAttribute::parent
         * =============================================================================================================
         */

        /**
         * @brief Get a pointer to the parent XMLNode for the current object.
         * @return A pointer to an XMLNode. The returned pointer is not const, as the parent is not owned by the object.
         */
        virtual XMLNode *parent() const;

        /*
         * =============================================================================================================
         * XMLAttribute::previousAttribute
         * =============================================================================================================
         */

        /**
         * @brief Get a pointer to the previous attribute of the current object.
         * @return A pointer to an XMLAttribute. The returned pointer is not const, as the sibling is not owned by the object.
         */
        virtual XMLAttribute *previousAttribute() const;

        /*
         * =============================================================================================================
         * XMLAttribute::nextAttribute
         * =============================================================================================================
         */

        /**
         * @brief Get a pointer to the next attribute of the current object.
         * @return A pointer to an XMLAttribute. The returned pointer is not const, as the sibling is not owned by the object.
         */
        virtual XMLAttribute *nextAttribute() const;

        /*
         * =============================================================================================================
         * XMLAttribute::deleteAttribute
         * =============================================================================================================
         */

        /**
         * @brief Delete the current attribute. This will delete the node from the XMLDocument and free the underlying resource.
         * Remember to set the variable to = nullptr after the call to deleteNode().
         */
        virtual void deleteAttribute();

    protected:
        /*
         * =============================================================================================================
         * Member Variables
         * =============================================================================================================
         */

        rapidxml::xml_attribute<> *m_attribute; /**< A pointer to the underlying xml_attribute resource. */
        XMLDocument *m_parentXMLDocument; /**< A pointer to the parent XMLDocument object. */

        mutable std::string m_valueCache;
        mutable bool m_valueLoaded;

        mutable std::string m_nameCache;
        mutable bool m_nameLoaded;
    };

}

#endif //OPENXL_XMLATTRIBUTE_H
