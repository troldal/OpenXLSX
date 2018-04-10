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

#ifndef OPENXL_XMLATTRIBUTE_H
#define OPENXL_XMLATTRIBUTE_H

#include <string>
#include "rapidxml.hpp"
#include "XMLNode.h"
#include "XMLDocument.h"

namespace OpenXLSX
{
    class XMLNode;
    class XMLDocument;

//======================================================================================================================
//========== XMLAttribute Class =========================================================================================
//======================================================================================================================

    /**
     * @brief Encapsulates the concept of an XML attribute
     */
    class XMLAttribute
    {
        friend class XMLNode;

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor
         * @param parentDocument A pointer to the parent XMLDocument.
         * @param attribute A pointer to the underlying xml_attribute resource.
         * @return Nothing
         */
        explicit XMLAttribute(XMLDocument *parentDocument,
                              rapidxml::xml_attribute<> *attribute);

        /**
         * @brief Copy constructor
         * @param other Reference to the object to copy
         * @return Nothing
         */
        XMLAttribute(const XMLAttribute &other);

        /**
         * @brief Dectructor
         */
        virtual ~XMLAttribute() = default;

        /**
         * @brief Assignment operator.
         * @param other Reference to the original object
         * @return A reference to the new object
         */
        virtual XMLAttribute &operator=(const XMLAttribute &other);

        /**
         * @brief
         * @return
         * @todo For some reason, this doesn't work properly.
         */
        virtual operator bool() const;

        /**
         * @brief Determines hether or not the object is valid, i.e. that the underlying xml_attribute is valid.
         * @return true if the object is valid, otherwise false.
         */
        virtual bool isValid() const;

        /**
         * @brief Sets the name of the XMLAttribute
         * @param name The name
         */
        virtual void setName(const std::string &name);

        /**
         * @brief Gets the name of the XMLAttribute
         * @return The name
         */
        virtual const std::string &name() const;

        /**
         * @brief Sets the value of the XMLAttribute.
         * @param value The value
         */
        virtual void setValue(const std::string &value);

        /**
         * @brief Sets the value of the XMLAttribute.
         * @param value The value
         */
        virtual void setValue(int value);

        /**
         * @brief Sets the value of the XMLAttribute.
         * @param value The value
         */
        virtual void setValue(double value);

        /**
         * @brief Gets the value of the XMLAttribute.
         * @return The value
         */
        virtual const std::string &value() const;

        /**
         * @brief Get a pointer to the parent XMLNode for the current object.
         * @return A pointer to an XMLNode. The returned pointer is not const, as the parent is not owned by the object.
         */
        virtual XMLNode *parent() const;

        /**
         * @brief Get a pointer to the previous attribute of the current object.
         * @return A pointer to an XMLAttribute. The returned pointer is not const, as the sibling is not owned by the object.
         */
        virtual XMLAttribute *previousAttribute() const;

        /**
         * @brief Get a pointer to the next attribute of the current object.
         * @return A pointer to an XMLAttribute. The returned pointer is not const, as the sibling is not owned by the object.
         */
        virtual XMLAttribute *nextAttribute() const;

        /**
         * @brief Delete the current attribute. This will delete the node from the XMLDocument and free the underlying resource.
         * Remember to set the variable to = nullptr after the call to deleteNode().
         */
        virtual void deleteAttribute();

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:

        rapidxml::xml_attribute<> *m_attribute; /**< A pointer to the underlying xml_attribute resource. */
        XMLDocument *m_parentXMLDocument; /**< A pointer to the parent XMLDocument object. */

        mutable std::string m_valueCache; /**< */
        mutable bool m_valueLoaded; /**< */

        mutable std::string m_nameCache; /**< */
        mutable bool m_nameLoaded; /**< */
    };

}

#endif //OPENXL_XMLATTRIBUTE_H
