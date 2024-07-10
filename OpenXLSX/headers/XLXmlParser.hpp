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

#ifndef OPENXLSX_XLXMLPARSER_HPP
#define OPENXLSX_XLXMLPARSER_HPP

// ===== pugixml.hpp needed for pugi::impl::xml_memory_page_type_mask, pugi::xml_node_type, pugi::char_t, pugi::node_element, pugi::xml_node, pugi::xml_attribute, pugi::xml_document
#include <external/pugixml/pugixml.hpp> // not sure why the full include path is needed within the header file

// 2024-05-29: forward declarations now pointless with include on pugixml being needed anyways
// namespace pugi
// {
//     class xml_node;
//     class xml_attribute;
//     class xml_document;
// }    // namespace pugi

namespace OpenXLSX
{
    // ===== Copy definition of PUGI_IMPL_NODETYPE, which is defined in pugixml.cpp, within a namespace, and somehow doesn't work here
#   define PUGI_IMPL_NODETYPE(n) static_cast<pugi::xml_node_type>((n)->header & pugi::impl::xml_memory_page_type_mask)


    // disable this line to use original (non-augmented) pugixml
#   define PUGI_AUGMENTED

    // ===== Using statements to switch between pugixml and augmented pugixml implementation
#   ifdef PUGI_AUGMENTED
        // ===== Forward declarations for using statements below
        class OpenXLSX_xml_node;
        class OpenXLSX_xml_document;

        using XMLNode      = OpenXLSX_xml_node;
        using XMLAttribute = pugi::xml_attribute;
        using XMLDocument  = OpenXLSX_xml_document;
#   else
        using XMLNode      = pugi::xml_node;
        using XMLAttribute = pugi::xml_attribute;
        using XMLDocument  = pugi::xml_document;
#   endif

    // ===== Custom OpenXLSX_xml_node to add functionality to pugi::xml_node
    class OpenXLSX_xml_node : public pugi::xml_node {
    public:
        /**
         * @brief Default constructor. Constructs a null object.
         */
        OpenXLSX_xml_node() : pugi::xml_node() {}

        /**
         * @brief Inherit all constructors with parameters from pugi::xml_node
         */
        template<class base>
        // explicit OpenXLSX_xml_node(base b) : xml_node(b) // TBD
        OpenXLSX_xml_node(base b) : pugi::xml_node(b)
        {}

        // ===== BEGIN: Wrappers for xml_node member functions to ensure OpenXLSX_xml_node return values
        // ===== CAUTION: this section is incomplete, only implementing those functions actually used by OpenXLSX to date
        /**
         * @brief for all functions: invoke the base class function, but with a return type of OpenXLSX_xml_node
         */
        XMLNode parent() { return pugi::xml_node::parent(); }
        XMLNode child(const pugi::char_t* name) const { return pugi::xml_node::child(name); }
        template <typename Predicate> XMLNode find_child(Predicate pred) const { return pugi::xml_node::find_child(pred); }
        // ===== END: Wrappers for xml_node member functions

        /**
         * @brief get first node child that matches type
         * @param type_ the pugi::xml_node_type to match
         * @return a valid child matching the node type or an empty XMLNode
         */
        XMLNode first_child_of_type(pugi::xml_node_type type_ = pugi::node_element) const;

        /**
         * @brief get last node child that matches type
         * @param type_ the pugi::xml_node_type to match
         * @return a valid child matching the node type or an empty XMLNode
         */
        XMLNode last_child_of_type(pugi::xml_node_type type_ = pugi::node_element) const;

        /**
         * @brief count node children that match type
         * @param type_ the pugi::xml_node_type to match
         * @return the amount of node children matching type
         */
        size_t child_count_of_type(pugi::xml_node_type type_ = pugi::node_element) const;

        /**
         * @brief get next node sibling that matches type
         * @param type_ the pugi::xml_node_type to match
         * @return a valid sibling matching the node type or an empty XMLNode
         */
        XMLNode next_sibling_of_type(pugi::xml_node_type type_ = pugi::node_element) const;

        /**
         * @brief get previous node sibling that matches type
         * @param type_ the pugi::xml_node_type to match
         * @return a valid sibling matching the node type or an empty XMLNode
         */
        XMLNode previous_sibling_of_type(pugi::xml_node_type type_ = pugi::node_element) const;

        /**
         * @brief get next node sibling that matches name_ and type
         * @param name_ the xml_node::name() to match
         * @param type_ the pugi::xml_node_type to match
         * @return a valid sibling matching the node type or an empty XMLNode
         */
        XMLNode next_sibling_of_type(const pugi::char_t* name_, pugi::xml_node_type type_ = pugi::node_element) const;

        /**
         * @brief get previous node sibling that matches name_ and type
         * @param name_ the xml_node::name() to match
         * @param type_ the pugi::xml_node_type to match
         * @return a valid sibling matching the node type or an empty XMLNode
         */
        XMLNode previous_sibling_of_type(const pugi::char_t* name_, pugi::xml_node_type type_ = pugi::node_element) const;
    };

    // ===== Custom OpenXLSX_xml_document to override relevant pugi::xml_document member functions with OpenXLSX_xml_node return value
    class OpenXLSX_xml_document : public pugi::xml_document {
    public:
        /**
         * @brief Default constructor. Constructs a null object.
         */
        OpenXLSX_xml_document() : pugi::xml_document() {}

        /**
         * @brief Inherit all constructors with parameters from pugi::xml_document
         */
        template<class base>
        // explicit OpenXLSX_xml_document(base b) : xml_document(b) // TBD
        OpenXLSX_xml_document(base b) : pugi::xml_document(b)
        {}

        // ===== BEGIN: Wrappers for xml_document member functions to ensure OpenXLSX_xml_node return values
        // ===== CAUTION: this section is incomplete, only implementing those functions actually used by OpenXLSX to date
        /**
         * @brief for all functions: invoke the base class function, but with a return type of OpenXLSX_xml_node
         */
        XMLNode document_element() const { return pugi::xml_document::document_element(); }
        // ===== END: Wrappers for xml_document member functions
    };

}    // namespace OpenXLSX
#endif    // OPENXLSX_XLXMLPARSER_HPP
