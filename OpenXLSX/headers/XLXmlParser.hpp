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

#include <memory> // shared_ptr

// ===== pugixml.hpp needed for pugi::impl::xml_memory_page_type_mask, pugi::xml_node_type, pugi::node_element, pugi::xml_node, pugi::xml_attribute, pugi::xml_document
#include <external/pugixml/pugixml.hpp> // not sure why the full include path is needed within the header file
#include "XLException.hpp"

namespace { // anonymous namespace to define constants / functions that shall not be exported from this module
    constexpr const int XLMaxNamespacedNameLen = 100;
} // anonymous namespace

namespace OpenXLSX
{
    typedef char char_t;   // Character type used for all internal storage and operations; single location to change it everywhere

    enum class xml_node_type // declare OpenXLSX interface to pugi::xml_node_type
    {
        node_null,            // Empty (null) node handle
        node_document,        // A document tree's absolute root
        node_element,        // Element tag, i.e. '<node/>'
        node_pcdata,        // Plain character data, i.e. 'text'
        node_cdata,            // Character data, i.e. '<![CDATA[text]]>'
        node_comment,        // Comment tag, i.e. '<!-- text -->'
        node_pi,            // Processing instruction, i.e. '<?name?>'
        node_declaration,    // Document declaration, i.e. '<?xml version="1.0"?>'
        node_doctype        // Document type declaration, i.e. '<!DOCTYPE doc>'
    };


#   define ENABLE_XML_NAMESPACES 1    // disable this line to control behavior via compiler flag
#   define NO_MULTITHREADING_SAFETY 1 // if this is defined, the function namespaced_name_char will be used for XML namespace support,
//                                    //  using a static character array for improved performance when multithreading conflicts are not a risk

#   ifdef ENABLE_XML_NAMESPACES
        // ===== Macro for NAMESPACED_NAME when node names might need to be prefixed with the current node's namespace
#       ifdef NO_MULTITHREADING_SAFETY
#           define NAMESPACED_NAME(name_, force_ns_) namespaced_name_char(name_, force_ns_)
#       else
#           define NAMESPACED_NAME(name_, force_ns_) namespaced_name_shared_ptr(name_, force_ns_).get()
#       endif
#   else
        // ===== Optimized version when no namespace support is desired - ignores force_ns_ setting
#       define NAMESPACED_NAME(name_, force_ns_) name_
#   endif

    constexpr const bool XLForceNamespace = true;
    /**
     * With namespace support: where OpenXLSX already addresses nodes by their namespace, doubling the namespace in a node name
     *   upon node create can be avoided by passing the optional parameter XLForceNamespace - this will use the node name passed to the
     *   insertion function "as-is".
     * Affected XMLNode methods: ::set_name, ::append_child, ::prepend_child, ::insert_child_after, ::insert_child_before
     */

    extern bool NO_XML_NS; // defined in XLXmlParser.cpp - default: no XML namespaces
    /**
     * @brief Set NO_XML_NS to false
     * @return always true
     * @note CAUTION: this setting should be established before any other OpenXLSX function is used
     */
    bool enable_xml_namespaces();
    /**
     * @brief Set NO_XML_NS to true
     * @return always true
     * @note CAUTION: this setting should be established before any other OpenXLSX function is used
     */
    bool disable_xml_namespaces();

    // ===== Using statements to switch between pugixml and augmented pugixml implementation

    // ===== Forward declarations for using statements below
    class OpenXLSX_xml_node;
    class OpenXLSX_xml_document;

    using XMLNode      = OpenXLSX_xml_node;
    using XMLAttribute = pugi::xml_attribute;
    using XMLDocument  = OpenXLSX_xml_document;

    // ===== Custom OpenXLSX_xml_node to add functionality to pugi::xml_node
    class OpenXLSX_xml_node : public pugi::xml_node {
    public:
        /**
         * @brief Default constructor. Constructs a null object.
         */
        OpenXLSX_xml_node() : pugi::xml_node(), name_begin(0) {}

        /**
         * @brief Inherit all constructors with parameters from pugi::xml_node
         */
        template<class base>
        // explicit OpenXLSX_xml_node(base b) : pugi::xml_node(b), name_begin(0) // TBD on explicit keyword
        OpenXLSX_xml_node(base b) : pugi::xml_node(b), name_begin(0)
        {
            if (NO_XML_NS) return;
                const char *name = xml_node::name();
            int pos = 0;
            while (name[pos] && name[pos] != ':') ++pos; // find name delimiter
            if (name[pos] == ':') name_begin = pos + 1;  // if delimiter was found: update name_begin to point behind that position
        }

        /**
         * @brief Strip any namespace from name_
         * @param name_ A node name which may be prefixed with any namespace like so "namespace:nodename"
         * @return The name_ stripped of a namespace prefix
         */
        const char_t* name_without_namespace(const char_t* name_) const;

        /**
         * @brief add this node's namespace to name_
         * @param name_ a node name which shall be prefixed with this node's current namespace
         * @param force_ns if true, will return name_ unmodified
         * @return this node's current namespace + ":" + name_ as a const char_t *
         */
        const char_t* namespaced_name_char(const char_t* name_, bool force_ns) const;

        /**
         * @brief add this node's namespace to name_
         * @param name_ a node name which shall be prefixed with this node's current namespace
         * @param force_ns if true, will return name_ unmodified
         * @return this node's current namespace + ":" + name_ as a shared_ptr to char_t
         */
        std::shared_ptr<char_t> namespaced_name_shared_ptr(const char_t* name_, bool force_ns) const;


        // ===== BEGIN: Wrappers for xml_node member functions to ensure OpenXLSX_xml_node return values
        //                and overrides for xml_node member functions to support ignoring the node namespace
        // ===== CAUTION: this section might be incomplete, only functions actually used by OpenXLSX to date have been checked

        /**
         * @brief get node name while stripping namespace, if so configured (name_begin > 0)
         * @return the node name without a namespace
         */
        const char_t* name() const;

        /**
         * @brief get actual node name, including namespace, if any
         * @return the raw node name
         */
        const char_t* raw_name() const;

        /**
         * @brief get xml_node_type
         * @return the raw node name
         */
        xml_node_type type() const;

        /**
         * @brief for all functions returning xml_node: invoke base class function, but with a return type of XMLNode (OpenXLSX_xml_node)
         */
        XMLNode parent();

        /**
         * @brief wrapper for finding a node based on a predicate
         * @note find_child is used exactly once in XLSheet.cpp line 1222
         */
        template <typename Predicate> XMLNode find_child(Predicate pred) const {
            if (!_root) return XMLNode();

            for (XMLNode node = first_child(); node; node = node.next_sibling())
                if (pred(node))
                    return node;

            return XMLNode();
        }
        XMLNode child(const char_t* name_) const;
        XMLNode next_sibling(const char_t* name_) const;
        XMLNode next_sibling() const;
        XMLNode previous_sibling(const char_t* name_) const;
        XMLNode previous_sibling() const;
        const char_t* child_value() const;
        const char_t* child_value(const char_t* name_) const;
        bool set_name(const char_t* rhs, bool force_ns = false);
        bool set_name(const char_t* rhs, size_t size, bool force_ns = false);
        XMLNode append_child(xml_node_type type_);
        XMLNode prepend_child(xml_node_type type_);
        XMLNode append_child(const char_t* name_, bool force_ns = false);
        XMLNode prepend_child(const char_t* name_, bool force_ns = false);
        XMLNode insert_child_after(xml_node_type type_, const xml_node& node);
        XMLNode insert_child_before(xml_node_type type_, const xml_node& node);
        XMLNode insert_child_after(const char_t* name_, const xml_node& node, bool force_ns = false);
        XMLNode insert_child_before(const char_t* name_, const xml_node& node, bool force_ns = false);
        bool remove_child(const char_t* name_);
        bool remove_child(const xml_node& n);
        XMLNode find_child_by_attribute(const char_t* name_, const char_t* attr_name, const char_t* attr_value) const;
        XMLNode find_child_by_attribute(const char_t* attr_name, const char_t* attr_value) const;
        // DISCLAIMER: unused by OpenXLSX, but theoretically impacted by xml namespaces:
        // PUGI_IMPL_FN xml_node xml_node::first_element_by_path(const char_t* path_, char_t delimiter) const
        // ===== END: Wrappers for xml_node member functions

        /**
         * @brief get first node child that matches type
         * @param type_ the xml_node_type to match
         * @return a valid child matching the node type or an empty XMLNode
         */
        XMLNode first_child_of_type(xml_node_type type_ = xml_node_type::node_element) const;

        /**
         * @brief get last node child that matches type
         * @param type_ the xml_node_type to match
         * @return a valid child matching the node type or an empty XMLNode
         */
        XMLNode last_child_of_type(xml_node_type type_ = xml_node_type::node_element) const;

        /**
         * @brief count node children that match type
         * @param type_ the xml_node_type to match
         * @return the amount of node children matching type
         */
        size_t child_count_of_type(xml_node_type type_ = xml_node_type::node_element) const;

        /**
         * @brief get next node sibling that matches type
         * @param type_ the xml_node_type to match
         * @return a valid sibling matching the node type or an empty XMLNode
         */
        XMLNode next_sibling_of_type(xml_node_type type_ = xml_node_type::node_element) const;

        /**
         * @brief get previous node sibling that matches type
         * @param type_ the xml_node_type to match
         * @return a valid sibling matching the node type or an empty XMLNode
         */
        XMLNode previous_sibling_of_type(xml_node_type type_ = xml_node_type::node_element) const;

        /**
         * @brief get next node sibling that matches name_ and type
         * @param name_ the xml_node::name() to match
         * @param type_ the xml_node_type to match
         * @return a valid sibling matching the node type or an empty XMLNode
         */
        XMLNode next_sibling_of_type(const char_t* name_, xml_node_type type_ = xml_node_type::node_element) const;

        /**
         * @brief get previous node sibling that matches name_ and type
         * @param name_ the xml_node::name() to match
         * @param type_ the xml_node_type to match
         * @return a valid sibling matching the node type or an empty XMLNode
         */
        XMLNode previous_sibling_of_type(const char_t* name_, xml_node_type type_ = xml_node_type::node_element) const;

    private:
        int  name_begin; // nameBegin holds the position in xml_node::name() where the actual node name begins - 0 for non-namespaced nodes
                         // for nodes with a namespace: the position following namespace + delimiter colon, e.g. "x:c" -> nameBegin = 2
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
        XMLNode document_element() const;
        XMLNode prepend_child(xml_node_type type_);
        XMLNode prepend_child(const char_t* name_);
        // ===== END: Wrappers for xml_document member functions
    };

}    // namespace OpenXLSX
#endif    // OPENXLSX_XLXMLPARSER_HPP
