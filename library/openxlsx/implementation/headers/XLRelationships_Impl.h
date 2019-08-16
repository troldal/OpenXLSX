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

#ifndef OPENXLSX_IMPL_XLRELATIONSHIPS_H
#define OPENXLSX_IMPL_XLRELATIONSHIPS_H

// ===== Standard Library Includes ===== //
#include <string>
#include <vector>

// ===== OpenXLSX Includes ===== //
#include "XLAbstractXMLFile_Impl.h"
#include "XLEnums_impl.h"
#include "XLXml_Impl.h"

namespace OpenXLSX::Impl
{
    class XLDocument;

    class XLRelationships;

    class XLRelationshipItem;

    // ================================================================================
    // XLRelationshipItem Class
    // ================================================================================

    /**
     * @brief An encapsulation of a relationship item, i.e. an XML file in the document, its type and an ID number.
     */
    class XLRelationshipItem
    {

        friend class XLRelationships;

    public: // ---------- Public Member Functions ---------- //

        /**
         * @brief Constructor. New items should only be created through an XLRelationship object.
         * @param node An XMLNode object with the relationship item. If no input is provided, a null node is used.
         */
        explicit XLRelationshipItem(XMLNode node = XMLNode());

        /**
         * @brief Copy Constructor.
         * @param other Object to be copied.
         */
        XLRelationshipItem(const XLRelationshipItem& other) = default;

        /**
         * @brief Move Constructor.
         * @param other Object to be moved.
         */
        XLRelationshipItem(XLRelationshipItem&& other) noexcept = default;

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         */
        XLRelationshipItem& operator=(const XLRelationshipItem& other) = default;

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         */
        XLRelationshipItem& operator=(XLRelationshipItem&& other) noexcept = default;

        /**
         * @brief Get the type of the current relationship item.
         * @return An XLRelationshipType enum object, corresponding to the type.
         */
        XLRelationshipType Type() const;

        /**
         * @brief Get the target, i.e. the path to the XML file the relationship item refers to.
         * @return An XMLAttribute object containing the Target value.
         */
        XMLAttribute Target() const;

        /**
         * @brief Get the id of the relationship item.
         * @return An XMLAttribute object containing the Id value.
         */
        XMLAttribute Id() const;

    private: // ---------- Private Member Functions ---------- //

        /**
         * @brief A static helper function for conversion of type string to XLRelationshipType.
         * @param typeString The string to convert.
         * @return A XLRelationshipType enum object.
         */
        static XLRelationshipType GetTypeFromString(const std::string& typeString);

        /**
         * @brief A static helper function for conversion of XLRelationshipType to type string.
         * @param type The XLRelationshipType object to convert
         * @return A std::string with the type
         */
        static std::string GetStringFromType(XLRelationshipType type);

    private: // ---------- Private Member Variables ---------- //

        XMLNode m_relationshipNode; /**< An XMLNode object with the relationship item */
    };



    // ================================================================================
    // XLRelationships Class
    // ================================================================================

    /**
     * @brief An encapsulation of relationship files (.rels files) in an Excel document package.
     */
    class XLRelationships : public XLAbstractXMLFile
    {

        friend class XLRelationshipItem;

    public: // ---------- Public Member Functions ---------- //

        /**
         * @brief Constructor
         * @param parent A pointer to the parent XLDocument object.
         * @param filePath The (relative) path to the relationship file.
         */
        explicit XLRelationships(XLDocument& parent, const std::string& filePath);

        /**
         * @brief Destructor
         */
        ~XLRelationships() override = default;

        /**
         * @brief Look up a relationship item by ID.
         * @param id The ID string of the relationship item to retrieve.
         * @return An XLRelationshipItem object.
         */
        XLRelationshipItem RelationshipByID(const std::string& id) const;

        /**
         * @brief Look up a relationship item by Target.
         * @param target The Target string of the relationship item to retrieve.
         * @return An XLRelationshipItem object.
         */
        XLRelationshipItem RelationshipByTarget(const std::string& target) const;

        /**
         * @brief Get the std::map with the relationship items, ordered by ID.
         * @return A const reference to the std::map with relationship items.
         */
        std::vector<const XLRelationshipItem*> Relationships() const;

        /**
         * @brief Delete an item from the Relationships register
         * @param item The XLRelationshipItem object to delete.
         */
        void DeleteRelationship(XLRelationshipItem& item);

        /**
         * @brief Add a new relationship item to the XLRelationships object.
         * @param type The type of the new relationship item.
         * @param target The target (or path) of the XML file for the relationship item.
         */
        XLRelationshipItem* AddRelationship(XLRelationshipType type, const std::string& target);

        /**
         * @brief Check if a XLRelationshipItem with the given Target string exists.
         * @param target The Target string to look up.
         * @return true if the XLRelationshipItem exists; otherwise false.
         */
        bool TargetExists(const std::string& target) const;

        /**
         * @brief Check if a XLRelationshipItem with the given Id string exists.
         * @param id The Id string to look up.
         * @return true if the XLRelationshipItem exists; otherwise false.
         */
        bool IdExists(const std::string& id) const;

    protected: // ---------- Protected Member Functions ---------- //

        /**
         * @brief Parse the contents of the underlying XML file with the relationship items.
         * @return true on success, otherwise false.
         */
        bool ParseXMLData() override;

    private: // ------------ Private Member Functions ---------- //

        /**
         * @brief Generate a new relationship ID, ensuring that every item has a unique ID.
         * @return A new ID number.
         */
        unsigned long GetNewRelsID() const;

    private: // ---------- Private Member Variables ---------- //

        std::vector<XLRelationshipItem> m_relationships; /**< A std::vector with the relationship items */
    };
} // namespace OpenXLSX::Impl

#endif //OPENXLSX_IMPL_XLRELATIONSHIPS_H
