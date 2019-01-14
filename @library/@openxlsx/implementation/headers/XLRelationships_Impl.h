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

#include <string>
#include <vector>
#include <map>

#include "XLAbstractXMLFile_Impl.h"
#include "XLSpreadsheetElement_Impl.h"
#include "XLXml_Impl.h"

namespace OpenXLSX::Impl
{
    class XLDocument;
    class XLRelationships;
    class XLRelationshipItem;

    using XLRelationshipMap = std::map<std::string, std::unique_ptr<XLRelationshipItem>>;


//======================================================================================================================
//========== XLRelationshipType Enum ===================================================================================
//======================================================================================================================

/**
 * @brief An enum of the possible relationship (or XML document) types used in relationship (.rels) XML files.
 */
    enum class XLRelationshipType
    {
        CoreProperties,
        ExtendedProperties,
        CustomProperties,
        Workbook,
        Worksheet,
        ChartSheet,
        DialogSheet,
        MacroSheet,
        CalculationChain,
        ExternalLink,
        ExternalLinkPath,
        Theme,
        Styles,
        Chart,
        ChartStyle,
        ChartColorStyle,
        Image,
        Drawing,
        VMLDrawing,
        SharedStrings,
        PrinterSettings,
        VBAProject,
        ControlProperties,
        Unknown
    };


//======================================================================================================================
//========== XLRelationshipItem Class ==================================================================================
//======================================================================================================================

/**
 * @brief An encapsulation of a relationship item, i.e. an XML file in the document, its type and an ID number.
 */
    class XLRelationshipItem
    {
        friend class XLRelationships;

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Copy Constructor
         * @param other Object to be copied
         * @note The copy constructor has been explicitly deleted
         */
        XLRelationshipItem(const XLRelationshipItem &other) = delete;

        /**
         * @brief Move Constructor
         * @param other Object to be moved
         * @note The move constructor has been explicitly deleted
         */
        XLRelationshipItem(XLRelationshipItem &&other) noexcept = delete;

        /**
         * @brief Copy assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to the lhs object.
         * @note The copy assignment operator has been explicitly deleted
         */
        XLRelationshipItem &operator=(const XLRelationshipItem &other) = delete;

        /**
         * @brief Move assignment operator.
         * @param other Right hand side of assignment operation.
         * @return A reference to lhs object.
         * @note The move assignment operator has been explicitly deleted
         */
        XLRelationshipItem &operator=(XLRelationshipItem &&other) noexcept = delete;

        /**
         * @brief Get the type of the current relationship item.
         * @return An XLRelationshipType enum object, corresponding to the type.
         */
        XLRelationshipType Type() const;

        /**
         * @brief Get the target, i.e. the path to the XML file the relationship item refers to.
         * @return A const reference to the target string.
         */
        const std::string &Target() const;

        /**
         * @brief Get the id of the relationship item.
         * @return A const reference to the id string.
         */
        const std::string &Id() const;


//----------------------------------------------------------------------------------------------------------------------
//           Protected Member Functions
//----------------------------------------------------------------------------------------------------------------------

    protected:

        /**
         * @brief Delete the current relationship item.
         * @todo Should this be done in the parent XLRelationship object?
         * @todo Is there a more elegant way to do this?
         */
        void Delete();

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Functions
//----------------------------------------------------------------------------------------------------------------------

    private:

        /**
         * @brief Constructor (Private). New items should only be created through an XLRelationship object.
         * @param node A pointer to the XML node with thr relationship item.
         * @param type The type of the relationship item
         * @param target The target of the relationship item
         * @param id The id of the relationship item
         */
        XLRelationshipItem(XMLNode node,
                           XLRelationshipType type,
                           const std::string &target,
                           const std::string &id);

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:

        std::unique_ptr<XMLNode> m_relationshipNode; /**< A pointer to the XML node with the relationship item */
        XLRelationshipType m_relationshipType; /**< The type of the relationship item */
        std::string m_relationshipTarget; /**< The target of the relationship item */
        std::string m_relationshipId; /**< The ID of the relationship item */
    };



//======================================================================================================================
//========== XLRelationships Class =====================================================================================
//======================================================================================================================

/**
 * @brief An encapsulation of relationship files (.rels files) in an Excel document package.
 */
    class XLRelationships: public XLAbstractXMLFile,
                           public XLSpreadsheetElement
    {
        friend class XLRelationshipItem;
//----------------------------------------------------------------------------------------------------------------------
//          Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief Constructor
         * @param parent A pointer to the parent XLDocument object.
         * @param filePath The (relative) path to the relationship file.
         */
        explicit XLRelationships(XLDocument &parent,
                                 const std::string &filePath);

        /**
         * @brief Destructor
         */
        ~XLRelationships() override = default;

        /**
         * @brief Look up a relationship item by ID.
         * @param id The ID string of the relationship item to retrieve.
         * @return A pointer-to-const XLRelationshipItem object.
         */
        const XLRelationshipItem *RelationshipByID(const std::string &id) const;

        /**
         * @brief Look up a relationship item by ID.
         * @param id The ID string of the relationship item to retrieve.
         * @return A pointer to the XLRelationshipItem object.
         */
        XLRelationshipItem *RelationshipByID(const std::string &id);

        /**
         * @brief
         * @param target
         * @return
         */
        const XLRelationshipItem *RelationshipByTarget(const std::string &target) const;

        /**
         * @brief
         * @param target
         * @return
         */
        XLRelationshipItem *RelationshipByTarget(const std::string &target);

        /**
         * @brief Get the std::map with the relationship items, ordered by ID.
         * @return A const reference to the std::map with relationship items.
         */
        const XLRelationshipMap *Relationships() const;

        /**
         * @brief Delete a relationship item with the given ID.
         * @param id The ID of the relationship item to delete.
         */
        void DeleteRelationship(const std::string &id);

        /**
         * @brief Add a new relationship item to the XLRelationships object.
         * @param type The type of the new relationship item.
         * @param target The target (or path) of the XML file for the relationship item.
         */
        XLRelationshipItem *AddRelationship(XLRelationshipType type,
                                             const std::string &target);

        /**
         * @brief
         * @param target
         * @return
         */
        bool TargetExists(const std::string &target) const;

        /**
         * @brief
         * @param id
         * @return
         */
        bool IdExists(const std::string &id) const;

//----------------------------------------------------------------------------------------------------------------------
//           Protected Member Functions
//----------------------------------------------------------------------------------------------------------------------

    protected:

        /**
         * @brief Parse the contents of the underlying XML file with the relationship items.
         * @return true on success, otherwise false.
         */
        bool ParseXMLData() override;

        /**
         * @brief Get the std::map with the relationship items, ordered by ID.
         * @return A reference to the std::map with relationship items.
         * @todo Is there a more elegant way? Using an ordinary overload doesn't work.
         */
        XLRelationshipMap *relationshipsMutable();

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:

        XLRelationshipMap m_relationships; /**< A std::map with the relationship items, ordered by ID */
        unsigned long m_relationshipCount; /**< The number of relationship items in the XLRelationship object */
    };
}

#endif //OPENXLSX_IMPL_XLRELATIONSHIPS_H
