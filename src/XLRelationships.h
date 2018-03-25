//
// Created by Troldal on 25/07/16.
//

#ifndef OPENXL_XLRELATIONSHIPS_H
#define OPENXL_XLRELATIONSHIPS_H

#include <string>
#include <vector>
#include <map>

#include "XML/XMLNode.h"
#include "XLAbstractXMLFile.h"
#include "XLSpreadsheetElement.h"

namespace RapidXLSX
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
        explicit XLRelationshipItem(XMLNode &node,
                                    XLRelationshipType type,
                                    const std::string &target,
                                    const std::string &id);

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:

        XMLNode *m_relationshipNode; /**< A pointer to the XML node with the relationship item */
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
        const XLRelationshipItem &RelationshipByID(const std::string &id) const;

        /**
         * @brief Look up a relationship item by ID.
         * @param id The ID string of the relationship item to retrieve.
         * @return A pointer to the XLRelationshipItem object.
         */
        XLRelationshipItem &RelationshipByID(const std::string &id);

        /**
         * @brief
         * @param target
         * @return
         */
        const XLRelationshipItem &RelationshipByTarget(const std::string &target) const;

        /**
         * @brief
         * @param target
         * @return
         */
        XLRelationshipItem &RelationshipByTarget(const std::string &target);

        /**
         * @brief Get the std::map with the relationship items, ordered by ID.
         * @return A const reference to the std::map with relationship items.
         */
        const XLRelationshipMap &Relationships() const;

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
        XLRelationshipItem &AddRelationship(XLRelationshipType type,
                                            const std::string &target);

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
        XLRelationshipMap &relationshipsMutable();

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:

        XLRelationshipMap m_relationships; /**< A std::map with the relationship items, ordered by ID */
        int m_relationshipCount; /**< The number of relationship items in the XLRelationship object */
    };
}

#endif //OPENXL_XLRELATIONSHIPS_H
