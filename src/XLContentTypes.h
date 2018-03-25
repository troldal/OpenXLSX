//
// Created by Troldal on 13/08/16.
//

#ifndef OPENXL_XLCONTENTTYPES_H
#define OPENXL_XLCONTENTTYPES_H

#include "XLAbstractXMLFile.h"
#include "XLSpreadsheetElement.h"

#include <map>
#include <string>

namespace RapidXLSX
{
    class XLContentItem;

    using XLContentItemMap = std::map<std::string, std::unique_ptr<XLContentItem>>;

//======================================================================================================================
//========== XLColumnVector Enum =======================================================================================
//======================================================================================================================

    /**
     * @brief
     */
    enum class XLContentType
    {
        Workbook,
        WorkbookMacroEnabled,
        Worksheet,
        Chartsheet,
        ExternalLink,
        Theme,
        Styles,
        SharedStrings,
        Drawing,
        Chart,
        ChartStyle,
        ChartColorStyle,
        ControlProperties,
        CalculationChain,
        VBAProject,
        CoreProperties,
        ExtendedProperties,
        CustomProperties,
        Comments,
        Table,
        VMLDrawing,
        Unknown
    };

//======================================================================================================================
//========== XLContentItem Class =======================================================================================
//======================================================================================================================

    /**
     * @brief
     */
    class XLContentItem
    {
        friend class XLContentTypes;

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief
         * @param other
         * @return
         */
        explicit XLContentItem(const XLContentItem &other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLContentItem(XLContentItem &&other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLContentItem &operator=(const XLContentItem &other);

        /**
         * @brief
         * @param other
         * @return
         */
        XLContentItem &operator=(XLContentItem &&other);

        /**
         * @brief
         * @return
         */
        XLContentType Type() const;

        /**
         * @brief
         * @return
         */
        const std::string &Path() const;

        /**
         * @brief
         */
        void DeleteItem();

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Functions
//----------------------------------------------------------------------------------------------------------------------

    private:

        /**
         * @brief
         * @param node
         * @param path
         * @param type
         * @return
         */
        explicit XLContentItem(XMLNode &node,
                               const std::string &path,
                               XLContentType type);

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:

        XMLNode *m_contentNode; /**< */
        std::string m_contentPath; /**< */
        XLContentType m_contentType; /**< */
    };


//======================================================================================================================
//========== XLContentTypes Class ==============================================================================
//======================================================================================================================

/**
 * @brief The purpose of this class is to load, store add and save item in the [Content_Types].xml file.
 */
    class XLContentTypes: public XLAbstractXMLFile,
                          public XLSpreadsheetElement
    {

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         *
         * @param parent
         * @param filePath
         */
        explicit XLContentTypes(XLDocument &parent,
                                const std::string &filePath);

        /**
         * @brief Destructor
         */
        virtual ~XLContentTypes() = default;

        /**
         * @brief Add a new default key/value pair to the data store.
         * @param key The key
         * @param node The value
         */
        void AddDefault(const std::string &key,
                        XMLNode &node);

        /**
         * @brief Add a new override key/value pair to the data store.
         * @param path The key
         * @param type The value
         */
        void addOverride(const std::string &path,
                         XLContentType type);

        /**
         * @brief Clear the overrides register.
         */
        void ClearOverrides();

        /**
         * @brief
         * @return
         */
        const XLContentItemMap &contentItems() const;

        /**
         * @brief
         * @param path
         * @return
         */
        XLContentItem &ContentItem(const std::string &path);

//----------------------------------------------------------------------------------------------------------------------
//           Protected Member Functions
//----------------------------------------------------------------------------------------------------------------------

    protected:

        /**
         * @brief
         * @return
         */
        bool ParseXMLData() override;

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Variables
//----------------------------------------------------------------------------------------------------------------------

    private:
        std::map<std::string, XMLNode *> m_defaults; /**< */
        XLContentItemMap m_overrides; /**< */
    };
}

#endif //OPENXL_XLCONTENTTYPES_H
