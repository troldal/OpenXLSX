//
// Created by Troldal on 15/08/16.
//

#ifndef OPENXL_XLDOCAPPPROPERTIES_H
#define OPENXL_XLDOCAPPPROPERTIES_H

#include "XLAbstractXMLFile.h"
#include "XML/XMLNode.h"
#include "XML/XMLAttribute.h"
#include "XLSpreadsheetElement.h"

#include <string>
#include <vector>
#include <map>

namespace RapidXLSX
{

//======================================================================================================================
//========== XLWorksheet Class =========================================================================================
//======================================================================================================================

    /**
     * @brief This class is a specialization of the XLAbstractXMLFile, with the purpose of the representing the
     * document app properties in the app.xml file (docProps folder) in the .xlsx package.
     */
    class XLDocAppProperties: public XLAbstractXMLFile,
                              public XLSpreadsheetElement
    {
        friend class XLDocument;

//----------------------------------------------------------------------------------------------------------------------
//           Public Member Functions
//----------------------------------------------------------------------------------------------------------------------

    public:

        /**
         * @brief
         * @param parent
         * @param filePath
         */
        explicit XLDocAppProperties(XLDocument &parent,
                                    const std::string &filePath);

        /**
         * @brief
         * @param other
         */
        XLDocAppProperties(const RapidXLSX::XLDocAppProperties &other) = default;

        /**
         * @brief
         */
        virtual ~XLDocAppProperties();

        /**
         * @brief
         * @param other
         * @return
         */
        XLDocAppProperties &operator=(const XLDocAppProperties &other) = default;

        /**
         * @brief
         * @param title
         * @return
         */
        XMLNode &AddSheetName(const std::string &title);

        /**
         * @brief
         * @param title
         */
        void DeleteSheetName(const std::string &title);

        /**
         * @brief
         * @param oldTitle
         * @param newTitle
         */
        void SetSheetName(const std::string &oldTitle,
                          const std::string &newTitle);

        /**
         * @brief
         * @param title
         * @return
         */
        XMLNode &SheetNameNode(const std::string &title);

        /**
         * @brief
         * @param name
         * @param value
         */
        void AddHeadingPair(const std::string &name,
                            int value);

        /**
         * @brief
         * @param name
         */
        void DeleteHeadingPair(const std::string &name);

        /**
         * @brief
         * @param name
         * @param newValue
         */
        void SetHeadingPair(const std::string &name,
                            int newValue);

        /**
         * @brief
         * @param name
         * @param value
         * @return
         */
        bool SetProperty(const std::string &name,
                         const std::string &value);

        /**
         * @brief
         * @param name
         * @param value
         * @return
         */
        bool SetProperty(const std::string &name,
                         int value);

        /**
         * @brief
         * @param name
         * @param value
         * @return
         */
        bool SetProperty(const std::string &name,
                         double value);

        /**
         * @brief
         * @param name
         * @return
         */
        XMLNode &Property(const std::string &name) const;

        /**
         * @brief
         * @param name
         */
        void DeleteProperty(const std::string &name);

        /**
         * @brief
         * @param sheetName
         * @return
         */
        XMLNode *AppendWorksheetName(const std::string &sheetName);

        /**
         * @brief
         * @param sheetName
         * @return
         */
        XMLNode *PrependWorksheetName(const std::string &sheetName);

        /**
         * @brief
         * @param sheetName
         * @param index
         * @return
         */
        XMLNode *InsertWorksheetName(const std::string &sheetName,
                                     unsigned int index);

        /**
         * @brief
         * @param sheetName
         * @param index
         * @return
         */
        XMLNode *MoveWorksheetName(const std::string &sheetName,
                                   unsigned int index);

        /**
         * @brief
         * @param sheetName
         * @return
         */
        XMLNode *AppendChartsheetName(const std::string &sheetName);

        /**
         * @brief
         * @param sheetName
         * @return
         */
        XMLNode *PrependChartsheetName(const std::string &sheetName);

        /**
         * @brief
         * @param sheetName
         * @param index
         * @return
         */
        XMLNode *InsertChartsheetName(const std::string &sheetName,
                                      unsigned int index);

        /**
         * @brief
         * @param sheetName
         * @param index
         * @return
         */
        XMLNode *MoveChartsheetName(const std::string &sheetName,
                                    unsigned int index);

        /**
         * @brief
         * @return
         */
        virtual bool ParseXMLData() override;

        /**
         * @brief
         * @param sheetName
         * @return
         */
        XMLNode &AppendSheetName(const std::string &sheetName);

        /**
         * @brief
         * @param sheetName
         * @return
         */
        XMLNode &PrependSheetName(const std::string &sheetName);

        /**
         * @brief
         * @param sheetName
         * @param index
         * @return
         */
        XMLNode &InsertSheetName(const std::string &sheetName,
                                 unsigned int index);

        /**
         * @brief
         * @param sheetName
         * @param index
         * @return
         */
        XMLNode *MoveSheetName(const std::string &sheetName,
                               unsigned int index);

//----------------------------------------------------------------------------------------------------------------------
//           Private Member Functions
//----------------------------------------------------------------------------------------------------------------------

    private:

        XMLAttribute *m_sheetCountAttribute; /**< */
        XMLNode *m_sheetNamesParent; /**< */
        std::map<std::string, XMLNode *> m_sheetNameNodes; /**< */

        XMLAttribute *m_headingPairsSize; /**< */
        XMLNode *m_headingPairsCategoryParent; /**< */
        XMLNode *m_headingPairsCountParent; /**< */
        std::vector<std::pair<XMLNode *, XMLNode *>> m_headingPairs; /**< */

        std::map<std::string, XMLNode *> m_properties; /**< */

        unsigned int m_worksheetCount; /**< */
        unsigned int m_chartsheetCount; /**< */
    };

}

#endif //OPENXL_XLDOCAPPPROPERTIES_H
